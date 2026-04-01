const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 8080;

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// ─── GOOGLE OAUTH ─────────────────────────────────────────────────────────────

function getOAuthClient() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
}

app.get('/auth/google', (req, res) => {
  const oauth2Client = getOAuthClient();
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: [
      'https://www.googleapis.com/auth/documents',
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive'
    ],
    prompt: 'consent'
  });
  res.json({ url });
});

app.get('/auth/callback', async (req, res) => {
  const { code } = req.query;
  const oauth2Client = getOAuthClient();
  try {
    const { tokens } = await oauth2Client.getToken(code);
    res.send(`<html><body><script>window.opener.postMessage(${JSON.stringify({ tokens })}, '*');window.close();</script></body></html>`);
  } catch (err) {
    res.send(`<html><body><script>window.opener.postMessage({ error: '${err.message}' }, '*');window.close();</script></body></html>`);
  }
});

// ─── HEALTH + CLAUDE PROXY ────────────────────────────────────────────────────

app.get('/health', (req, res) => res.json({ status: 'online' }));

app.post('/api/claude', async (req, res) => {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'API key not configured on server.' });
  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: 'claude-sonnet-4-5', max_tokens: 8000, system: req.body.system, messages: req.body.messages })
    });
    const data = await response.json();
    if (!response.ok) return res.status(response.status).json({ error: data.error?.message || 'Anthropic API error' });
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── HELPERS ──────────────────────────────────────────────────────────────────

async function findOrCreateFolder(drive, folderName) {
  const search = await drive.files.list({
    q: `mimeType='application/vnd.google-apps.folder' and name='${folderName}' and trashed=false`,
    fields: 'files(id, name)'
  });
  if (search.data.files.length > 0) return search.data.files[0].id;
  const folder = await drive.files.create({
    requestBody: { name: folderName, mimeType: 'application/vnd.google-apps.folder' },
    fields: 'id'
  });
  return folder.data.id;
}

async function getNextBatchNumber(sheets, spreadsheetId) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'Creative Roadmap'!A:A"
  });
  const rows = response.data.values || [];
  let maxNum = 0;
  for (const row of rows) {
    const match = (row[0] || '').toString().match(/BATCH\s*#(\d+)/i);
    if (match) {
      const num = parseInt(match[1]);
      if (num > maxNum) maxNum = num;
    }
  }
  return maxNum + 1;
}

function mapAdFormat(format) {
  const f = (format || '').toLowerCase();
  if (f.includes('static')) return '🖼️ Static';
  if (f.includes('video')) return '🎬 Video';
  if (f.includes('promo')) return '🏷️ Promo';
  if (f.includes('gif')) return '📱 GIF';
  return '🎬 Video';
}

function mapAdType(adType) {
  const t = (adType || '').toLowerCase();
  if (t.includes('iteration')) return '🔄 Iteration';
  if (t.includes('imitation')) return '🎭 Imitation';
  return '💡 Ideation';
}

// ─── SPREADSHEET WRITES ───────────────────────────────────────────────────────

async function writeSpreadsheetRow(sheets, spreadsheetId, batchNum, concept, docUrl) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'Creative Roadmap'!A:A"
  });
  const rows = response.data.values || [];
  const nextRow = rows.length + 1;

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'Creative Roadmap'!A${nextRow}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        `BATCH #${batchNum}`,
        'Working',
        'Williams AI',
        concept.name,
        concept.desire,
        concept.angles,
        concept.testing,
        concept.awareness,
        mapAdFormat(concept.format),
        mapAdType(concept.adType),
        '',
        docUrl
      ]]
    }
  });
}

async function writeDesire(sheets, spreadsheetId, desire) {
  if (!desire) return;
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'Desires'!A:A"
  });
  const rows = response.data.values || [];
  const existing = rows.map(r => (r[0] || '').toString().trim().toLowerCase());
  if (existing.includes(desire.trim().toLowerCase())) return;
  const nextRow = rows.length + 1;
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'Desires'!A${nextRow}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[desire]] }
  });
}

// ─── BUILD GOOGLE DOC ─────────────────────────────────────────────────────────

async function createGoogleDoc(docs, drive, folderId, title, concept, batchNum) {
  const doc = await docs.documents.create({ requestBody: { title } });
  const docId = doc.data.documentId;

  await drive.files.update({
    fileId: docId,
    addParents: folderId,
    removeParents: 'root',
    fields: 'id, parents'
  });

  const requests = buildDocRequests(concept, batchNum);

  // Send in batches of 50 to avoid limits
  for (let i = 0; i < requests.length; i += 50) {
    await docs.documents.batchUpdate({
      documentId: docId,
      requestBody: { requests: requests.slice(i, i + 50) }
    });
  }

  return `https://docs.google.com/document/d/${docId}/edit`;
}

// ─── DOC CONTENT BUILDER ──────────────────────────────────────────────────────
//
// Outline strategy:
//   HEADING_1 → doc title only (not in outline nav)
//   HEADING_2 → the 3 section banners ONLY — these are the 3 outline items
//   Everything else → NORMAL_TEXT with bold/size styling
//   This gives the left nav exactly 3 jump links: Brief / Creation / Posting
//
// ──────────────────────────────────────────────────────────────────────────────

function buildDocRequests(concept, batchNum) {
  const requests = [];
  let cursor = 1;
  const isVideo = /video/i.test(concept.format || '');
  const isReels = /reels/i.test(concept.platform || '');

  // ── Low-level primitives ────────────────────────────────────────────────────

  function ins(text) {
    requests.push({ insertText: { location: { index: cursor }, text } });
    cursor += text.length;
    return cursor - text.length; // returns start index
  }

  function style(start, end, textStyle, fields) {
    if (end <= start) return;
    requests.push({
      updateTextStyle: {
        range: { startIndex: start, endIndex: end },
        textStyle,
        fields
      }
    });
  }

  function paraStyle(start, end, paragraphStyle, fields) {
    if (end <= start) return;
    requests.push({
      updateParagraphStyle: {
        range: { startIndex: start, endIndex: end },
        paragraphStyle,
        fields
      }
    });
  }

  function namedStyle(start, end, namedStyleType) {
    requests.push({
      updateParagraphStyle: {
        range: { startIndex: start, endIndex: end },
        paragraphStyle: { namedStyleType },
        fields: 'namedStyleType'
      }
    });
  }

  function blank() { ins('\n'); }

  // ── High-level content helpers ──────────────────────────────────────────────

  // Doc title — large, bold, HEADING_1 (shows in outline but below the 3 section headings visually)
  function docTitle(text) {
    const start = cursor;
    ins(text + '\n');
    namedStyle(start, cursor, 'HEADING_1');
    style(start, cursor - 1, { fontSize: { magnitude: 20, unit: 'PT' }, bold: true }, 'fontSize,bold');
  }

  // The 3 section banners — HEADING_2 so they appear as the outline nav items
  function sectionHeading(text) {
    const start = cursor;
    ins(text + '\n');
    namedStyle(start, cursor, 'HEADING_2');
    // Override the visual style: white text on dark green background
    style(start, cursor - 1, {
      bold: true,
      foregroundColor: { color: { rgbColor: { red: 1, green: 1, blue: 1 } } },
      fontSize: { magnitude: 13, unit: 'PT' }
    }, 'bold,foregroundColor,fontSize');
    paraStyle(start, cursor, {
      shading: { backgroundColor: { color: { rgbColor: { red: 0.106, green: 0.369, blue: 0.271 } } } },
      spaceAbove: { magnitude: 16, unit: 'PT' },
      spaceBelow: { magnitude: 8, unit: 'PT' },
      indentStart: { magnitude: 6, unit: 'PT' },
      indentEnd: { magnitude: 6, unit: 'PT' }
    }, 'shading,spaceAbove,spaceBelow,indentStart,indentEnd');
  }

  // Sub-section label — bold, slightly larger, normal text (NOT a heading — won't clutter outline)
  function subLabel(text) {
    const start = cursor;
    ins(text + '\n');
    style(start, cursor - 1, {
      bold: true,
      fontSize: { magnitude: 11, unit: 'PT' },
      foregroundColor: { color: { rgbColor: { red: 0.106, green: 0.369, blue: 0.271 } } }
    }, 'bold,fontSize,foregroundColor');
    paraStyle(start, cursor, {
      spaceAbove: { magnitude: 10, unit: 'PT' },
      spaceBelow: { magnitude: 2, unit: 'PT' }
    }, 'spaceAbove,spaceBelow');
  }

  // Item label — bold label + black value on same line: "Label: value"
  function lv(label, value) {
    const start = cursor;
    const labelEnd = start + label.length + 2; // "Label: "
    ins(label + ': ' + (value || '') + '\n');
    style(start, labelEnd, { bold: true }, 'bold');
    // Force value portion to black — only if there is actual value text to style
    if (value && value.length > 0) {
      style(labelEnd, cursor - 1, {
        foregroundColor: { color: { rgbColor: { red: 0, green: 0, blue: 0 } } }
      }, 'foregroundColor');
    }
  }

  // Hook label — bold dark label for Hook A/B/C, visual direction, etc.
  function hookLabel(text) {
    const start = cursor;
    ins(text + '\n');
    style(start, cursor - 1, {
      bold: true,
      fontSize: { magnitude: 10.5, unit: 'PT' },
      foregroundColor: { color: { rgbColor: { red: 0.15, green: 0.15, blue: 0.15 } } }
    }, 'bold,fontSize,foregroundColor');
    paraStyle(start, cursor, {
      spaceAbove: { magnitude: 8, unit: 'PT' },
      spaceBelow: { magnitude: 2, unit: 'PT' }
    }, 'spaceAbove,spaceBelow');
  }

  // Variation block header — used in Section 3 for Variation 1 / Variation 2 / Caption 1 / Caption 2
  // Light grey pill background so the media buyer can immediately see where one ends and the next begins
  function variationBlock(text) {
    const start = cursor;
    ins(text + '\n');
    style(start, cursor - 1, {
      bold: true,
      fontSize: { magnitude: 11, unit: 'PT' },
      foregroundColor: { color: { rgbColor: { red: 0.1, green: 0.1, blue: 0.1 } } }
    }, 'bold,fontSize,foregroundColor');
    paraStyle(start, cursor, {
      shading: { backgroundColor: { color: { rgbColor: { red: 0.918, green: 0.918, blue: 0.918 } } } },
      spaceAbove: { magnitude: 14, unit: 'PT' },
      spaceBelow: { magnitude: 6, unit: 'PT' },
      indentStart: { magnitude: 6, unit: 'PT' },
      indentEnd: { magnitude: 6, unit: 'PT' }
    }, 'shading,spaceAbove,spaceBelow,indentStart,indentEnd');
  }

  // Plain body text — explicitly black so it never inherits grey from the Docs theme
  function body(text) {
    if (!text) return;
    const start = cursor;
    ins((text || '') + '\n');
    style(start, cursor - 1, {
      foregroundColor: { color: { rgbColor: { red: 0, green: 0, blue: 0 } } },
      fontSize: { magnitude: 10.5, unit: 'PT' }
    }, 'foregroundColor,fontSize');
  }

  // ── BUILD DOCUMENT ──────────────────────────────────────────────────────────

  // Doc title
  docTitle(`BATCH #${batchNum} — ${concept.name}`);
  blank();
  lv('Brand', concept.brand || 'Kleen Bio');
  lv('Format', mapAdFormat(concept.format));
  lv('Platform', concept.platform || '');
  lv('Awareness', concept.awareness || '');
  lv('Zone', concept.zone || '');
  blank();

  // ── SECTION 1 ───────────────────────────────────────────────────────────────
  sectionHeading('01 — BATCH BRIEF');
  blank();

  subLabel('STRATEGY');
  lv('Avatar', concept.avatar || '');
  lv('Mass Desire', concept.desire || '');
  lv('Awareness Level', concept.awareness || '');
  blank();

  subLabel('CREATIVE STRATEGY');
  lv('Angle(s)', concept.angles || '');
  lv('Concept', concept.concept || '');
  lv('Emotional Zone', concept.zone || '');
  lv('Testing Method', concept.testing || '');
  lv('Ad Type', mapAdType(concept.adType));
  blank();

  subLabel('BREAKTHROUGH MEMO');
  hookLabel('Why we\'re making it');
  body(concept.testing || '');
  blank();
  hookLabel('What it\'s going to say');
  body(concept.concept || '');
  blank();
  hookLabel('How it\'s going to execute');
  body(`${mapAdFormat(concept.format)} on ${concept.platform || ''}`);
  blank();

  subLabel('COPYWRITER\'S NOTE');
  body(concept.copywritersNote || '');
  blank();

  // ── SECTION 2 ───────────────────────────────────────────────────────────────
  sectionHeading(`02 — CREATION INSTRUCTIONS — ${isVideo ? 'VIDEO EDITOR' : 'GRAPHIC DESIGNER'}`);
  blank();

  if (isVideo) {
    subLabel('HOOKS — 3 variations, test simultaneously');
    blank();

    for (const hook of (concept.hooks || [])) {
      hookLabel(hook.label || 'Hook');
      lv('Voiceover', hook.vo || '');
      blank();
      lv('Frame 1 Visual', hook.visual || '');
      blank();
      lv('Rationale', hook.rationale || '');
      blank();
    }

    subLabel('MAIN BODY');
    body(concept.mainBody || '');
    blank();

  } else {
    subLabel('VARIATIONS');
    blank();

    for (const v of (concept.variations || [])) {
      hookLabel(`VARIATION ${v.num}${v.label ? ' — ' + v.label : ''}`);
      lv('Headline', v.headline || '');
      lv('Subheadline', v.sub || '');
      if (v.body) lv('Body Text', v.body);
      blank();
      hookLabel('Visual Direction');
      body(v.visual || '');
      blank();
      lv('Why this works', v.why || '');
      blank();
    }
  }

  // ── SECTION 3 ───────────────────────────────────────────────────────────────
  sectionHeading('03 — FACEBOOK POSTING — MEDIA BUYER');
  blank();

  subLabel('POSTING DETAILS');
  lv('Frame Link', '');
  lv('Page Name', '');
  lv('Landing Page', '');
  lv('Adset Name', '');
  blank();

  if (isReels) {
    subLabel('INSTAGRAM CAPTIONS');
    variationBlock('CAPTION 1 — LONG');
    body(concept.caption1 || '');
    blank();
    variationBlock('CAPTION 2 — SHORT');
    body(concept.caption2 || '');

  } else {
    subLabel('PRIMARY TEXT (Body Copy)');
    variationBlock('VARIATION 1');
    body(concept.copy1 || '');
    blank();
    variationBlock('VARIATION 2');
    body(concept.copy2 || '');
    blank();

    subLabel('HEADLINES');
    lv('Headline 1', concept.headline1 || '');
    lv('Headline 2', concept.headline2 || '');
    blank();

    subLabel('SUBHEADLINES (Link Description)');
    lv('Subheadline 1', concept.sub1 || '');
    lv('Subheadline 2', concept.sub2 || '');
    blank();

  }

  return requests;
}

// ─── MAIN GENERATE ENDPOINT ───────────────────────────────────────────────────

app.post('/api/generate', async (req, res) => {
  const { tokens, concepts, spreadsheetId, brand } = req.body;
  if (!tokens || !concepts || !spreadsheetId) {
    return res.status(400).json({ error: 'Missing tokens, concepts, or spreadsheetId' });
  }

  try {
    const oauth2Client = getOAuthClient();
    oauth2Client.setCredentials(tokens);

    const docs = google.docs({ version: 'v1', auth: oauth2Client });
    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
    const drive = google.drive({ version: 'v3', auth: oauth2Client });

    const folderId = await findOrCreateFolder(drive, 'Batches');
    let nextBatchNum = await getNextBatchNumber(sheets, spreadsheetId);

    const results = [];

    for (const concept of concepts) {
      const batchNum = nextBatchNum++;
      const brandName = brand || 'Kleen Bio';
      const title = `Batch #${batchNum} — ${concept.name} — ${brandName}`;

      const docUrl = await createGoogleDoc(docs, drive, folderId, title, concept, batchNum);
      await writeSpreadsheetRow(sheets, spreadsheetId, batchNum, concept, docUrl);
      await writeDesire(sheets, spreadsheetId, concept.desire);

      results.push({ batchNum, name: concept.name, docUrl });
    }

    res.json({ success: true, results });

  } catch (err) {
    console.error('Generate error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Creative Command + Brief Generator backend running on port ${PORT}`);
});
