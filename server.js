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

function buildDocRequests(concept, batchNum) {
  const requests = [];
  let cursor = 1;
  const isVideo = /video/i.test(concept.format || '');
  const isReels = /reels/i.test(concept.platform || '');

  function ins(text) {
    requests.push({ insertText: { location: { index: cursor }, text } });
    cursor += text.length;
    return cursor - text.length; // return start index
  }

  function style(start, end, textStyle, fields) {
    requests.push({
      updateTextStyle: {
        range: { startIndex: start, endIndex: end },
        textStyle,
        fields
      }
    });
  }

  function paraStyle(start, end, namedStyle) {
    requests.push({
      updateParagraphStyle: {
        range: { startIndex: start, endIndex: end },
        paragraphStyle: { namedStyleType: namedStyle },
        fields: 'namedStyleType'
      }
    });
  }

  function h1(text) {
    const start = cursor;
    ins(text + '\n');
    paraStyle(start, cursor, 'HEADING_1');
  }

  function h2(text) {
    const start = cursor;
    ins(text + '\n');
    paraStyle(start, cursor, 'HEADING_2');
  }

  function h3(text) {
    const start = cursor;
    ins(text + '\n');
    paraStyle(start, cursor, 'HEADING_3');
  }

  function lv(label, value) {
    const start = cursor;
    const labelEnd = cursor + label.length + 2; // "Label: "
    ins(label + ': ' + (value || '') + '\n');
    style(start, labelEnd, { bold: true }, 'bold');
  }

  function body(text) {
    if (!text) return;
    ins((text || '') + '\n');
  }

  function blank() { ins('\n'); }

  function divider() {
    const start = cursor;
    ins('\n');
    requests.push({
      updateParagraphStyle: {
        range: { startIndex: start, endIndex: cursor },
        paragraphStyle: {
          borderBottom: {
            color: { color: { rgbColor: { red: 0.8, green: 0.8, blue: 0.8 } } },
            dashStyle: 'SOLID',
            padding: { magnitude: 4, unit: 'PT' },
            width: { magnitude: 1, unit: 'PT' }
          }
        },
        fields: 'borderBottom'
      }
    });
  }

  function sectionHeader(text) {
    const start = cursor;
    ins('  ' + text + '  \n');
    style(start, cursor, {
      bold: true,
      foregroundColor: { color: { rgbColor: { red: 1, green: 1, blue: 1 } } },
      fontSize: { magnitude: 13, unit: 'PT' }
    }, 'bold,foregroundColor,fontSize');
    requests.push({
      updateParagraphStyle: {
        range: { startIndex: start, endIndex: cursor },
        paragraphStyle: {
          shading: { backgroundColor: { color: { rgbColor: { red: 0.176, green: 0.416, blue: 0.31 } } } },
          spaceAbove: { magnitude: 12, unit: 'PT' },
          spaceBelow: { magnitude: 6, unit: 'PT' }
        },
        fields: 'shading,spaceAbove,spaceBelow'
      }
    });
  }

  // ── DOC TITLE ────────────────────────────────────────────────────────────────
  h1(`BATCH #${batchNum} — ${concept.name}`);
  const start = cursor - (`BATCH #${batchNum} — ${concept.name}`).length - 1;
  style(start + 1, cursor - 1, {
    fontSize: { magnitude: 18, unit: 'PT' },
    bold: true
  }, 'fontSize,bold');
  blank();

  lv('Brand', 'Kleen Bio');
  lv('Format', mapAdFormat(concept.format));
  lv('Platform', concept.platform || '');
  lv('Awareness', concept.awareness || '');
  lv('Zone', concept.zone || '');
  blank();

  // ── SECTION 1: BATCH BRIEF ───────────────────────────────────────────────────
  sectionHeader('01 — BATCH BRIEF');
  blank();

  h2('STRATEGY');
  lv('Avatar', concept.avatar || '');
  lv('Mass Desire', concept.desire || '');
  lv('Awareness Level', concept.awareness || '');
  blank();

  h2('CREATIVE STRATEGY');
  lv('Angle(s)', concept.angles || '');
  lv('Concept', concept.concept || '');
  lv('Emotional Zone', concept.zone || '');
  lv('Testing Method', concept.testing || '');
  lv('Ad Type', mapAdType(concept.adType));
  blank();

  h2('BREAKTHROUGH MEMO');
  h3('Why we\'re making it');
  body(concept.testing || '');
  blank();
  h3('What it\'s going to say');
  body(concept.concept || '');
  blank();
  h3('How it\'s going to execute');
  body(`${mapAdFormat(concept.format)} on ${concept.platform || ''}`);
  blank();

  h2('COPYWRITER\'S NOTE');
  body(concept.copywritersNote || '');
  blank();

  // ── SECTION 2: CREATION INSTRUCTIONS ─────────────────────────────────────────
  sectionHeader(`02 — CREATION INSTRUCTIONS — ${isVideo ? 'VIDEO EDITOR' : 'GRAPHIC DESIGNER'}`);
  blank();

  if (isVideo) {
    h2('HOOKS — 3 variations, test simultaneously');
    blank();

    for (const hook of (concept.hooks || [])) {
      h3(hook.label || 'Hook');
      lv('Voiceover', hook.vo || '');
      blank();
      lv('Frame 1 Visual', hook.visual || '');
      blank();
      lv('Rationale', hook.rationale || '');
      blank();
    }

    h2('MAIN BODY');
    body(concept.mainBody || '');
    blank();

  } else {
    h2('VARIATIONS');
    blank();

    for (const v of (concept.variations || [])) {
      h3(`VARIATION ${v.num}${v.label ? ' — ' + v.label : ''}`);
      lv('Headline', v.headline || '');
      lv('Subheadline', v.sub || '');
      if (v.body) lv('Body Text', v.body);
      blank();
      h3('Visual Direction');
      body(v.visual || '');
      blank();
      lv('Why this works', v.why || '');
      blank();
    }
  }

  // ── SECTION 3: FACEBOOK POSTING ───────────────────────────────────────────────
  sectionHeader('03 — FACEBOOK POSTING — MEDIA BUYER');
  blank();

  h2('POSTING DETAILS');
  lv('Frame Link', '');
  lv('Page Name', '');
  lv('Landing Page', '');
  lv('Adset Name', '');
  blank();

  if (isReels) {
    h2('INSTAGRAM CAPTIONS');
    h3('Caption 1 — Long');
    body(concept.caption1 || '');
    blank();
    h3('Caption 2 — Short');
    body(concept.caption2 || '');

  } else {
    h2('PRIMARY TEXT (Body Copy)');
    h3('Variation 1');
    body(concept.copy1 || '');
    blank();
    h3('Variation 2');
    body(concept.copy2 || '');
    blank();

    h2('HEADLINES');
    lv('Headline 1', concept.headline1 || '');
    lv('Headline 2', concept.headline2 || '');
    blank();

    h2('SUBHEADLINES (Link Description)');
    lv('Subheadline 1', concept.sub1 || '');
    lv('Subheadline 2', concept.sub2 || '');
    blank();

    h2('LANDING PAGE');
    lv('Landing Page', '');
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
