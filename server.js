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

// ─── AD FORMAT + TYPE MAPPING ─────────────────────────────────────────────────

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

// ─── WRITE SPREADSHEET ROW ────────────────────────────────────────────────────

async function writeSpreadsheetRow(sheets, spreadsheetId, batchNum, concept, docUrl) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'Creative Roadmap'!A:A"
  });
  const rows = response.data.values || [];
  const nextRow = rows.length + 1;

  const values = [[
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
  ]];

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'Creative Roadmap'!A${nextRow}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values }
  });
}

// ─── WRITE DESIRE TO DESIRES SHEET ───────────────────────────────────────────

async function writeDesire(sheets, spreadsheetId, desire) {
  if (!desire) return;
  // Read existing desires
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "'Desires'!A:A"
  });
  const rows = response.data.values || [];
  // Find first empty row
  const nextRow = rows.length + 1;

  // Check desire doesn't already exist
  const existing = rows.map(r => (r[0] || '').toString().trim().toLowerCase());
  if (existing.includes(desire.trim().toLowerCase())) return; // skip duplicate

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'Desires'!A${nextRow}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[desire]] }
  });
}

// ─── CREATE GOOGLE DOC WITH TABS ──────────────────────────────────────────────

async function createGoogleDoc(docs, drive, folderId, title, concept, batchNum) {
  // Create empty doc
  const doc = await docs.documents.create({ requestBody: { title } });
  const docId = doc.data.documentId;

  // Move to Batches folder
  await drive.files.update({
    fileId: docId,
    addParents: folderId,
    removeParents: 'root',
    fields: 'id, parents'
  });

  // Create the 3 tabs and populate them
  await buildDocWithTabs(docs, docId, concept, batchNum);

  return `https://docs.google.com/document/d/${docId}/edit`;
}

// ─── BUILD DOC WITH TABS ──────────────────────────────────────────────────────

async function buildDocWithTabs(docs, docId, concept, batchNum) {
  // Step 1: Create tabs 2 and 3 (tab 1 exists by default)
  // Get current doc to find default tab ID
  const docData = await docs.documents.get({ documentId: docId });
  const defaultTabId = docData.data.tabs?.[0]?.tabProperties?.tabId || null;

  // Create tab 2: Creation Instructions
  const tab2Res = await docs.documents.batchUpdate({
    documentId: docId,
    requestBody: {
      requests: [{
        createTab: {
          insertionIndex: 1,
          tabProperties: { title: 'Creation Instructions' }
        }
      }]
    }
  });

  // Create tab 3: Facebook Posting
  const tab3Res = await docs.documents.batchUpdate({
    documentId: docId,
    requestBody: {
      requests: [{
        createTab: {
          insertionIndex: 2,
          tabProperties: { title: 'Facebook Posting' }
        }
      }]
    }
  });

  // Re-fetch doc to get all tab IDs
  const docData2 = await docs.documents.get({ documentId: docId });
  const tabs = docData2.data.tabs || [];

  // Find tab IDs by title (tabs are in order)
  let briefTabId = null, creationTabId = null, postingTabId = null;
  for (const tab of tabs) {
    const title = tab.tabProperties?.title || '';
    const id = tab.tabProperties?.tabId;
    if (title === 'Tab 1' || title === 'Untitled' || tabs.indexOf(tab) === 0) briefTabId = id;
    if (title === 'Creation Instructions') creationTabId = id;
    if (title === 'Facebook Posting') postingTabId = id;
  }

  // Rename default tab to "Batch Brief"
  if (briefTabId) {
    await docs.documents.batchUpdate({
      documentId: docId,
      requestBody: {
        requests: [{
          updateTabProperties: {
            tabProperties: { tabId: briefTabId, title: 'Batch Brief' },
            fields: 'title'
          }
        }]
      }
    });
  }

  // Populate all three tabs
  const allRequests = [
    ...buildBriefTab(concept, batchNum, briefTabId),
    ...buildCreationTab(concept, batchNum, creationTabId),
    ...buildPostingTab(concept, batchNum, postingTabId)
  ];

  if (allRequests.length > 0) {
    // Split into batches of 50 to avoid API limits
    for (let i = 0; i < allRequests.length; i += 50) {
      await docs.documents.batchUpdate({
        documentId: docId,
        requestBody: { requests: allRequests.slice(i, i + 50) }
      });
    }
  }
}

// ─── TAB CONTENT BUILDERS ─────────────────────────────────────────────────────

function tabInsert(tabId, index, text) {
  return { insertText: { location: { tabId, index }, text } };
}

function tabStyle(tabId, start, end, style, fields) {
  return { updateTextStyle: { range: { tabId, startIndex: start, endIndex: end }, textStyle: style, fields } };
}

function tabParaStyle(tabId, start, end, namedStyle) {
  return {
    updateParagraphStyle: {
      range: { tabId, startIndex: start, endIndex: end },
      paragraphStyle: { namedStyleType: namedStyle },
      fields: 'namedStyleType'
    }
  };
}

function buildTabContent(tabId, sections) {
  // sections = array of { type: 'heading1'|'heading2'|'heading3'|'label-value'|'body'|'blank', text, value }
  const requests = [];
  let cursor = 1;

  for (const s of sections) {
    if (s.type === 'heading1') {
      const text = s.text + '\n';
      requests.push(tabInsert(tabId, cursor, text));
      requests.push(tabParaStyle(tabId, cursor, cursor + text.length, 'HEADING_1'));
      requests.push(tabStyle(tabId, cursor, cursor + text.length - 1, { bold: true, fontSize: { magnitude: 14, unit: 'PT' } }, 'bold,fontSize'));
      cursor += text.length;
    } else if (s.type === 'heading2') {
      const text = s.text + '\n';
      requests.push(tabInsert(tabId, cursor, text));
      requests.push(tabParaStyle(tabId, cursor, cursor + text.length, 'HEADING_2'));
      cursor += text.length;
    } else if (s.type === 'heading3') {
      const text = s.text + '\n';
      requests.push(tabInsert(tabId, cursor, text));
      requests.push(tabParaStyle(tabId, cursor, cursor + text.length, 'HEADING_3'));
      cursor += text.length;
    } else if (s.type === 'label-value') {
      const label = s.text + ': ';
      const value = (s.value || '') + '\n';
      const full = label + value;
      requests.push(tabInsert(tabId, cursor, full));
      requests.push(tabStyle(tabId, cursor, cursor + label.length, { bold: true }, 'bold'));
      cursor += full.length;
    } else if (s.type === 'body') {
      const text = (s.text || '') + '\n';
      requests.push(tabInsert(tabId, cursor, text));
      cursor += text.length;
    } else if (s.type === 'blank') {
      requests.push(tabInsert(tabId, cursor, '\n'));
      cursor += 1;
    }
  }

  return requests;
}

function buildBriefTab(concept, batchNum, tabId) {
  if (!tabId) return [];
  const isVideo = /video/i.test(concept.format || '');
  const sections = [
    { type: 'heading1', text: `BATCH #${batchNum} — ${concept.name}` },
    { type: 'blank' },

    { type: 'heading2', text: 'PART 1: STRATEGY' },
    { type: 'label-value', text: 'Brand', value: 'Kleen Bio' },
    { type: 'label-value', text: 'Batch', value: `#${batchNum}` },
    { type: 'label-value', text: 'Date', value: new Date().toLocaleDateString('en-GB') },
    { type: 'blank' },

    { type: 'heading3', text: 'TARGET' },
    { type: 'label-value', text: 'Avatar', value: concept.avatar || '' },
    { type: 'label-value', text: 'Mass Desire', value: concept.desire || '' },
    { type: 'label-value', text: 'Awareness Level', value: concept.awareness || '' },
    { type: 'blank' },

    { type: 'heading3', text: 'CREATIVE STRATEGY' },
    { type: 'label-value', text: 'Angle(s)', value: concept.angles || '' },
    { type: 'label-value', text: 'Concept', value: concept.concept || '' },
    { type: 'label-value', text: 'Emotional Zone', value: concept.zone || '' },
    { type: 'label-value', text: 'Testing Method', value: concept.testing || '' },
    { type: 'label-value', text: 'New or Variation', value: mapAdType(concept.adType) },
    { type: 'blank' },

    { type: 'heading3', text: 'BREAKTHROUGH MEMO' },
    { type: 'label-value', text: 'Why we\'re making it', value: concept.testing || '' },
    { type: 'label-value', text: 'What it\'s going to say', value: concept.concept || '' },
    { type: 'label-value', text: 'How it\'s going to execute', value: `${concept.format} — ${concept.platform}` },
    { type: 'blank' },

    { type: 'heading2', text: 'CLASSIFICATION' },
    { type: 'label-value', text: 'Format', value: mapAdFormat(concept.format) },
    { type: 'label-value', text: 'Ad Type', value: mapAdType(concept.adType) },
    { type: 'label-value', text: 'Platform', value: concept.platform || '' },
    { type: 'blank' },

    { type: 'heading3', text: 'COPYWRITER\'S NOTE' },
    { type: 'body', text: concept.copywritersNote || '' },
  ];
  return buildTabContent(tabId, sections);
}

function buildCreationTab(concept, batchNum, tabId) {
  if (!tabId) return [];
  const isVideo = /video/i.test(concept.format || '');
  const sections = [
    { type: 'heading1', text: `BATCH #${batchNum} — ${concept.name}` },
    { type: 'label-value', text: 'For', value: isVideo ? 'Video Editor' : 'Graphic Designer' },
    { type: 'label-value', text: 'Platform', value: concept.platform || '' },
    { type: 'label-value', text: 'Format', value: mapAdFormat(concept.format) },
    { type: 'blank' },
  ];

  if (isVideo) {
    sections.push({ type: 'heading2', text: 'HOOKS — 3 variations, test simultaneously' });
    sections.push({ type: 'blank' });

    for (const hook of (concept.hooks || [])) {
      sections.push({ type: 'heading3', text: hook.label || 'Hook' });
      sections.push({ type: 'label-value', text: 'Voiceover', value: hook.vo || '' });
      sections.push({ type: 'label-value', text: 'Frame 1 Visual', value: hook.visual || '' });
      sections.push({ type: 'label-value', text: 'Rationale', value: hook.rationale || '' });
      sections.push({ type: 'blank' });
    }

    sections.push({ type: 'heading2', text: 'MAIN BODY' });
    sections.push({ type: 'body', text: concept.mainBody || '' });

  } else {
    sections.push({ type: 'heading2', text: 'VARIATIONS' });
    sections.push({ type: 'blank' });

    for (const v of (concept.variations || [])) {
      sections.push({ type: 'heading3', text: `VARIATION ${v.num} — ${v.label || ''}` });
      sections.push({ type: 'label-value', text: 'Headline', value: v.headline || '' });
      sections.push({ type: 'label-value', text: 'Subheadline', value: v.sub || '' });
      if (v.body) sections.push({ type: 'label-value', text: 'Body Text', value: v.body });
      sections.push({ type: 'label-value', text: 'Visual Direction', value: v.visual || '' });
      sections.push({ type: 'label-value', text: 'Why this works', value: v.why || '' });
      sections.push({ type: 'blank' });
    }
  }

  return buildTabContent(tabId, sections);
}

function buildPostingTab(concept, batchNum, tabId) {
  if (!tabId) return [];
  const isReels = /reels/i.test(concept.platform || '');
  const sections = [
    { type: 'heading1', text: `BATCH #${batchNum} — ${concept.name}` },
    { type: 'blank' },
    { type: 'heading2', text: 'POSTING DETAILS' },
    { type: 'label-value', text: 'Frame Link', value: '' },
    { type: 'label-value', text: 'Page Name', value: '' },
    { type: 'label-value', text: 'Landing Page', value: '' },
    { type: 'label-value', text: 'Adset Name', value: '' },
    { type: 'blank' },
  ];

  if (isReels) {
    sections.push({ type: 'heading2', text: 'INSTAGRAM CAPTIONS' });
    sections.push({ type: 'heading3', text: 'Caption 1 (Long)' });
    sections.push({ type: 'body', text: concept.caption1 || '' });
    sections.push({ type: 'blank' });
    sections.push({ type: 'heading3', text: 'Caption 2 (Short)' });
    sections.push({ type: 'body', text: concept.caption2 || '' });
  } else {
    sections.push({ type: 'heading2', text: 'PRIMARY TEXT (Body Copy)' });
    sections.push({ type: 'heading3', text: 'Variation 1' });
    sections.push({ type: 'body', text: concept.copy1 || '' });
    sections.push({ type: 'blank' });
    sections.push({ type: 'heading3', text: 'Variation 2' });
    sections.push({ type: 'body', text: concept.copy2 || '' });
    sections.push({ type: 'blank' });

    sections.push({ type: 'heading2', text: 'HEADLINES' });
    sections.push({ type: 'label-value', text: 'Headline 1', value: concept.headline1 || '' });
    sections.push({ type: 'label-value', text: 'Headline 2', value: concept.headline2 || '' });
    sections.push({ type: 'blank' });

    sections.push({ type: 'heading2', text: 'SUBHEADLINES (Link Description)' });
    sections.push({ type: 'label-value', text: 'Subheadline 1', value: concept.sub1 || '' });
    sections.push({ type: 'label-value', text: 'Subheadline 2', value: concept.sub2 || '' });
    sections.push({ type: 'blank' });

    sections.push({ type: 'heading2', text: 'LANDING PAGE' });
    sections.push({ type: 'label-value', text: 'Landing Page', value: '' });
  }

  return buildTabContent(tabId, sections);
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

      // Create Google Doc with tabs
      const docUrl = await createGoogleDoc(docs, drive, folderId, title, concept, batchNum);

      // Write row to Creative Roadmap
      await writeSpreadsheetRow(sheets, spreadsheetId, batchNum, concept, docUrl);

      // Write desire to Desires sheet (skip duplicates)
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
