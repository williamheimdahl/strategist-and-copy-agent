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

// Step 1: Get auth URL
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

// Step 2: Handle callback, return tokens
app.get('/auth/callback', async (req, res) => {
  const { code } = req.query;
  const oauth2Client = getOAuthClient();
  try {
    const { tokens } = await oauth2Client.getToken(code);
    // Return tokens to frontend via postMessage
    res.send(`
      <html><body><script>
        window.opener.postMessage(${JSON.stringify({ tokens })}, '*');
        window.close();
      </script></body></html>
    `);
  } catch (err) {
    res.send(`<html><body><script>
      window.opener.postMessage({ error: '${err.message}' }, '*');
      window.close();
    </script></body></html>`);
  }
});

// ─── HEALTH ───────────────────────────────────────────────────────────────────

app.get('/health', (req, res) => res.json({ status: 'online' }));

// ─── CLAUDE API PROXY ─────────────────────────────────────────────────────────

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

// ─── FIND OR CREATE FOLDER ────────────────────────────────────────────────────

async function findOrCreateFolder(drive, folderName) {
  // Search for existing folder
  const search = await drive.files.list({
    q: `mimeType='application/vnd.google-apps.folder' and name='${folderName}' and trashed=false`,
    fields: 'files(id, name)'
  });
  if (search.data.files.length > 0) return search.data.files[0].id;

  // Create folder
  const folder = await drive.files.create({
    requestBody: { name: folderName, mimeType: 'application/vnd.google-apps.folder' },
    fields: 'id'
  });
  return folder.data.id;
}

// ─── CREATE GOOGLE DOC ────────────────────────────────────────────────────────

async function createGoogleDoc(docs, drive, folderId, title, concept) {
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

  // Build content requests
  const requests = buildDocRequests(concept);
  if (requests.length > 0) {
    await docs.documents.batchUpdate({ documentId: docId, requestBody: { requests } });
  }

  return `https://docs.google.com/document/d/${docId}/edit`;
}

// ─── BUILD DOC CONTENT ────────────────────────────────────────────────────────

function buildDocRequests(concept) {
  const requests = [];
  let cursor = 1;

  function insertText(text, bold = false, fontSize = 11, rgb = null) {
    const req = {
      insertText: { location: { index: cursor }, text }
    };
    requests.push(req);

    const styleReq = {
      updateTextStyle: {
        range: { startIndex: cursor, endIndex: cursor + text.length },
        textStyle: {
          bold,
          fontSize: { magnitude: fontSize, unit: 'PT' },
          ...(rgb ? { foregroundColor: { color: { rgbColor: rgb } } } : {})
        },
        fields: 'bold,fontSize' + (rgb ? ',foregroundColor' : '')
      }
    };
    requests.push(styleReq);
    cursor += text.length;
  }

  function insertHeading(text, level) {
    const namedStyle = level === 1 ? 'HEADING_1' : level === 2 ? 'HEADING_2' : 'HEADING_3';
    requests.push({ insertText: { location: { index: cursor }, text: text + '\n' } });
    requests.push({
      updateParagraphStyle: {
        range: { startIndex: cursor, endIndex: cursor + text.length + 1 },
        paragraphStyle: { namedStyleType: namedStyle },
        fields: 'namedStyleType'
      }
    });
    cursor += text.length + 1;
  }

  function insertParagraph(text) {
    const content = text + '\n';
    requests.push({ insertText: { location: { index: cursor }, text: content } });
    cursor += content.length;
  }

  function insertLabelValue(labelText, valueText) {
    const line = labelText + ': ' + valueText + '\n';
    requests.push({ insertText: { location: { index: cursor }, text: line } });
    // Bold the label part
    requests.push({
      updateTextStyle: {
        range: { startIndex: cursor, endIndex: cursor + labelText.length + 2 },
        textStyle: { bold: true },
        fields: 'bold'
      }
    });
    cursor += line.length;
  }

  function insertPageBreak() {
    requests.push({ insertPageBreak: { location: { index: cursor } } });
    cursor += 1;
  }

  function insertDivider() {
    insertParagraph('─────────────────────────────────────────');
  }

  // ── PAGE 1: BATCH BRIEF ────────────────────────────────────────────────────
  insertHeading(`BATCH #${concept.batchNum} — ${concept.name}`, 1);
  insertParagraph(`${concept.format}  ·  ${concept.platform}  ·  ${concept.awareness}  ·  ${concept.zone}`);
  insertParagraph('');

  insertHeading('BATCH BRIEF', 2);
  insertLabelValue('Brand', 'Kleen Bio');
  insertLabelValue('Batch', `#${concept.batchNum}`);
  insertLabelValue('Format', concept.format);
  insertLabelValue('Platform', concept.platform);
  insertLabelValue('Awareness Level', concept.awareness);
  insertLabelValue('Emotional Zone', concept.zone);
  insertLabelValue('Testing Method', concept.testing);
  insertDivider();
  insertLabelValue('Concept', concept.concept);
  insertLabelValue('Desire', concept.desire);
  insertLabelValue('Angle(s)', concept.angles);
  insertLabelValue('Ad Format', concept.adFormat);
  insertLabelValue('Target Avatar', concept.avatar);
  insertDivider();
  insertHeading('COPYWRITER\'S NOTE', 3);
  insertParagraph(concept.copywritersNote);

  insertPageBreak();

  // ── PAGE 2: CREATION INSTRUCTIONS ─────────────────────────────────────────
  insertHeading(`BATCH #${concept.batchNum} — ${concept.name}`, 1);
  insertHeading('CREATION INSTRUCTIONS', 2);
  insertLabelValue('For', concept.format === 'Video' ? 'Video Editor' : 'Graphic Designer');
  insertLabelValue('Platform', concept.platform);
  insertDivider();

  if (concept.isVideo) {
    insertHeading('HOOKS (3 variations — test simultaneously)', 2);
    for (const hook of concept.hooks) {
      insertHeading(hook.label, 3);
      insertHeading('VOICEOVER', 3);
      insertParagraph(hook.vo);
      insertHeading('FRAME 1 VISUAL', 3);
      insertParagraph(hook.visual);
      insertHeading('RATIONALE', 3);
      insertParagraph(hook.rationale);
      insertParagraph('');
    }
    insertHeading('MAIN BODY', 2);
    insertParagraph(concept.mainBody);
  } else {
    insertHeading('VARIATIONS', 2);
    for (const v of concept.variations) {
      insertHeading(`VARIATION ${v.num} — ${v.label}`, 3);
      insertLabelValue('Headline', v.headline);
      insertLabelValue('Subheadline', v.sub);
      if (v.body) insertLabelValue('Body Text', v.body);
      insertHeading('VISUAL DIRECTION', 3);
      insertParagraph(v.visual);
      insertHeading('WHY THIS WORKS', 3);
      insertParagraph(v.why);
      insertParagraph('');
    }
  }

  insertPageBreak();

  // ── PAGE 3: FACEBOOK POSTING ───────────────────────────────────────────────
  insertHeading(`BATCH #${concept.batchNum} — ${concept.name}`, 1);
  insertHeading('FACEBOOK POSTING INSTRUCTIONS', 2);

  insertHeading('POSTING DETAILS', 3);
  insertLabelValue('Frame Link', '_______________________________________________');
  insertLabelValue('Page Name', '_______________________________________________');
  insertLabelValue('Landing Page', '_______________________________________________');
  insertLabelValue('Adset Name', '_______________________________________________');
  insertDivider();

  if (concept.isReels) {
    insertHeading('INSTAGRAM CAPTION', 2);
    insertHeading('CAPTION 1 (Long)', 3);
    insertParagraph(concept.caption1);
    insertHeading('CAPTION 2 (Short)', 3);
    insertParagraph(concept.caption2);
  } else {
    insertHeading('PRIMARY TEXT (Body Copy)', 2);
    insertHeading('VARIATION 1', 3);
    insertParagraph(concept.copy1);
    insertHeading('VARIATION 2', 3);
    insertParagraph(concept.copy2);
    insertDivider();
    insertHeading('HEADLINES', 2);
    insertHeading('HEADLINE 1', 3);
    insertParagraph(concept.headline1);
    insertHeading('HEADLINE 2', 3);
    insertParagraph(concept.headline2);
    insertDivider();
    insertHeading('SUBHEADLINES (Link Description)', 2);
    insertHeading('SUBHEADLINE 1', 3);
    insertParagraph(concept.sub1);
    insertHeading('SUBHEADLINE 2', 3);
    insertParagraph(concept.sub2);
  }

  return requests;
}

// ─── WRITE SPREADSHEET ROW ────────────────────────────────────────────────────

async function writeSpreadsheetRow(sheets, spreadsheetId, concept, docUrl) {
  // Find next empty row in the sheet
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: 'A:A'
  });
  const rows = response.data.values || [];
  const nextRow = rows.length + 1;

  const adTypeMap = { ideation: 'Ideation', iteration: 'Iteration', imitation: 'Imitation' };
  const adType = adTypeMap[concept.adType?.toLowerCase()] || 'Ideation';

  const values = [[
    `BATCH #${concept.batchNum}`,  // A
    'Working',                      // B
    'Williams AI',                  // C
    concept.name,                   // D — Ad Concept
    concept.desire,                 // E — Desire
    concept.angles,                 // F — Angle(s)
    concept.testing,                // G — What are you testing
    concept.awareness,              // H — Awareness Level
    concept.format,                 // I — Ad Format
    adType,                         // J — Ad Type
    '',                             // K — Video/Graphic Editor (leave empty)
    docUrl                          // L — Link To Brief
  ]];

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `A${nextRow}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values }
  });
}

// ─── MAIN GENERATE ENDPOINT ───────────────────────────────────────────────────

app.post('/api/generate', async (req, res) => {
  const { tokens, concepts, spreadsheetId } = req.body;

  if (!tokens || !concepts || !spreadsheetId) {
    return res.status(400).json({ error: 'Missing tokens, concepts, or spreadsheetId' });
  }

  try {
    const oauth2Client = getOAuthClient();
    oauth2Client.setCredentials(tokens);

    const docs = google.docs({ version: 'v1', auth: oauth2Client });
    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
    const drive = google.drive({ version: 'v3', auth: oauth2Client });

    // Find or create Batches folder
    const folderId = await findOrCreateFolder(drive, 'Batches');

    const results = [];

    for (const concept of concepts) {
      const title = `Batch #${concept.batchNum} — ${concept.name} — Kleen Bio`;

      // Create Google Doc
      const docUrl = await createGoogleDoc(docs, drive, folderId, title, concept);

      // Write to spreadsheet
      await writeSpreadsheetRow(sheets, spreadsheetId, concept, docUrl);

      results.push({
        batchNum: concept.batchNum,
        name: concept.name,
        docUrl
      });
    }

    res.json({ success: true, results });

  } catch (err) {
    console.error('Generate error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Brief Generator backend running on port ${PORT}`);
});
