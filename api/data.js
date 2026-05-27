/**
 * Vercel Serverless Function: /api/data
 * Reads from the regional "Current month" sheet.
 *
 * Required env vars:
 *   GOOGLE_SA_KEY   — base64-encoded service account JSON key
 *   SPREADSHEET_ID  — Regional sheet ID: 1eXaylD_21vAYRB_Q76PeGL5RlXGhbSlq7XZxzfKkJgk
 *
 * Optional env vars:
 *   CHECKINS_SPREADSHEET_ID — Sheet ID for field check-ins (defaults to SPREADSHEET_ID)
 *   PERSONAL_TARGET — Monthly personal NMV target in LCY (default: 0, meaning show from sheet)
 *
 * Supported query params:
 *   ?action=dashboard&email=...&name=...
 *   ?action=teamlist
 */

const { google } = require('googleapis');

const DATA_TAB     = 'Current month';
const CHECKINS_TAB = 'RCC_Field_Checkins';

// ── Auth ─────────────────────────────────────────────────────────────────────
function getAuth() {
  const key = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(key),
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
  });
}

// ── Handler ───────────────────────────────────────────────────────────────────
module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'GET')    return res.status(405).json({ success: false, error: 'Method not allowed' });

  if (!process.env.GOOGLE_SA_KEY)  return res.status(500).json({ success: false, error: 'GOOGLE_SA_KEY not set' });
  if (!process.env.SPREADSHEET_ID) return res.status(500).json({ success: false, error: 'SPREADSHEET_ID not set' });

  try {
    const auth   = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    const sid    = process.env.SPREADSHEET_ID;
    const action = req.query.action || 'dashboard';

    // ── Load sheet: find headers + all rows ──────────────────────────────────
    const raw = await sheets.spreadsheets.values.get({
      spreadsheetId:    sid,
      range:            `'${DATA_TAB}'!A1:AZ`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const allRows = raw.data.values || [];

    // Find header row — the row that contains both "email" and "name"
    let headerIdx = -1;
    let cols      = {};   // { EMAIL: 4, NAME: 3, ... }

    for (let i = 0; i < Math.min(allRows.length, 6); i++) {
      const row = allRows[i].map(c => String(c || '').toLowerCase().trim());
      if (row.some(c => c.includes('email')) && row.some(c => c.includes('name'))) {
        headerIdx = i;
        // Map every column we care about
        cols = mapColumns(row);
        break;
      }
    }

    if (headerIdx < 0) {
      return res.status(500).json({ success: false, error: 'Could not find header row in "' + DATA_TAB + '" tab. Make sure it has Email and Name columns.' });
    }

    const dataRows = allRows.slice(headerIdx + 1).filter(r => r && r.length > 0 && r[0]);

    // ── Filter to current month ───────────────────────────────────────────────
    const currentMonth = getCurrentMonth();   // e.g. 202605
    const monthRows = dataRows.filter(row => {
      const monthVal = String(row[cols.MONTH] || '').replace(/[^0-9]/g, '');
      return monthVal === String(currentMonth);
    });

    // If no rows for current month, fall back to most recent month available
    const workingRows = monthRows.length > 0
      ? monthRows
      : getMostRecentMonthRows(dataRows, cols.MONTH);

    if (action === 'teamlist') {
      return res.status(200).json(buildTeamList(workingRows, cols));
    }

    // ── Dashboard ─────────────────────────────────────────────────────────────
    const userEmail = (req.query.email || '').toLowerCase().trim();
    const userName  = req.query.name  || getUserDisplayName(userEmail);

    // Find the logged-in user's row
    const userRow = workingRows.find(r =>
      String(r[cols.EMAIL] || '').toLowerCase().trim() === userEmail
    ) || null;

    const kpis         = buildKpis(userRow, workingRows, cols);
    const orderPoints  = buildOrderPointsSummary(userRow, workingRows, cols);
    const agents       = buildAgentsSummary(userRow, workingRows, cols);
    const fieldCheckins = await getFieldCheckinsSummary(sheets, sid, userEmail);

    // Resolved display name — prefer sheet name over signup name
    const resolvedName = (userRow && userRow[cols.NAME])
      ? String(userRow[cols.NAME]).trim()
      : userName;

    return res.status(200).json({
      success: true,
      user: {
        email:    userEmail,
        name:     resolvedName,
        initials: getInitials(resolvedName)
      },
      kpis,
      orderPoints,
      agents,
      fieldCheckins,
      timestamp: new Date().toISOString()
    });

  } catch (err) {
    console.error('api/data error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};

// ── Column mapper ─────────────────────────────────────────────────────────────
// Finds the index of each key column by name (case-insensitive partial match)
function mapColumns(headerRow) {
  const find = (...keywords) => {
    for (const kw of keywords) {
      const idx = headerRow.findIndex(h => h.includes(kw));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  return {
    MONTH:        find('month'),
    NAME:         find('name'),
    EMAIL:        find('email'),
    REGION:       find('region'),
    LOCATION:     find('location'),
    PERSONAL_NMV: find('rcc nmv (lcy)', 'rcc nmv lcy', 'personal nmv'),
    TEAM_NMV:     find('rcc team nmv (lcy)', 'team nmv (lcy)', 'rcc team nmv'),
    NMV_TARGET:   find('nmv target', 'rcc + team nmv target', 'target'),
    NEW_AGENTS:   find('newly active agents', 'new active agents'),
    AGENT_TARGET: find('new active agent recruitment', 'agent recruitment'),
    NEW_OPS:      find('newly active order points', 'new active order points'),
    OPS_ENROLLED: find('order points enrolled', 'op enrolled'),
    ORDERS_MTD:   find('orders mtd'),
    ORDERS_LM:    find('last month orders', 'orders last month'),
    RUN_RATE:     find('% mtd new active', 'run rate', '% mtd')
  };
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function buildKpis(userRow, allRows, cols) {
  const personalNmv = userRow ? toNum(userRow[cols.PERSONAL_NMV]) : 0;

  // Team NMV: use user's "RCC Team NMV" column if available, else sum all personal NMVs
  let teamNmv = userRow ? toNum(userRow[cols.TEAM_NMV]) : 0;
  if (!teamNmv && cols.PERSONAL_NMV >= 0) {
    teamNmv = allRows.reduce((sum, r) => sum + toNum(r[cols.PERSONAL_NMV]), 0);
  }

  // Personal NMV target: from sheet column if available, else env var, else 0
  const personalTarget = (userRow && cols.NMV_TARGET >= 0)
    ? toNum(userRow[cols.NMV_TARGET])
    : Number(process.env.PERSONAL_TARGET || 0);

  // Team target = sum of all personal targets or env var
  const teamTarget = Number(process.env.TEAM_TARGET || 0)
    || (cols.NMV_TARGET >= 0
      ? allRows.reduce((sum, r) => sum + toNum(r[cols.NMV_TARGET]), 0)
      : 0);

  return {
    personalNmv,
    personalTarget,
    teamNmv,
    teamTarget,
    activeTeam:         allRows.length,
    teamTarget_members: allRows.length
  };
}

// ── Order Points ──────────────────────────────────────────────────────────────
function buildOrderPointsSummary(userRow, allRows, cols) {
  const newActive     = userRow ? toNum(userRow[cols.NEW_OPS])      : 0;
  const lastTwoMonths = userRow ? toNum(userRow[cols.OPS_ENROLLED])  : 0;

  // Build list of all team members with their OP data
  const list = allRows.map(r => ({
    email:          String(r[cols.EMAIL]    || ''),
    city:           String(r[cols.LOCATION] || r[cols.REGION] || ''),
    monthRecruited: String(r[cols.MONTH]    || ''),
    firstActive:    toNum(r[cols.NEW_OPS])  > 0 ? String(r[cols.MONTH] || '') : ''
  })).filter(r => r.email);

  return { lastTwoMonths, newActive, list };
}

// ── Agents ────────────────────────────────────────────────────────────────────
function buildAgentsSummary(userRow, allRows, cols) {
  const newActive     = userRow ? toNum(userRow[cols.NEW_AGENTS])      : 0;
  const lastTwoMonths = userRow ? toNum(userRow[cols.AGENT_TARGET])     : 0;

  const list = allRows.map(r => ({
    email:          String(r[cols.EMAIL]    || ''),
    city:           String(r[cols.LOCATION] || r[cols.REGION] || ''),
    monthRecruited: String(r[cols.MONTH]    || ''),
    firstActive:    toNum(r[cols.NEW_AGENTS]) > 0 ? String(r[cols.MONTH] || '') : ''
  })).filter(r => r.email);

  return { lastTwoMonths, newActive, list };
}

// ── Team list ─────────────────────────────────────────────────────────────────
function buildTeamList(allRows, cols) {
  const list = allRows.map(row => {
    const nmvMtd          = toNum(row[cols.PERSONAL_NMV]);
  
/**
 * Vercel Serverless Function: /api/checkin
 * Appends a field check-in row to the sheet and optionally uploads
 * the photo to Google Drive — all via service account.
 * No Apps Script — no Workspace restrictions.
 *
 * Required environment variables in Vercel:
 *   GOOGLE_SA_KEY    — base64-encoded service account JSON key
 *   SPREADSHEET_ID   — Google Sheet ID
 *
 * Body (JSON): {
 *   userEmail, userName,
 *   activityType, location, notes,
 *   photoBase64, photoName   ← optional
 * }
 */

const { google } = require('googleapis');

const CHECKINS_SHEET = 'RCC_Field_Checkins';

function getAuth(scopes) {
  const keyJson     = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  const credentials = JSON.parse(keyJson);
  return new google.auth.GoogleAuth({ credentials, scopes });
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')   return res.status(405).json({ success: false, error: 'Method not allowed' });

  if (!process.env.GOOGLE_SA_KEY)  return res.status(500).json({ success: false, error: 'GOOGLE_SA_KEY not set' });
  if (!process.env.SPREADSHEET_ID) return res.status(500).json({ success: false, error: 'SPREADSHEET_ID not set' });

  try {
    const payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;

    const userEmail = payload.userEmail || 'unknown@rcc';
    const userName  = (payload.userName && payload.userName.trim()) || getUserDisplayName(userEmail);
    const id        = generateId();
    const now       = new Date();

    // ── Optional photo upload to Drive ───────────────────────────────────────
    let photoUrl        = '';
    let photoPreviewUrl = '';

    if (payload.photoBase64) {
      const driveAuth = getAuth([
        'https://www.googleapis.com/auth/drive.file'
      ]);
      const drive = google.drive({ version: 'v3', auth: driveAuth });

      // Strip the data-URI prefix  (data:image/jpeg;base64,...)
      const base64Data = payload.photoBase64.replace(/^data:image\/\w+;base64,/, '');
      const mimeType   = payload.photoBase64.startsWith('data:image/png') ? 'image/png' : 'image/jpeg';
      const fileName   = `${id}_${now.getTime()}.jpg`;

      const fileBuffer = Buffer.from(base64Data, 'base64');
      const { Readable } = require('stream');
      const stream = Readable.from(fileBuffer);

      const driveRes = await drive.files.create({
        requestBody: {
          name:    fileName,
          mimeType: mimeType
        },
        media: {
          mimeType: mimeType,
          body:     stream
        },
        fields: 'id, webViewLink'
      });

      const fileId = driveRes.data.id;

      // Make the file publicly readable (anyone with link)
      await drive.permissions.create({
        fileId:      fileId,
        requestBody: { role: 'reader', type: 'anyone' }
      });

      photoUrl        = driveRes.data.webViewLink;
      photoPreviewUrl = `https://drive.google.com/uc?id=${fileId}`;
    }

    // ── Append row to RCC_Field_Checkins ─────────────────────────────────────
    // Columns: A=id, B=timestamp, C=email, D=name, E=time, F-G=empty,
    //          H=location, I=photoUrl, J=photoPreviewUrl, K=empty,
    //          L=activityType, M=notes
    const timeStr   = now.toTimeString().slice(0, 8);         // HH:MM:SS
    const isoTs     = now.toISOString();

    const newRow = [
      id,                             // A — unique ID
      isoTs,                          // B — full timestamp
      userEmail,                      // C — RCC email
      userName,                       // D — RCC name
      timeStr,                        // E — time
      '',                             // F — empty
      '',                             // G — empty
      payload.location || '',         // H — GPS coords
      photoUrl,                       // I — Drive view link
      photoPreviewUrl,                // J — Drive direct preview
      '',                             // K — empty
      payload.activityType || '',     // L — activity type
      payload.notes || ''             // M — notes
    ];

    const sheetsAuth = getAuth([
      'https://www.googleapis.com/auth/spreadsheets'
    ]);
    const sheets = google.sheets({ version: 'v4', auth: sheetsAuth });

    await sheets.spreadsheets.values.append({
      spreadsheetId:      process.env.SPREADSHEET_ID,
      range:              `'${CHECKINS_SHEET}'!A:M`,
      valueInputOption:   'USER_ENTERED',
      insertDataOption:   'INSERT_ROWS',
      requestBody: {
        values: [newRow]
      }
    });

    return res.status(200).json({
      success:  true,
      id:       id,
      message:  'Check-in saved successfully',
      photoUrl: photoUrl
    });

  } catch (err) {
    console.error('api/checkin error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};

// ── Helpers ──────────────────────────────────────────────────────────────────
function generateId() {
  return Math.random().toString(36).substring(2, 10);
}

function getUserDisplayName(email) {
  if (!email) return 'RCC User';
  return email.split('@')[0].split(/[._]/)
    .map(s => s.charAt(0).toUpperCase() + s.slice(1).toLowerCase()).join(' ');
}
