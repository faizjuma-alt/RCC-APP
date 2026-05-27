/**
 * Vercel Serverless Function: /api/data
 *
 * Sheet: "Current month" tab
 *   B1  = current month selector (e.g. 202605) — user switches this in the sheet
 *   Row 2 = headers: Start date | End date | Name | Email | Region | Location |
 *            New active agent recruitment | Recruitment actual | Newly active agents |
 *            % MTD | RCC NMV(LCY) | RCC NMV ($) | RCC Team NMV (LCY) | Team NMV ($) |
 *            RCC+Team NMV $ | RCC + team NMV target | ... |
 *            Order Point enrolled | Newly active order point | New active OP target | ...
 *   Row 3+ = one row per RCC for the selected month (sheet already filters by B1)
 *
 * "Start date" = when the RCC joined — NOT a month column, not used for filtering.
 * The sheet handles month filtering via B1; the API just reads whatever rows are visible.
 *
 * Required env vars:
 *   GOOGLE_SA_KEY   — base64-encoded service account JSON key
 *   SPREADSHEET_ID  — Regional sheet ID
 * Optional:
 *   CHECKINS_SPREADSHEET_ID — for field check-ins tab
 *   TEAM_TARGET     — NMV target override (if not in sheet)
 */

const { google } = require('googleapis');

const DATA_TAB     = 'Current month';
const CHECKINS_TAB = 'RCC_Field_Checkins';

function getAuth() {
  const key = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(key),
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
  });
}

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

    const raw = await sheets.spreadsheets.values.get({
      spreadsheetId:     sid,
      range:             `'${DATA_TAB}'!A1:AZ`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const allRows = raw.data.values || [];

    // B1 = currently selected month (e.g. 202605)
    const selectedMonth = allRows[0] && allRows[0][1]
      ? String(allRows[0][1]).replace(/[^0-9]/g, '')
      : '';

    // Find header row (contains both "email" and "name")
    let headerIdx = -1;
    let cols      = {};
    for (let i = 0; i < Math.min(allRows.length, 6); i++) {
      const row = allRows[i].map(c => String(c || '').toLowerCase().trim());
      if (row.some(c => c.includes('email')) && row.some(c => c.includes('name'))) {
        headerIdx = i;
        cols = mapColumns(row);
        break;
      }
    }
    if (headerIdx < 0) {
      return res.status(500).json({ success: false, error: 'Header row not found. Sheet must have Name and Email columns.' });
    }

    // Data rows — sheet already shows only the selected month's RCCs
    const dataRows = allRows.slice(headerIdx + 1)
      .filter(r => r && r.length > 0 && r[cols.EMAIL]);

    if (action === 'teamlist') {
      return res.status(200).json(buildTeamList(dataRows, cols, selectedMonth));
    }

    // ── Dashboard ──────────────────────────────────────────────────────────────
    const userEmail = (req.query.email || '').toLowerCase().trim();
    const userName  = req.query.name  || getUserDisplayName(userEmail);

    const userRow = dataRows.find(r =>
      String(r[cols.EMAIL] || '').toLowerCase().trim() === userEmail
    ) || null;

    const kpis          = buildKpis(userRow, dataRows, cols);
    const orderPoints   = buildOrderPointsSummary(userRow, cols);
    const agents        = buildAgentsSummary(userRow, cols);
    const progress      = buildProgress(userRow, cols);
    const fieldCheckins = await getFieldCheckinsSummary(sheets, sid, userEmail);

    const resolvedName = (userRow && userRow[cols.NAME])
      ? String(userRow[cols.NAME]).trim()
      : userName;

    return res.status(200).json({
      success: true,
      user: { email: userEmail, name: resolvedName, initials: getInitials(resolvedName) },
      month: selectedMonth,
      kpis, orderPoints, agents, progress, fieldCheckins,
      timestamp: new Date().toISOString()
    });

  } catch (err) {
    console.error('api/data error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};

// ── Column mapper ─────────────────────────────────────────────────────────────
function mapColumns(headerRow) {
  const find = (...kws) => {
    for (const kw of kws) {
      const idx = headerRow.findIndex(h => h.includes(kw));
      if (idx >= 0) return idx;
    }
    return -1;
  };
  return {
    NAME:         find('name'),
    EMAIL:        find('email'),
    REGION:       find('region'),
    LOCATION:     find('location'),
    PERSONAL_NMV: find('rcc nmv(lcy)', 'rcc nmv (lcy)', 'rcc nmv lcy', 'personal nmv'),
    TEAM_NMV:     find('rcc team nmv (lcy)', 'rcc team nmv(lcy)', 'team nmv (lcy)', 'rcc team nmv'),
    NMV_TARGET:   find('rcc + team nmv target', 'nmv target'),
    NEW_AGENTS:   find('newly active agents', 'new active agents'),
    AGENT_TARGET: find('new active agent recruitment', 'recruitment actual', 'agent recruitment'),
    NEW_OPS:      find('newly active order point', 'new active order point', 'newly active order points'),
    OPS_ENROLLED: find('order point enrolled', 'order points enrolled', 'op enrolled'),
    OP_TARGET:    find('new active op target', 'op target', 'new active order point target')
  };
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function buildKpis(userRow, allRows, cols) {
  const personalNmv    = userRow ? toNum(userRow[cols.PERSONAL_NMV]) : 0;
  const teamNmv        = userRow ? toNum(userRow[cols.TEAM_NMV])     : 0;
  const sheetTarget    = (userRow && cols.NMV_TARGET >= 0) ? toNum(userRow[cols.NMV_TARGET]) : 0;
  const teamTarget     = sheetTarget || Number(process.env.TEAM_TARGET || 0);
  const personalTarget = Number(process.env.PERSONAL_TARGET || 0);
  return {
    personalNmv, personalTarget, teamNmv, teamTarget,
    activeTeam: allRows.length, teamTarget_members: allRows.length
  };
}

// ── Progress vs Target (3-bar section) ───────────────────────────────────────
function buildProgress(userRow, cols) {
  if (!userRow) return {
    nmv:    { actual: 0, target: 0 },
    agents: { actual: 0, target: 0 },
    ops:    { actual: 0, target: 0 }
  };

  const personalNmv = toNum(userRow[cols.PERSONAL_NMV]);
  const teamNmv     = toNum(userRow[cols.TEAM_NMV]);
  const nmvTarget   = toNum(userRow[cols.NMV_TARGET]);

  const newAgents   = toNum(userRow[cols.NEW_AGENTS]);
  const agentTarget = toNum(userRow[cols.AGENT_TARGET]);

  const newOps      = toNum(userRow[cols.NEW_OPS]);
  const opTarget    = cols.OP_TARGET >= 0 ? toNum(userRow[cols.OP_TARGET]) : 0;

  return {
    nmv:    { actual: personalNmv + teamNmv, target: nmvTarget },
    agents: { actual: newAgents,  target: agentTarget },
    ops:    { actual: newOps,     target: opTarget }
  };
}

// ── Order Points ──────────────────────────────────────────────────────────────
function buildOrderPointsSummary(userRow, cols) {
  return {
    newActive:     userRow ? toNum(userRow[cols.NEW_OPS])      : 0,
    lastTwoMonths: userRow ? toNum(userRow[cols.OPS_ENROLLED])  : 0,
    list: []
  };
}

// ── Agents ────────────────────────────────────────────────────────────────────
function buildAgentsSummary(userRow, cols) {
  return {
    newActive:     userRow ? toNum(userRow[cols.NEW_AGENTS])   : 0,
    lastTwoMonths: userRow ? toNum(userRow[cols.AGENT_TARGET])  : 0,
    list: []
  };
}

// ── Team list ─────────────────────────────────────────────────────────────────
function buildTeamList(rows, cols, month) {
  const list = rows.map(row => {
    const emailVal        = String(row[cols.EMAIL] || '');
    const nameVal         = String(row[cols.NAME]  || emailVal);
    const ordersMtd       = toNum(row[cols.NEW_AGENTS]);
    const ordersLastMonth = toNum(row[cols.AGENT_TARGET]);
    return {
      email: emailVal, name: nameVal, initials: getInitials(nameVal || emailVal),
      nmvMtd:        toNum(row[cols.PERSONAL_NMV]),
      nmvLastMonth:  0,
      ordersMtd, ordersLastMonth,
      runRate: ordersLastMonth > 0 ? Math.round((ordersMtd / ordersLastMonth) * 100) : 0,
      location: String(row[cols.LOCATION] || row[cols.REGION] || '')
    };
  }).filter(r => r.email);
  return { success: true, month, list };
}

// ── Field check-ins ───────────────────────────────────────────────────────────
async function getFieldCheckinsSummary(sheets, sid, userEmail) {
  const checkSid = process.env.CHECKINS_SPREADSHEET_ID || sid;
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: checkSid,
      range:         `'${CHECKINS_TAB}'!A2:M`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const rows = r.data.values || [];
    if (!rows.length) return emptyCheckins();

    const now          = new Date();
    const startOfDay   = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const startOfWeek  = new Date(startOfDay);
    startOfWeek.setDate(startOfDay.getDate() - startOfDay.getDay());
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

    let today = 0, thisWeek = 0, thisMonth = 0;
    const locations = new Set();
    const recent    = [];

    rows.forEach(row => {
      const ts = row[1] ? new Date(row[1]) : null;
      if (!ts || isNaN(ts.getTime())) return;
      if (ts >= startOfMonth) thisMonth++;
      if (ts >= startOfWeek)  thisWeek++;
      if (ts >= startOfDay)   today++;
      if (row[7]) locations.add(String(row[7]).substring(0, 20));
      recent.push({ id: row[0]||'', timestamp: ts.toISOString(),
        rccEmail: row[2]||'', rccName: row[3]||'', location: row[7]||'',
        photo: row[8]||'', activityType: row[11]||'', notes: row[12]||'' });
    });

    recent.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return { today, thisWeek, thisMonth, locations: locations.size, recent: recent.slice(0, 10) };
  } catch (e) {
    return emptyCheckins();
  }
}

function emptyCheckins() {
  return { today: 0, thisWeek: 0, thisMonth: 0, locations: 0, recent: [] };
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function toNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  const n = Number(String(val).replace(/[^0-9.-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function getUserDisplayName(email) {
  if (!email) return 'RCC User';
  return email.split('@')[0].split(/[._]/)
    .map(s => s.charAt(0).toUpperCase() + s.slice(1).toLowerCase()).join(' ');
}

function getInitials(nameOrEmail) {
  if (!nameOrEmail) return '??';
  const name = nameOrEmail.includes('@') ? getUserDisplayName(nameOrEmail) : nameOrEmail;
  return name.trim().split(/\s+/).filter(p => p.length > 0).slice(0, 2)
    .map(p => p[0].toUpperCase()).join('');
                }
