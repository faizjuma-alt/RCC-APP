/**
 * Vercel Serverless Function: /api/data
 * Reads from the regional "Current month" sheet.
 *
 * Sheet structure (headers in row 2):
 *   Start date | End date | Name | Email | Region | Location |
 *   New active agent recruitment | Recruitment actual | Newly active agents |
 *   % MTD new active agents | RCC NMV(LCY) | RCC NMV ($) |
 *   RCC Team NMV (LCY) | Team NMV ($) | RCC+Team NMV $ |
 *   RCC + team NMV target | ... | Order Point enrolled |
 *   Newly active order point | New active OP target | ...
 *
 * Each row = one RCC (independent, not hierarchical).
 *
 * Required env vars:
 *   GOOGLE_SA_KEY   — base64-encoded service account JSON key
 *   SPREADSHEET_ID  — Regional sheet ID
 *
 * Optional env vars:
 *   CHECKINS_SPREADSHEET_ID — Sheet ID for field check-ins (defaults to SPREADSHEET_ID)
 *   TEAM_TARGET — Regional NMV target override (if not in sheet)
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
      spreadsheetId:    sid,
      range:            `'${DATA_TAB}'!A1:AZ`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const allRows = raw.data.values || [];

    // Find header row — contains both "email" and "name"
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
      return res.status(500).json({ success: false, error: 'Could not find header row in "' + DATA_TAB + '" tab. Make sure it has Email and Name columns.' });
    }

    const dataRows = allRows.slice(headerIdx + 1).filter(r => r && r.length > 0 && r[0]);

    // Filter to current month
    const currentMonth = getCurrentMonth();   // e.g. 202605
    const monthRows = dataRows.filter(row => {
      const monthVal = String(row[cols.MONTH] || '').replace(/[^0-9]/g, '');
      return monthVal === String(currentMonth);
    });

    // Fall back to most recent month if no current-month rows
    const workingRows = monthRows.length > 0
      ? monthRows
      : getMostRecentMonthRows(dataRows, cols.MONTH);

    if (action === 'teamlist') {
      return res.status(200).json(buildTeamList(workingRows, cols));
    }

    // Dashboard — find logged-in user's row
    const userEmail = (req.query.email || '').toLowerCase().trim();
    const userName  = req.query.name  || getUserDisplayName(userEmail);

    const userRow = workingRows.find(r =>
      String(r[cols.EMAIL] || '').toLowerCase().trim() === userEmail
    ) || null;

    const kpis          = buildKpis(userRow, workingRows, cols);
    const orderPoints   = buildOrderPointsSummary(userRow, cols);
    const agents        = buildAgentsSummary(userRow, cols);
    const fieldCheckins = await getFieldCheckinsSummary(sheets, sid, userEmail);

    const resolvedName = (userRow && userRow[cols.NAME])
      ? String(userRow[cols.NAME]).trim()
      : userName;

    return res.status(200).json({
      success: true,
      user: { email: userEmail, name: resolvedName, initials: getInitials(resolvedName) },
      kpis, orderPoints, agents, fieldCheckins,
      timestamp: new Date().toISOString()
    });

  } catch (err) {
    console.error('api/data error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};

// ── Column mapper ─────────────────────────────────────────────────────────────
// Uses case-insensitive substring matching on the header row.
// Multiple keywords tried in order — first match wins.
function mapColumns(headerRow) {
  const find = (...keywords) => {
    for (const kw of keywords) {
      const idx = headerRow.findIndex(h => h.includes(kw));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  return {
    // "Start date" column holds YYYYMM month codes (e.g. 202605)
    MONTH:        find('start date', 'month'),
    NAME:         find('name'),
    EMAIL:        find('email'),
    REGION:       find('region'),
    LOCATION:     find('location'),
    // RCC NMV(LCY) — no space before paren
    PERSONAL_NMV: find('rcc nmv(lcy)', 'rcc nmv (lcy)', 'rcc nmv lcy', 'personal nmv'),
    // RCC Team NMV (LCY) — the RCC's own team of agents' NMV
    TEAM_NMV:     find('rcc team nmv (lcy)', 'rcc team nmv(lcy)', 'team nmv (lcy)', 'rcc team nmv'),
    // RCC + team NMV target
    NMV_TARGET:   find('rcc + team nmv target', 'nmv target'),
    // Agents
    NEW_AGENTS:   find('newly active agents', 'new active agents'),
    AGENT_TARGET: find('new active agent recruitment', 'agent recruitment', 'recruitment actual'),
    // Order Points — headers use singular "order point"
    NEW_OPS:      find('newly active order point', 'new active order point', 'newly active order points'),
    OPS_ENROLLED: find('order point enrolled', 'order points enrolled', 'op enrolled'),
    RUN_RATE:     find('% mtd new active', 'run rate', '% mtd')
  };
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
// Each RCC is independent. Personal = their own sales. Team = their agents' NMV.
function buildKpis(userRow, allRows, cols) {
  const personalNmv = userRow ? toNum(userRow[cols.PERSONAL_NMV]) : 0;

  // Team NMV = this RCC's own team of agents (NOT sum of all RCCs)
  const teamNmv = userRow ? toNum(userRow[cols.TEAM_NMV]) : 0;

  // Target comes from the user's own row, or env var
  const rawTarget = (userRow && cols.NMV_TARGET >= 0)
    ? toNum(userRow[cols.NMV_TARGET])
    : 0;
  const teamTarget = rawTarget || Number(process.env.TEAM_TARGET || 0);
  const personalTarget = Number(process.env.PERSONAL_TARGET || 0);

  return {
    personalNmv,
    personalTarget,
    teamNmv,
    teamTarget,
    activeTeam:         allRows.length,   // total RCCs in region
    teamTarget_members: allRows.length
  };
}

// ── Order Points ──────────────────────────────────────────────────────────────
// Returns counts from the user's own row only.
// No individual OP sub-records in this sheet — list is empty.
function buildOrderPointsSummary(userRow, cols) {
  const newActive     = userRow ? toNum(userRow[cols.NEW_OPS])      : 0;
  const lastTwoMonths = userRow ? toNum(userRow[cols.OPS_ENROLLED])  : 0;
  return { lastTwoMonths, newActive, list: [] };
}

// ── Agents ────────────────────────────────────────────────────────────────────
// Returns counts from the user's own row only.
// No individual agent sub-records in this sheet — list is empty.
function buildAgentsSummary(userRow, cols) {
  const newActive     = userRow ? toNum(userRow[cols.NEW_AGENTS])   : 0;
  const lastTwoMonths = userRow ? toNum(userRow[cols.AGENT_TARGET])  : 0;
  return { lastTwoMonths, newActive, list: [] };
}

// ── Team list (manager view) ──────────────────────────────────────────────────
// Returns all RCCs in the region for the Team tab.
function buildTeamList(allRows, cols) {
  const list = allRows.map(row => {
    const nmvMtd          = toNum(row[cols.PERSONAL_NMV]);
    const ordersMtd       = toNum(row[cols.NEW_AGENTS]);
    const ordersLastMonth = toNum(row[cols.AGENT_TARGET]);
    const runRate = ordersLastMonth > 0
      ? Math.round((ordersMtd / ordersLastMonth) * 100)
      : 0;
    const emailVal = String(row[cols.EMAIL] || '');
    const nameVal  = String(row[cols.NAME]  || emailVal);
    return {
      email: emailVal, name: nameVal, initials: getInitials(nameVal || emailVal),
      ordersMtd, ordersLastMonth, nmvMtd, nmvLastMonth: 0, runRate,
      location: String(row[cols.LOCATION] || row[cols.REGION] || '')
    };
  }).filter(r => r.email);
  return { success: true, list };
}

// ── Field check-ins ───────────────────────────────────────────────────────────
async function getFieldCheckinsSummary(sheets, sid, userEmail) {
  const checkSid = process.env.CHECKINS_SPREADSHEET_ID || sid;
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: checkSid,
      range: `'${CHECKINS_TAB}'!A2:M`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const rows = r.data.values || [];
    if (!rows.length) return emptyCheckins();

    const now = new Date();
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
      recent.push({
        id: row[0]||'', timestamp: ts.toISOString(),
        rccEmail: row[2]||'', rccName: row[3]||'',
        location: row[7]||'', photo: row[8]||'',
        activityType: row[11]||'', notes: row[12]||''
      });
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
function getCurrentMonth() {
  const now = new Date();
  return now.getFullYear() * 100 + (now.getMonth() + 1);
}

function getMostRecentMonthRows(dataRows, monthCol) {
  if (monthCol < 0 || !dataRows.length) return dataRows;
  const months = dataRows
    .map(r => Number(String(r[monthCol] || '').replace(/[^0-9]/g, '')))
    .filter(m => m > 0);
  if (!months.length) return dataRows;
  const latest = Math.max(...months);
  return dataRows.filter(r =>
    String(r[monthCol] || '').replace(/[^0-9]/g, '') === String(latest)
  );
}

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
