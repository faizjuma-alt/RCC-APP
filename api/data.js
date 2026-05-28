/**
 * Vercel Serverless Function: /api/data
 *
 * Sheet: "Current month" tab
 *   B1  = current month selector (e.g. 202605)
 *   Row 2 = headers with Name, Email, NMV columns, target columns, etc.
 *   Row 3+ = one row per RCC for the selected month
 *
 * RCC_Field_Checkins tab column layout (matches checkin.js):
 *   A(0)=ID  B(1)=Timestamp  C(2)=Email  D(3)=Name  E(4)=Time
 *   F(5)=PhotoPreview  G(6)=PhotoURL  H(7)=Location
 *   I(8)=''  J(9)=Notes  K(10)=''  L(11)=ActivityType
 *
 * Required env: GOOGLE_SA_KEY, SPREADSHEET_ID
 * Optional env: CHECKINS_SPREADSHEET_ID, TEAM_TARGET
 */

const { google } = require('googleapis');
const jwt         = require('jsonwebtoken');

const DATA_TAB     = 'Current month';
const CHECKINS_TAB = 'RCC_Field_Checkins';

function getAuth() {
  const key = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(key),
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
  });
}

function requireAuth(req, res) {
  const auth  = req.headers.authorization || '';
  const token = auth.startsWith('Bearer ') ? auth.slice(7).trim() : null;
  if (!token) {
    res.status(401).json({ success: false, error: 'Authentication required' });
    return null;
  }
  try {
    const payload = jwt.verify(token, process.env.JWT_SECRET);
    if (payload.type !== 'session') throw new Error('Invalid token type');
    return payload;
  } catch (err) {
    res.status(401).json({ success: false, error: 'Session expired. Please sign in again.' });
    return null;
  }
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'GET')    return res.status(405).json({ success: false, error: 'Method not allowed' });

  // Auth check
  const authUser = requireAuth(req, res);
  if (!authUser) return;

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

    const selectedMonth = allRows[0] && allRows[0][1]
      ? String(allRows[0][1]).replace(/[^0-9]/g, '')
      : '';

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

    const dataRows = allRows.slice(headerIdx + 1)
      .filter(r => r && r.length > 0 && r[cols.EMAIL]);

    if (action === 'teamlist') {
      return res.status(200).json(buildTeamList(dataRows, cols, selectedMonth));
    }

    // Use email from verified JWT — not from query string (prevents spoofing)
    const userEmail = authUser.email;
    const userName  = authUser.name  || getUserDisplayName(userEmail);

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
  const find = function() {
    var kws = Array.prototype.slice.call(arguments);
    for (var i = 0; i < kws.length; i++) {
      var idx = headerRow.findIndex(function(h) { return h.includes(kws[i]); });
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
    NMV_TARGET:   find('rcc + team nmv target', 'rcc+team nmv target', 'nmv target'),
    NEW_AGENTS:   find('newly active agents', 'new active agents'),
    AGENT_TARGET: find('new active agent recruitment', 'recruitment actual', 'agent recruitment'),
    NEW_OPS:      find('newly active order point', 'new active order point', 'newly active order points'),
    OPS_ENROLLED: find('order point enrolled', 'order points enrolled', 'op enrolled'),
    OP_TARGET:    find('new active op target', 'op target', 'new active order point target')
  };
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function buildKpis(userRow, allRows, cols) {
  const personalNmv = userRow ? toNum(userRow[cols.PERSONAL_NMV]) : 0;
  const teamNmv     = userRow ? toNum(userRow[cols.TEAM_NMV])     : 0;
  const sheetTarget = (userRow && cols.NMV_TARGET >= 0) ? toNum(userRow[cols.NMV_TARGET]) : 0;
  const teamTarget  = sheetTarget || Number(process.env.TEAM_TARGET || 0);
  return {
    personalNmv,
    personalTarget: Number(process.env.PERSONAL_TARGET || 0),
    teamNmv,
    teamTarget,
    activeTeam:         allRows.length,
    teamTarget_members: allRows.length
  };
}

// ── Progress ──────────────────────────────────────────────────────────────────
function buildProgress(userRow, cols) {
  if (!userRow) return {
    nmv:    { actual: 0, target: 0 },
    agents: { actual: 0, target: 0 },
    ops:    { actual: 0, target: 0 }
  };
  return {
    nmv: {
      actual: toNum(userRow[cols.PERSONAL_NMV]) + toNum(userRow[cols.TEAM_NMV]),
      target: toNum(userRow[cols.NMV_TARGET])
    },
    agents: {
      actual: toNum(userRow[cols.NEW_AGENTS]),
      target: toNum(userRow[cols.AGENT_TARGET])
    },
    ops: {
      actual: toNum(userRow[cols.NEW_OPS]),
      target: cols.OP_TARGET >= 0 ? toNum(userRow[cols.OP_TARGET]) : 0
    }
  };
}

// ── Order Points ──────────────────────────────────────────────────────────────
function buildOrderPointsSummary(userRow, cols) {
  return {
    newActive:     userRow ? toNum(userRow[cols.NEW_OPS])      : 0,
    lastTwoMonths: userRow ? toNum(userRow[cols.OPS_ENROLLED]) : 0,
    list: []
  };
}

// ── Agents ────────────────────────────────────────────────────────────────────
function buildAgentsSummary(userRow, cols) {
  return {
    newActive:     userRow ? toNum(userRow[cols.NEW_AGENTS])   : 0,
    lastTwoMonths: userRow ? toNum(userRow[cols.AGENT_TARGET]) : 0,
    list: []
  };
}

// ── Team list ─────────────────────────────────────────────────────────────────
function buildTeamList(rows, cols, month) {
  return { success: true, month, list: [] };
}

// ── Field check-ins ───────────────────────────────────────────────────────────
async function getFieldCheckinsSummary(sheets, sid, userEmail) {
  const checkSid = process.env.CHECKINS_SPREADSHEET_ID || sid;
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId:     checkSid,
      range:             `'${CHECKINS_TAB}'!A2:L`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const rows = r.data.values || [];
    if (!rows.length) return emptyCheckins();

    const now          = new Date();
    const startOfDay   = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const startOfWeek  = new Date(startOfDay);
    startOfWeek.setDate(startOfDay.getDate() - startOfDay.getDay());
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

    var today = 0, thisWeek = 0, thisMonth = 0;
    var locations = new Set();
    var recent    = [];

    rows.forEach(function(row) {
      // Filter to this user's check-ins only
      var rowEmail = row[2] ? String(row[2]).toLowerCase().trim() : '';
      if (userEmail && rowEmail !== userEmail) return;

      var ts = row[1] ? new Date(row[1]) : null;
      if (!ts || isNaN(ts.getTime())) return;
      if (ts >= startOfMonth) thisMonth++;
      if (ts >= startOfWeek)  thisWeek++;
      if (ts >= startOfDay)   today++;

      var loc = row[7] ? String(row[7]) : '';
      if (loc) locations.add(loc.substring(0, 20));

      // Column layout: F(5)=PhotoPreview  H(7)=Location  J(9)=Notes  L(11)=ActivityType
      var photo        = row[5]  ? String(row[5])  : '';
      var activityType = row[11] ? String(row[11]) : '';
      var notes        = row[9]  ? String(row[9])  : '';

      recent.push({
        id:           row[0]  || '',
        timestamp:    ts.toISOString(),
        rccEmail:     row[2]  || '',
        rccName:      row[3]  || '',
        location:     loc,
        photo:        photo,
        activityType: activityType,
        notes:        notes
      });
    });

    recent.sort(function(a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });
    return { today: today, thisWeek: thisWeek, thisMonth: thisMonth, locations: locations.size, recent: recent.slice(0, 10) };

  } catch (e) {
    console.error('getFieldCheckinsSummary error:', e.message);
    return emptyCheckins();
  }
}

function emptyCheckins() {
  return { today: 0, thisWeek: 0, thisMonth: 0, locations: 0, recent: [] };
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function toNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  var n = Number(String(val).replace(/[^0-9.-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function getUserDisplayName(email) {
  if (!email) return 'RCC User';
  return email.split('@')[0].split(/[._]/)
    .map(function(s) { return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); })
    .join(' ');
}

function getInitials(nameOrEmail) {
  if (!nameOrEmail) return '??';
  var name = nameOrEmail.includes('@') ? getUserDisplayName(nameOrEmail) : nameOrEmail;
  return name.trim().split(/\s+/).filter(function(p) { return p.length > 0; }).slice(0, 2)
    .map(function(p) { return p[0].toUpperCase(); }).join('');
}
