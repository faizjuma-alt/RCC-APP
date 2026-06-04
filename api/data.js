/**
 * Vercel Serverless Function: /api/data
 *
 * Sheet tabs:
 *   "Current month" — B1=month selector, row 2=headers, row 3+=RCC data
 *   "RCC_Field_Checkins" — A=ID B=Timestamp C=Email D=Name E=Time
 *                          F=PhotoPreview G=PhotoURL H=Location I='' J=Notes K='' L=ActivityType
 *   "Managers" — col A=email, col B=name, col C=market (KE / NG / ALL)
 *                "ALL" = admin/head access — sees every region
 *
 * Actions:
 *   GET ?action=dashboard  — personal RCC dashboard (default)
 *   GET ?action=manager    — full team view (managers only)
 *   GET ?action=teamlist   — lightweight team list
 *   GET ?action=rcc&email= — single RCC detail (managers only)
 *   GET ?action=market     — regional leaderboard (managers + RCCs)
 *
 * Required env: GOOGLE_SA_KEY, SPREADSHEET_ID
 * Optional env: CHECKINS_SPREADSHEET_ID, TEAM_TARGET
 */

const { google } = require('googleapis');
const jwt         = require('jsonwebtoken');

const DATA_TAB     = 'Current month';
const CHECKINS_TAB = 'RCC_Field_Checkins';
const MANAGERS_TAB = 'Managers';

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
      return res.status(500).json({ success: false, error: 'Header row not found in sheet.' });
    }

    const dataRows = allRows.slice(headerIdx + 1)
      .filter(r => r && r.length > 0 && r[cols.EMAIL]);

    const userEmail = authUser.email;
    const userName  = authUser.name || getUserDisplayName(userEmail);
    const managerInfo = await getManagerInfo(sheets, sid, userEmail);
    const isManager   = managerInfo.isManager;
    const managerMarket = managerInfo.market;

    let userRegion = '';
    if (!isManager) {
      const userRowRaw = dataRows.find(r =>
        String(r[cols.EMAIL] || '').toLowerCase().trim() === userEmail
      );
      userRegion = userRowRaw && cols.REGION >= 0
        ? String(userRowRaw[cols.REGION] || '').trim().toUpperCase()
        : '';
    }

    let filteredRows = dataRows;
    if (isManager && managerMarket && managerMarket.toUpperCase() !== 'ALL') {
      filteredRows = dataRows.filter(r =>
        String(r[cols.REGION] || '').trim().toUpperCase() === managerMarket.toUpperCase()
      );
    } else if (!isManager && userRegion) {
      filteredRows = dataRows.filter(r =>
        String(r[cols.REGION] || '').trim().toUpperCase() === userRegion
      );
    }

    const resolvedName = getUserDisplayName(userEmail);

    if (action === 'teamlist') {
      return res.status(200).json(buildTeamList(filteredRows, cols, selectedMonth));
    }

    if (action === 'manager') {
      if (!isManager) return res.status(403).json({ success: false, error: 'Manager access required' });
      const teamData    = buildManagerTeamData(filteredRows, cols, selectedMonth);
      const teamCheckins = await getCheckins(sheets, sid, null);
      return res.status(200).json({
        success: true,
        role: 'manager',
        market: managerMarket,
        user: { email: userEmail, name: resolvedName, initials: getInitials(resolvedName) },
        month: selectedMonth,
        teamData,
        teamCheckins,
        timestamp: new Date().toISOString()
      });
    }

    if (action === 'market') {
      const regionMap = {};
      filteredRows.forEach(function(r) {
        var region = cols.REGION >= 0
          ? String(r[cols.REGION] || 'Unknown').trim().toUpperCase()
          : 'Unknown';
        if (!regionMap[region]) regionMap[region] = [];
        regionMap[region].push(r);
      });

      var regions = Object.keys(regionMap).map(function(region) {
        var rows = regionMap[region];
        var rccs = rows.map(function(r) {
          var personalNmv = toNum(r[cols.PERSONAL_NMV]);
          var teamNmv     = toNum(r[cols.TEAM_NMV]);
          return {
            name:       String(r[cols.NAME]  || '').trim(),
            email:      String(r[cols.EMAIL] || '').toLowerCase().trim(),
            location:   cols.LOCATION  >= 0 ? String(r[cols.LOCATION]  || '').trim() : '',
            personalNmv,
            teamNmv,
            totalNmv:   personalNmv + teamNmv,
            nmvTarget:  cols.NMV_TARGET >= 0 ? toNum(r[cols.NMV_TARGET]) : 0,
            audits:     cols.AUDITS     >= 0 ? toNum(r[cols.AUDITS])     : 0,
            payoutUsd:  cols.PAYOUT_USD >= 0 ? toNum(r[cols.PAYOUT_USD]) : 0,
            payoutLcy:  cols.PAYOUT_LCY >= 0 ? toNum(r[cols.PAYOUT_LCY]) : 0,
            agents:     toNum(r[cols.NEW_AGENTS]),
            ops:        toNum(r[cols.NEW_OPS])
          };
        });
        rccs.sort(function(a, b) { return b.totalNmv - a.totalNmv; });

        var totalNmv    = rccs.reduce(function(s, x) { return s + x.totalNmv;   }, 0);
        var totalAudits = rccs.reduce(function(s, x) { return s + x.audits;     }, 0);
        var totalPayout = rccs.reduce(function(s, x) { return s + x.payoutUsd;  }, 0);
        var totalAgents = rccs.reduce(function(s, x) { return s + x.agents;     }, 0);
        var totalOps    = rccs.reduce(function(s, x) { return s + x.ops;        }, 0);

        return {
          region,
          rccCount:   rccs.length,
          totalNmv,
          avgNmv:     rccs.length > 0 ? Math.round(totalNmv / rccs.length) : 0,
          totalAudits,
          totalPayout,
          totalAgents,
          totalOps,
          topRcc:     rccs[0] || null,
          rccs
        };
      });

      regions.sort(function(a, b) { return b.totalNmv - a.totalNmv; });

      return res.status(200).json({
        success: true,
        role:  isManager ? 'manager' : 'rcc',
        market: isManager ? managerMarket : userRegion,
        user:  { email: userEmail, name: resolvedName, initials: getInitials(resolvedName) },
        month: selectedMonth,
        regions,
        timestamp: new Date().toISOString()
      });
    }

    if (action === 'rcc') {
      if (!isManager) return res.status(403).json({ success: false, error: 'Manager access required' });
      const targetEmail = (req.query.email || '').toLowerCase().trim();
      const rccRow = filteredRows.find(r =>
        String(r[cols.EMAIL] || '').toLowerCase().trim() === targetEmail
      ) || null;
      const checkins = await getCheckins(sheets, sid, targetEmail);
      return res.status(200).json({
        success: true,
        email: targetEmail,
        name: rccRow ? String(rccRow[cols.NAME] || '').trim() : getUserDisplayName(targetEmail),
        progress: buildProgress(rccRow, cols),
        kpis:     buildKpis(rccRow, filteredRows, cols),
        payout:   buildPayout(rccRow, cols),
        checkins,
        timestamp: new Date().toISOString()
      });
    }

    const userRow = filteredRows.find(r =>
      String(r[cols.EMAIL] || '').toLowerCase().trim() === userEmail
    ) || null;

    const kpis        = buildKpis(userRow, filteredRows, cols);
    const orderPoints = buildOrderPointsSummary(userRow, cols);
    const agents      = buildAgentsSummary(userRow, cols);
    const progress    = buildProgress(userRow, cols);
    const payout      = buildPayout(userRow, cols);
    const fieldCheckins = await getCheckins(sheets, sid, userEmail);

    const displayName = (userRow && userRow[cols.NAME])
      ? String(userRow[cols.NAME]).trim()
      : userName;

    return res.status(200).json({
      success: true,
      role:   isManager ? 'manager' : 'rcc',
      market: isManager ? managerMarket : userRegion,
      user:   { email: userEmail, name: displayName, initials: getInitials(displayName) },
      month:  selectedMonth,
      kpis, orderPoints, agents, progress, payout,
      fieldCheckins,
      timestamp: new Date().toISOString()
    });

  } catch (err) {
    console.error('api/data error:', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};

async function getManagerInfo(sheets, sid, email) {
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId:     sid,
      range:             `'${MANAGERS_TAB}'!A1:C`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const rows = r.data.values || [];
    const row  = rows.find(function(row) {
      return row[0] && String(row[0]).toLowerCase().trim() === email;
    });
    if (!row) return { isManager: false, market: '' };
    var market = row[2] ? String(row[2]).trim().toUpperCase() : 'ALL';
    return { isManager: true, market };
  } catch (e) {
    console.warn('Managers tab error:', e.message);
    return { isManager: false, market: '' };
  }
}

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
    NAME:               find('name'),
    EMAIL:              find('email'),
    REGION:             find('markets', 'market', 'region'),
    LOCATION:           find('location'),
    PERSONAL_NMV:       find('rcc nmv(lcy)', 'rcc nmv (lcy)', 'rcc nmv lcy', 'personal nmv'),
    TEAM_NMV:           find('rcc team nmv (lcy)', 'rcc team nmv(lcy)', 'team nmv (lcy)', 'rcc team nmv'),
    NMV_TARGET:         find('target (lcy) rcc', 'rcc + team nmv target', 'rcc+team nmv target', 'nmv target'),
    NEW_AGENTS:         find('newly active agents', 'new active agents'),
    AGENT_TARGET:       find('new active agent recruitment', 'recruitment actual', 'agent recruitment'),
    NEW_OPS:            find('newly active order point', 'new active order point', 'newly active order points'),
    OPS_ENROLLED:       find('order point enrolled', 'order points enrolled', 'op enrolled'),
    OP_TARGET:          find('new active op target', 'op target', 'new active order point target'),
    AUDITS:             find('audit'),
    PAYOUT_USD:         find('total payout $', 'payout $', 'payout usd'),
    PAYOUT_LCY:         find('total payout lcy', 'payout lcy mtd', 'payout lcy'),
    AGENT_TEAM_NMV_USD: find('agent + team nmv ($)', 'agent + team nmv', 'agent+team nmv')
  };
}

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

function buildPayout(userRow, cols) {
  return {
    audits:         userRow && cols.AUDITS             >= 0 ? toNum(userRow[cols.AUDITS])             : 0,
    payoutUsd:      userRow && cols.PAYOUT_USD         >= 0 ? toNum(userRow[cols.PAYOUT_USD])         : 0,
    payoutLcy:      userRow && cols.PAYOUT_LCY         >= 0 ? toNum(userRow[cols.PAYOUT_LCY])         : 0,
    agentTeamNmvUsd:userRow && cols.AGENT_TEAM_NMV_USD >= 0 ? toNum(userRow[cols.AGENT_TEAM_NMV_USD]) : 0
  };
}

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

function buildOrderPointsSummary(userRow, cols) {
  return {
    newActive:     userRow ? toNum(userRow[cols.NEW_OPS])      : 0,
    lastTwoMonths: userRow ? toNum(userRow[cols.OPS_ENROLLED]) : 0,
    list: []
  };
}

function buildAgentsSummary(userRow, cols) {
  return {
    newActive:     userRow ? toNum(userRow[cols.NEW_AGENTS])   : 0,
    lastTwoMonths: userRow ? toNum(userRow[cols.AGENT_TARGET]) : 0,
    list: []
  };
}

function buildTeamList(rows, cols, month) {
  const list = rows.map(function(r) {
    return {
      name:   String(r[cols.NAME]  || '').trim(),
      email:  String(r[cols.EMAIL] || '').toLowerCase().trim(),
      region: cols.REGION >= 0 ? String(r[cols.REGION] || '').trim() : ''
    };
  });
  return { success: true, month, list };
}

function buildManagerTeamData(rows, cols, month) {
  var totalNmv         = 0;
  var totalPersonalNmv = 0;
  var totalTeamNmv     = 0;
  var totalAgents      = 0;
  var totalOps         = 0;
  var totalTarget      = 0;
  var totalAgentTarget = 0;
  var totalOpTarget    = 0;
  var totalAudits      = 0;
  var totalPayoutUsd   = 0;
  var totalPayoutLcy   = 0;

  var leaderboard = rows.map(function(r) {
    var personalNmv    = toNum(r[cols.PERSONAL_NMV]);
    var teamNmv        = toNum(r[cols.TEAM_NMV]);
    var totalNmvRow    = personalNmv + teamNmv;
    var target         = cols.NMV_TARGET >= 0 ? toNum(r[cols.NMV_TARGET]) : 0;
    var agents         = toNum(r[cols.NEW_AGENTS]);
    var agentTarget    = toNum(r[cols.AGENT_TARGET]);
    var ops            = toNum(r[cols.NEW_OPS]);
    var opTarget       = cols.OP_TARGET >= 0 ? toNum(r[cols.OP_TARGET]) : 0;
    var audits         = cols.AUDITS     >= 0 ? toNum(r[cols.AUDITS])   : 0;
    var payoutUsd      = cols.PAYOUT_USD >= 0 ? toNum(r[cols.PAYOUT_USD]) : 0;
    var payoutLcy      = cols.PAYOUT_LCY >= 0 ? toNum(r[cols.PAYOUT_LCY]) : 0;
    var agentTeamNmvUsd= cols.AGENT_TEAM_NMV_USD >= 0 ? toNum(r[cols.AGENT_TEAM_NMV_USD]) : 0;
    var pct            = target > 0 ? Math.round((totalNmvRow / target) * 100) : 0;

    totalNmv         += totalNmvRow;
    totalPersonalNmv += personalNmv;
    totalTeamNmv     += teamNmv;
    totalAgents      += agents;
    totalOps         += ops;
    totalTarget      += target;
    totalAgentTarget += agentTarget;
    totalOpTarget    += opTarget;
    totalAudits      += audits;
    totalPayoutUsd   += payoutUsd;
    totalPayoutLcy   += payoutLcy;

    return {
      name:           String(r[cols.NAME]  || '').trim(),
      email:          String(r[cols.EMAIL] || '').toLowerCase().trim(),
      region:         cols.REGION   >= 0 ? String(r[cols.REGION]   || '').trim() : '',
      location:       cols.LOCATION >= 0 ? String(r[cols.LOCATION] || '').trim() : '',
      personalNmv,
      teamNmv,
      totalNmv:       totalNmvRow,
      nmvTarget:      target,
      nmvPct:         pct,
      agents,
      agentTarget,
      ops,
      opTarget,
      audits,
      payoutUsd,
      payoutLcy,
      agentTeamNmvUsd
    };
  });

  leaderboard.sort(function(a, b) { return b.totalNmv - a.totalNmv; });

  var count   = leaderboard.length;
  var teamPct = totalTarget > 0 ? Math.round((totalNmv / totalTarget) * 100) : 0;

  return {
    month,
    totals: {
      personalNmv:  totalPersonalNmv,
      teamNmv:      totalTeamNmv,
      nmv:          totalNmv,
      nmvTarget:    totalTarget,
      nmvPct:       teamPct,
      agents:       totalAgents,
      agentTarget:  totalAgentTarget,
      ops:          totalOps,
      opTarget:     totalOpTarget,
      audits:       totalAudits,
      payoutUsd:    totalPayoutUsd,
      payoutLcy:    totalPayoutLcy,
      rccCount:     count,
      avgNmv:       count > 0 ? Math.round(totalNmv / count) : 0
    },
    leaderboard
  };
}

async function getCheckins(sheets, sid, filterEmail) {
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
      var rowEmail = row[2] ? String(row[2]).toLowerCase().trim() : '';
      if (filterEmail && rowEmail !== filterEmail) return;

      var ts = row[1] ? new Date(row[1]) : null;
      if (!ts || isNaN(ts.getTime())) return;

      if (ts >= startOfMonth) thisMonth++;
      if (ts >= startOfWeek)  thisWeek++;
      if (ts >= startOfDay)   today++;

      var loc = row[7] ? String(row[7]) : '';
      if (loc) locations.add(loc.substring(0, 20));

      recent.push({
        id:           row[0]  || '',
        timestamp:    ts.toISOString(),
        rccEmail:     rowEmail,
        rccName:      row[3]  ? String(row[3])  : '',
        location:     loc,
        photo:        row[5]  ? String(row[5])  : '',
        photoUrl:     row[6]  ? String(row[6])  : '',
        activityType: row[11] ? String(row[11]) : '',
        notes:        row[9]  ? String(row[9])  : ''
      });
    });

    recent.sort(function(a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });

    var limit = filterEmail ? 10 : 50;
    return {
      today,
      thisWeek,
      thisMonth,
      locations: locations.size,
      recent:   recent.slice(0, limit)
    };

  } catch (e) {
    console.error('getCheckins error:', e.message);
    return emptyCheckins();
  }
}

function emptyCheckins() {
  return { today: 0, thisWeek: 0, thisMonth: 0, locations: 0, recent: [] };
}

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
    .map(function(p) { return p[0].toUpperCase(); })
    .join('');
}
