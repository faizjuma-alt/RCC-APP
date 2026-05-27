/**
 * Vercel Serverless Function: /api/data
 *
 * Sheet: "Current month" tab
 *   B1  = current month selector (e.g. 202605) -- user switches this in the sheet
 *   Row 2 = headers: Start date | End date | Name | Email | Region | Location |
 *            New active agent recruitment | Recruitment actual | Newly active agents |
 *            RCC NMV(LCY) | RCC Team NMV (LCY) | RCC + team NMV target |
 *            Order Point enrolled | Newly active order point | New active OP target
 *   Row 3+ = one row per RCC (sheet already filters by B1)
 *
 * "Start date" = when the RCC joined, NOT a month column.
 * Required env: GOOGLE_SA_KEY, SPREADSHEET_ID
 */
const { google } = require('googleapis');
const DATA_TAB = 'Current month';
const CHECKINS_TAB = 'RCC_Field_Checkins';

function getAuth() {
  const key = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  return new google.auth.GoogleAuth({ credentials: JSON.parse(key), scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'] });
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'GET') return res.status(405).json({ success: false, error: 'Method not allowed' });
  if (!process.env.GOOGLE_SA_KEY) return res.status(500).json({ success: false, error: 'GOOGLE_SA_KEY not set' });
  if (!process.env.SPREADSHEET_ID) return res.status(500).json({ success: false, error: 'SPREADSHEET_ID not set' });
  try {
    const auth = getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    const sid = process.env.SPREADSHEET_ID;
    const action = req.query.action || 'dashboard';
    const raw = await sheets.spreadsheets.values.get({ spreadsheetId: sid, range: `'${DATA_TAB}'!A1:AZ`, valueRenderOption: 'UNFORMATTED_VALUE' });
    const allRows = raw.data.values || [];
    const selectedMonth = allRows[0] && allRows[0][1] ? String(allRows[0][1]).replace(/[^0-9]/g, '') : '';
    let headerIdx = -1, cols = {};
    for (let i = 0; i < Math.min(allRows.length, 6); i++) {
      const row = allRows[i].map(c => String(c || '').toLowerCase().trim());
      if (row.some(c => c.includes('email')) && row.some(c => c.includes('name'))) { headerIdx = i; cols = mapColumns(row); break; }
    }
    if (headerIdx < 0) return res.status(500).json({ success: false, error: 'Header row not found.' });
    const dataRows = allRows.slice(headerIdx + 1).filter(r => r && r.length > 0 && r[cols.EMAIL]);
    if (action === 'teamlist') return res.status(200).json(buildTeamList(dataRows, cols, selectedMonth));
    const userEmail = (req.query.email || '').toLowerCase().trim();
    const userName = req.query.name || getUserDisplayName(userEmail);
    const userRow = dataRows.find(r => String(r[cols.EMAIL] || '').toLowerCase().trim() === userEmail) || null;
    const kpis = buildKpis(userRow, dataRows, cols);
    const orderPoints = buildOrderPointsSummary(userRow, cols);
    const agents = buildAgentsSummary(userRow, cols);
    const progress = buildProgress(userRow, cols);
    const fieldCheckins = await getFieldCheckinsSummary(sheets, sid, userEmail);
    const resolvedName = (userRow && userRow[cols.NAME]) ? String(userRow[cols.NAME]).trim() : userName;
    return res.status(200).json({ success: true, user: { email: userEmail, name: resolvedName, initials: getInitials(resolvedName) }, month: selectedMonth, kpis, orderPoints, agents, progress, fieldCheckins, timestamp: new Date().toISOString() });
  } catch (err) { return res.status(500).json({ success: false, error: err.message }); }
};

function mapColumns(h) {
  const f = (...kws) => { for (const kw of kws) { const i = h.findIndex(x => x.includes(kw)); if (i >= 0) return i; } return -1; };
  return { NAME: f('name'), EMAIL: f('email'), REGION: f('region'), LOCATION: f('location'),
    PERSONAL_NMV: f('rcc nmv(lcy)', 'rcc nmv (lcy)', 'rcc nmv lcy', 'personal nmv'),
    TEAM_NMV: f('rcc team nmv (lcy)', 'rcc team nmv(lcy)', 'team nmv (lcy)', 'rcc team nmv'),
    NMV_TARGET: f('rcc + team nmv target', 'nmv target'),
    NEW_AGENTS: f('newly active agents', 'new active agents'),
    AGENT_TARGET: f('new active agent recruitment', 'recruitment actual', 'agent recruitment'),
    NEW_OPS: f('newly active order point', 'new active order point', 'newly active order points'),
    OPS_ENROLLED: f('order point enrolled', 'order points enrolled', 'op enrolled'),
    OP_TARGET: f('new active op target', 'op target', 'new active order point target') };
}
function buildKpis(u, all, c) {
  return { personalNmv: u?toNum(u[c.PERSONAL_NMV]):0, personalTarget: Number(process.env.PERSONAL_TARGET||0),
    teamNmv: u?toNum(u[c.TEAM_NMV]):0, teamTarget: ((u&&c.NMV_TARGET>=0)?toNum(u[c.NMV_TARGET]):0)||Number(process.env.TEAM_TARGET||0),
    activeTeam: all.length, teamTarget_members: all.length };
}
function buildProgress(u, c) {
  if (!u) return { nmv:{actual:0,target:0}, agents:{actual:0,target:0}, ops:{actual:0,target:0} };
  return { nmv:{actual:toNum(u[c.PERSONAL_NMV])+toNum(u[c.TEAM_NMV]),target:toNum(u[c.NMV_TARGET])},
    agents:{actual:toNum(u[c.NEW_AGENTS]),target:toNum(u[c.AGENT_TARGET])},
    ops:{actual:toNum(u[c.NEW_OPS]),target:c.OP_TARGET>=0?toNum(u[c.OP_TARGET]):0} };
}
function buildOrderPointsSummary(u,c){return{newActive:u?toNum(u[c.NEW_OPS]):0,lastTwoMonths:u?toNum(u[c.OPS_ENROLLED]):0,list:[]};}
function buildAgentsSummary(u,c){return{newActive:u?toNum(u[c.NEW_AGENTS]):0,lastTwoMonths:u?toNum(u[c.AGENT_TARGET]):0,list:[]};}
function buildTeamList(rows,cols,month){
  // Each row=one RCC, no sub-agent rows. Return empty list.
  return{success:true,month,list:[]};
}

async function getFieldCheckinsSummary(sheets,sid,email){
  try{
    const r=await sheets.spreadsheets.values.get({spreadsheetId:process.env.CHECKINS_SPREADSHEET_ID||sid,range:`'${CHECKINS_TAB}'!A2:M`,valueRenderOption:'UNFORMATTED_VALUE'});
    const rows=r.data.values||[];if(!rows.length)return emptyCheckins();
    const now=new Date(),sDay=new Date(now.getFullYear(),now.getMonth(),now.getDate()),sWeek=new Date(sDay);sWeek.setDate(sDay.getDate()-sDay.getDay());
    const sMon=new Date(now.getFullYear(),now.getMonth(),1);let today=0,week=0,mon=0;const locs=new Set(),rec=[];
    rows.forEach(row=>{const ts=row[1]?new Date(row[1]):null;if(!ts||isNaN(ts))return;if(ts>=sMon)mon++;if(ts>=sWeek)week++;if(ts>=sDay)today++;if(row[7])locs.add(String(row[7]).substring(0,20));rec.push({id:row[0]||'',timestamp:ts.toISOString(),rccEmail:row[2]||'',rccName:row[3]||'',location:row[7]||'',photo:row[8]||'',activityType:row[11]||'',notes:row[12]||''});});
    rec.sort((a,b)=>new Date(b.timestamp)-new Date(a.timestamp));
    return{today,thisWeek:week,thisMonth:mon,locations:locs.size,recent:rec.slice(0,10)};
  }catch(e){return emptyCheckins();}
}
function emptyCheckins(){return{today:0,thisWeek:0,thisMonth:0,locations:0,recent:[]};}
function toNum(v){if(v===null||v===undefined||v==='')return 0;const n=Number(String(v).replace(/[^0-9.-]/g,''));return isNaN(n)?0:n;}
function getUserDisplayName(e){if(!e)return 'RCC User';return e.split('@')[0].split(/[._]/).map(s=>s.charAt(0).toUpperCase()+s.slice(1).toLowerCase()).join(' ');}
function getInitials(s){if(!s)return '??';const n=s.includes('@')?getUserDisplayName(s):s;return n.trim().split(/\s+/).filter(p=>p.length>0).slice(0,2).map(p=>p[0].toUpperCase()).join('');}