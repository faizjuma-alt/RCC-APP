/**
 * Vercel Serverless Function: /api/message
 *
 * Sheet tab: "Messages" (auto-created on first use)
 * Columns: A=ID  B=Timestamp  C=FromEmail  D=FromName
 *          E=Market  F=Type  G=ParentID  H=Subject  I=Body
 *
 * Actions:
 *   GET  ?action=inbox     — inbox for current user's market
 *   POST ?action=broadcast — manager posts announcement to their market
 *   POST ?action=reply     — any user replies to a broadcast
 *
 * Required env: GOOGLE_SA_KEY, SPREADSHEET_ID, JWT_SECRET
 */

'use strict';

const { google } = require('googleapis');
const jwt        = require('jsonwebtoken');

const MESSAGES_TAB = 'Messages';
const DATA_TAB     = 'Current month';
const MANAGERS_TAB = 'Managers';

// ── Auth helpers ──────────────────────────────────────────────────────────────

function getAuth() {
  const key = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(key),
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
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

// ── Manager info ──────────────────────────────────────────────────────────────

async function getManagerInfo(sheets, sid, email) {
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: sid, range: `'${MANAGERS_TAB}'!A1:C`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const row = (r.data.values || []).find(r =>
      r[0] && String(r[0]).toLowerCase().trim() === email
    );
    if (!row) return { isManager: false, market: '' };
    return { isManager: true, market: row[2] ? String(row[2]).trim().toUpperCase() : 'ALL' };
  } catch (e) {
    return { isManager: false, market: '' };
  }
}

// ── RCC market lookup ─────────────────────────────────────────────────────────

async function getRccMarket(sheets, sid, email) {
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: sid, range: `'${DATA_TAB}'!A1:AZ`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const allRows = r.data.values || [];
    let headerIdx = -1, emailCol = -1, marketCol = -1;
    for (let i = 0; i < Math.min(allRows.length, 6); i++) {
      const row = allRows[i].map(c => String(c || '').toLowerCase().trim());
      if (row.some(c => c.includes('email'))) {
        headerIdx = i;
        emailCol  = row.findIndex(c => c.includes('email'));
        marketCol = row.findIndex(c => c === 'markets' || c === 'market' || c.includes('markets'));
        break;
      }
    }
    if (headerIdx < 0) return '';
    const userRow = allRows.slice(headerIdx + 1).find(r =>
      String(r[emailCol] || '').toLowerCase().trim() === email
    );
    return (userRow && marketCol >= 0)
      ? String(userRow[marketCol] || '').trim().toUpperCase()
      : '';
  } catch (e) {
    return '';
  }
}

// ── Ensure Messages tab exists ────────────────────────────────────────────────

async function ensureMessagesTab(sheets, sid) {
  try {
    const meta   = await sheets.spreadsheets.get({ spreadsheetId: sid });
    const exists = meta.data.sheets.some(s => s.properties.title === MESSAGES_TAB);
    if (!exists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: sid,
        requestBody: { requests: [{ addSheet: { properties: { title: MESSAGES_TAB } } }] }
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId: sid,
        range: `'${MESSAGES_TAB}'!A1:I1`,
        valueInputOption: 'RAW',
        requestBody: { values: [['ID','Timestamp','FromEmail','FromName','Market','Type','ParentID','Subject','Body']] }
      });
    }
  } catch (e) {
    console.warn('ensureMessagesTab:', e.message);
  }
}

// ── Parse a sheet row into a message object ───────────────────────────────────

function parseRow(r) {
  return {
    id:        String(r[0] || ''),
    timestamp: String(r[1] || ''),
    fromEmail: String(r[2] || '').toLowerCase().trim(),
    fromName:  String(r[3] || ''),
    market:    String(r[4] || '').toUpperCase(),
    type:      String(r[5] || ''),
    parentId:  String(r[6] || ''),
    subject:   String(r[7] || ''),
    body:      String(r[8] || '')
  };
}

// ── Handler ───────────────────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const authUser = requireAuth(req, res);
  if (!authUser) return;

  const sid = process.env.SPREADSHEET_ID;
  if (!sid) return res.status(500).json({ success: false, error: 'SPREADSHEET_ID not set' });

  const auth    = getAuth();
  const sheets  = google.sheets({ version: 'v4', auth });
  const action  = (req.query.action || '').toLowerCase();
  const email   = authUser.email.toLowerCase().trim();

  await ensureMessagesTab(sheets, sid);

  // ── GET inbox ───────────────────────────────────────────────────────────────
  if (action === 'inbox' && req.method === 'GET') {
    const mgrInfo = await getManagerInfo(sheets, sid, email);
    const market  = mgrInfo.isManager
      ? mgrInfo.market
      : await getRccMarket(sheets, sid, email);

    const raw = await sheets.spreadsheets.values.get({
      spreadsheetId: sid, range: `'${MESSAGES_TAB}'!A2:I`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });

    const all = (raw.data.values || []).filter(r => r && r[0]).map(parseRow);

    // Broadcasts visible to this user
    const broadcasts = all
      .filter(m => m.type === 'broadcast' && (
        m.market === market ||
        m.market === 'ALL'  ||
        (mgrInfo.isManager && m.fromEmail === email)
      ))
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
      .map(b => ({
        ...b,
        replies: all
          .filter(m => m.type === 'reply' && m.parentId === b.id)
          .sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp))
      }));

    return res.status(200).json({
      success: true,
      role:    mgrInfo.isManager ? 'manager' : 'rcc',
      market,
      messages: broadcasts,
      timestamp: new Date().toISOString()
    });
  }

  // ── POST broadcast (managers only) ─────────────────────────────────────────
  if (action === 'broadcast' && req.method === 'POST') {
    const mgrInfo = await getManagerInfo(sheets, sid, email);
    if (!mgrInfo.isManager)
      return res.status(403).json({ success: false, error: 'Manager access required' });

    let body;
    try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
    catch { return res.status(400).json({ success: false, error: 'Invalid JSON' }); }

    const subject = String(body.subject || '').trim();
    const msgBody = String(body.body    || '').trim();
    if (!subject || !msgBody)
      return res.status(400).json({ success: false, error: 'Subject and body required' });

    const id        = 'MSG-' + Date.now();
    const timestamp = new Date().toISOString();

    await sheets.spreadsheets.values.append({
      spreadsheetId: sid, range: `'${MESSAGES_TAB}'!A:I`,
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[
        id, timestamp, email, authUser.name || '',
        mgrInfo.market, 'broadcast', '', subject, msgBody
      ]] }
    });

    return res.status(200).json({ success: true, id, message: 'Broadcast sent' });
  }

  // ── POST reply (any user) ──────────────────────────────────────────────────
  if (action === 'reply' && req.method === 'POST') {
    let body;
    try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
    catch { return res.status(400).json({ success: false, error: 'Invalid JSON' }); }

    const parentId = String(body.parentId || '').trim();
    const msgBody  = String(body.body     || '').trim();
    if (!parentId || !msgBody)
      return res.status(400).json({ success: false, error: 'parentId and body required' });

    const mgrInfo    = await getManagerInfo(sheets, sid, email);
    const market     = mgrInfo.isManager
      ? mgrInfo.market
      : await getRccMarket(sheets, sid, email);

    const id        = 'RPL-' + Date.now();
    const timestamp = new Date().toISOString();

    await sheets.spreadsheets.values.append({
      spreadsheetId: sid, range: `'${MESSAGES_TAB}'!A:I`,
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[
        id, timestamp, email, authUser.name || '',
        market, 'reply', parentId, '', msgBody
      ]] }
    });

    return res.status(200).json({ success: true, id, message: 'Reply sent' });
  }

  return res.status(400).json({ success: false, error: `Unknown action: ${action}` });
};
