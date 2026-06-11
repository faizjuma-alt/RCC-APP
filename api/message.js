/**
 * Vercel Serverless Function: /api/message
 *
 * Sheet tabs (in SPREADSHEET_ID):
 *   "Comms"   — row 1: headers (Name | Link | Image | Price | Tips | Category)
 *               row 2+: product listings (manager-curated featured products)
 *   "Pushers" — auto-created: Timestamp | Email | Name | Product
 *   "Current month" — existing RCC data for top performers
 *
 * Actions:
 *   GET  ?action=inbox                      — products + top performers + push state
 *   POST { action:"push",   product:"..." } — mark RCC as pushing this product
 *   POST { action:"unpush", product:"..." } — unmark
 *
 * Required env: GOOGLE_SA_KEY, SPREADSHEET_ID, JWT_SECRET
 */

const { google } = require('googleapis');
const jwt        = require('jsonwebtoken');

const COMMS_TAB    = 'Comms';
const PUSHERS_TAB  = 'Pushers';
const DATA_TAB     = 'Current month';
const MANAGERS_TAB = 'Managers';

function getReadAuth() {
  const key = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(key),
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
  });
}

function getWriteAuth() {
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
  } catch {
    res.status(401).json({ success: false, error: 'Session expired. Please sign in again.' });
    return null;
  }
}

async function getManagerInfo(sheets, sid, email) {
  try {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: sid,
      range: `'${MANAGERS_TAB}'!A:C`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const rows  = resp.data.values || [];
    const found = rows.find(r =>
      String(r[0] || '').toLowerCase().trim() === email.toLowerCase().trim()
    );
    if (found) return { isManager: true, market: String(found[2] || '').trim().toUpperCase() };
  } catch {}
  return { isManager: false, market: '' };
}

async function getTopPerformers(sheets, sid, market) {
  try {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: sid,
      range: `'${DATA_TAB}'!A1:AZ`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const allRows = resp.data.values || [];

    // Locate header row
    let headerIdx = -1, cols = {};
    for (let i = 0; i < Math.min(allRows.length, 6); i++) {
      const row = allRows[i].map(c => String(c || '').toLowerCase().trim());
      if (row.some(c => c.includes('email')) && row.some(c => c.includes('name'))) {
        headerIdx = i;
        row.forEach((h, idx) => {
          if (h.includes('email') && cols.EMAIL === undefined)                      cols.EMAIL  = idx;
          else if ((h.includes('name') || h === 'rcc') && cols.NAME === undefined) cols.NAME   = idx;
          else if (h.includes('nmv') && !h.includes('team') &&
                   !h.includes('target') && cols.NMV === undefined)                cols.NMV    = idx;
          else if ((h.includes('market') || h.includes('region') ||
                    h.includes('country')) && cols.REGION === undefined)           cols.REGION = idx;
        });
        break;
      }
    }
    if (headerIdx < 0) return [];

    let rows = allRows.slice(headerIdx + 1).filter(r => r && r[cols.EMAIL]);

    if (market && market !== 'ALL') {
      rows = rows.filter(r =>
        String(r[cols.REGION] || '').trim().toUpperCase() === market
      );
    }

    rows.sort((a, b) => {
      const nmvA = parseFloat(String(a[cols.NMV] || '0').replace(/[^0-9.]/g, '')) || 0;
      const nmvB = parseFloat(String(b[cols.NMV] || '0').replace(/[^0-9.]/g, '')) || 0;
      return nmvB - nmvA;
    });

    return rows.slice(0, 5).map((r, i) => ({
      rank: i + 1,
      name: String(r[cols.NAME] || '').trim(),
      nmv:  parseFloat(String(r[cols.NMV] || '0').replace(/[^0-9.]/g, '')) || 0
    }));
  } catch {
    return [];
  }
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const authUser = requireAuth(req, res);
  if (!authUser) return;

  if (!process.env.GOOGLE_SA_KEY)  return res.status(500).json({ success: false, error: 'GOOGLE_SA_KEY not set' });
  if (!process.env.SPREADSHEET_ID) return res.status(500).json({ success: false, error: 'SPREADSHEET_ID not set' });

  const sid       = process.env.SPREADSHEET_ID;
  const userEmail = authUser.email.toLowerCase().trim();
  const userName  = authUser.name || userEmail.split('@')[0];

  try {
    // ── GET: products + top performers ──────────────────────────────────────────
    if (req.method === 'GET') {
      const readAuth = getReadAuth();
      const sheets   = google.sheets({ version: 'v4', auth: readAuth });

      const managerInfo   = await getManagerInfo(sheets, sid, userEmail);
      const isManager     = managerInfo.isManager;
      const market        = managerInfo.market;
      const topPerformers = await getTopPerformers(sheets, sid, market);

      // ── Read featured products from "Comms" tab ──────────────────────────────
      let products = [];
      try {
        const commsResp = await sheets.spreadsheets.values.get({
          spreadsheetId: sid,
          range: `'${COMMS_TAB}'!A:F`,
          valueRenderOption: 'UNFORMATTED_VALUE'
        });
        const commsRows = commsResp.data.values || [];
        // Row 1 = header (Name|Link|Image|Price|Tips|Category), row 2+ = data
        products = commsRows.slice(1)
          .filter(r => r && String(r[0] || '').trim())
          .map((r, i) => ({
            id:        i,
            name:      String(r[0] || '').trim(),
            link:      String(r[1] || '').trim(),
            imageUrl:  String(r[2] || '').trim(),
            price:     String(r[3] || '').trim(),
            tips:      String(r[4] || '').trim(),
            category:  String(r[5] || '').trim(),
            pushers:   [],
            pushCount: 0,
            isPushing: false
          }));
      } catch {
        // "Comms" tab doesn't exist yet — return empty list
      }

      // ── Read pushers to annotate each product ────────────────────────────────
      try {
        const pushersResp = await sheets.spreadsheets.values.get({
          spreadsheetId: sid,
          range: `'${PUSHERS_TAB}'!A:D`,
          valueRenderOption: 'UNFORMATTED_VALUE'
        });
        const pushersRows = (pushersResp.data.values || []).slice(1); // skip header

        // Build map: productName → [{ email, name }] (deduplicated per person)
        const pushMap = {};
        pushersRows.forEach(r => {
          const pEmail   = String(r[1] || '').toLowerCase().trim();
          const pName    = String(r[2] || '').trim();
          const pProduct = String(r[3] || '').trim();
          if (!pProduct || !pEmail) return;
          if (!pushMap[pProduct]) pushMap[pProduct] = [];
          if (!pushMap[pProduct].find(p => p.email === pEmail)) {
            pushMap[pProduct].push({ email: pEmail, name: pName });
          }
        });

        products.forEach(p => {
          const pushers  = pushMap[p.name] || [];
          p.pushers      = pushers.map(x => x.name);
          p.pushCount    = pushers.length;
          p.isPushing    = pushers.some(x => x.email === userEmail);
        });
      } catch {
        // "Pushers" tab doesn't exist yet — that's fine
      }

      return res.json({
        success:      true,
        role:         isManager ? 'manager' : 'rcc',
        products,
        topPerformers
      });
    }

    // ── POST: push / unpush ──────────────────────────────────────────────────────
    if (req.method === 'POST') {
      let body = req.body;
      if (typeof body === 'string') {
        try { body = JSON.parse(body); } catch { body = {}; }
      }

      const action      = body.action;
      const productName = String(body.product || '').trim();
      if (!productName) return res.status(400).json({ success: false, error: 'Product name required' });

      const writeAuth = getWriteAuth();
      const sheets    = google.sheets({ version: 'v4', auth: writeAuth });

      // Ensure Pushers sheet exists
      let pushersRows    = [];
      let pushersSheetId = null;
      try {
        const resp = await sheets.spreadsheets.values.get({
          spreadsheetId: sid,
          range: `'${PUSHERS_TAB}'!A:D`,
          valueRenderOption: 'UNFORMATTED_VALUE'
        });
        pushersRows = resp.data.values || [];
        const meta  = await sheets.spreadsheets.get({ spreadsheetId: sid });
        const found = meta.data.sheets.find(s => s.properties.title === PUSHERS_TAB);
        if (found) pushersSheetId = found.properties.sheetId;
      } catch {
        // Create the Pushers tab
        try {
          const addResp = await sheets.spreadsheets.batchUpdate({
            spreadsheetId: sid,
            requestBody: { requests: [{ addSheet: { properties: { title: PUSHERS_TAB } } }] }
          });
          pushersSheetId = addResp.data.replies[0].addSheet.properties.sheetId;
          await sheets.spreadsheets.values.update({
            spreadsheetId:    sid,
            range:            `'${PUSHERS_TAB}'!A1:D1`,
            valueInputOption: 'RAW',
            requestBody:      { values: [['Timestamp', 'Email', 'Name', 'Product']] }
          });
          pushersRows = [['Timestamp', 'Email', 'Name', 'Product']];
        } catch (e) {
          return res.status(500).json({ success: false, error: 'Could not create Pushers sheet: ' + e.message });
        }
      }

      if (action === 'push') {
        // Idempotent — check if already marked
        const already = pushersRows.slice(1).some(r =>
          String(r[1] || '').toLowerCase().trim() === userEmail &&
          String(r[3] || '').trim() === productName
        );
        if (already) return res.json({ success: true });

        await sheets.spreadsheets.values.append({
          spreadsheetId:    sid,
          range:            `'${PUSHERS_TAB}'!A:D`,
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS',
          requestBody: {
            values: [[new Date().toISOString(), userEmail, userName, productName]]
          }
        });
        return res.json({ success: true });
      }

      if (action === 'unpush') {
        const dataRows = pushersRows.slice(1);
        const rowIdx   = dataRows.findIndex(r =>
          String(r[1] || '').toLowerCase().trim() === userEmail &&
          String(r[3] || '').trim() === productName
        );
        if (rowIdx < 0) return res.json({ success: true }); // already not pushing

        // Sheet row is 0-based: rowIdx + 1 (skip header)
        const sheetRowZeroBased = rowIdx + 1;

        if (pushersSheetId === null) {
          const meta  = await sheets.spreadsheets.get({ spreadsheetId: sid });
          const found = meta.data.sheets.find(s => s.properties.title === PUSHERS_TAB);
          if (found) pushersSheetId = found.properties.sheetId;
        }

        if (pushersSheetId !== null) {
          await sheets.spreadsheets.batchUpdate({
            spreadsheetId: sid,
            requestBody: {
              requests: [{
                deleteDimension: {
                  range: {
                    sheetId:    pushersSheetId,
                    dimension:  'ROWS',
                    startIndex: sheetRowZeroBased,
                    endIndex:   sheetRowZeroBased + 1
                  }
                }
              }]
            }
          });
        }
        return res.json({ success: true });
      }

      return res.status(400).json({ success: false, error: 'Unknown action: ' + action });
    }

    return res.status(405).json({ success: false, error: 'Method not allowed' });

  } catch (err) {
    console.error('[message]', err);
    return res.status(500).json({ success: false, error: err.message });
  }
};
