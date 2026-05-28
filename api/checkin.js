/**
 * Vercel Serverless Function: /api/checkin
 *
 * Column layout (matches sheet headers):
 *   A(0)=ID  B(1)=Timestamp  C(2)=Email  D(3)=Name  E(4)=Time
 *   F(5)=PhotoPreview  G(6)=PhotoURL  H(7)=Location
 *   I(8)=''  J(9)=Notes  K(10)=''  L(11)=ActivityType
 *
 * Required env: GOOGLE_SA_KEY, SPREADSHEET_ID
 * Optional env: CHECKINS_SPREADSHEET_ID, DRIVE_FOLDER_ID
 */
'use strict';

const { google }   = require('googleapis');
const { Readable } = require('stream');
const jwt          = require('jsonwebtoken');

const CHECKINS_SHEET = 'RCC_Field_Checkins';

function requireAuth(req, res) {
  const auth  = req.headers.authorization || '';
  const token = auth.startsWith('Bearer ') ? auth.slice(7).trim() : null;
  if (!token) {
    res.status(401).json({ success: false, error: 'Authentication required' });
    return null;
  }
  try {
    const p = jwt.verify(token, process.env.JWT_SECRET);
    if (p.type !== 'session') throw new Error('Invalid token type');
    return p;
  } catch {
    res.status(401).json({ success: false, error: 'Session expired. Please sign in again.' });
    return null;
  }
}

async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')
    return res.status(405).json({ success: false, error: 'Method not allowed' });

  // Auth check — email comes from verified token, not request body
  const authUser = requireAuth(req, res);
  if (!authUser) return;

  if (!process.env.GOOGLE_SA_KEY)
    return res.status(500).json({ success: false, error: 'GOOGLE_SA_KEY not set' });
  if (!process.env.SPREADSHEET_ID)
    return res.status(500).json({ success: false, error: 'SPREADSHEET_ID not set' });

  try {
    let payload;
    try {
      payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    } catch (e) {
      return res.status(400).json({ success: false, error: 'Invalid JSON: ' + e.message });
    }
    if (!payload || typeof payload !== 'object')
      return res.status(400).json({ success: false, error: 'Empty or invalid body' });

    // Use email/name from JWT — not from body (prevents spoofing)
    const userEmail    = authUser.email;
    const userName     = authUser.name || getUserDisplayName(authUser.email);
    const activityType = payload.activityType || '';
    const location     = payload.location     || '';
    const notes        = payload.notes        || '';
    const photoBase64  = payload.photoBase64  || '';

    const id    = generateId();
    const now   = new Date();
    const isoTs = now.toISOString();
    const sid   = process.env.CHECKINS_SPREADSHEET_ID || process.env.SPREADSHEET_ID;

    // Optional Drive upload — non-fatal if it fails
    let photoUrl        = '';
    let photoPreviewUrl = '';

    if (photoBase64 && photoBase64.length > 100) {
      try {
        const driveAuth = getAuth([
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/spreadsheets'
        ]);
        const drive = google.drive({ version: 'v3', auth: driveAuth });

        const base64Data = photoBase64.replace(/^data:[^;]+;base64,/, '');
        const mimeType   = photoBase64.startsWith('data:image/png') ? 'image/png' : 'image/jpeg';
        const fileName   = id + '_' + now.getTime() + '.jpg';
        const fileBuffer = Buffer.from(base64Data, 'base64');
        const stream     = Readable.from(fileBuffer);
        const folderId   = process.env.DRIVE_FOLDER_ID || null;

        const driveRes = await drive.files.create({
          requestBody: {
            name:    fileName,
            mimeType: mimeType,
            parents: folderId ? [folderId] : undefined
          },
          media:  { mimeType: mimeType, body: stream },
          fields: 'id,webViewLink'
        });

        const fileId    = driveRes.data.id || '';
        photoUrl        = driveRes.data.webViewLink || '';
        photoPreviewUrl = fileId ? 'https://drive.google.com/uc?id=' + fileId : '';

        if (fileId) {
          await drive.permissions.create({
            fileId:      fileId,
            requestBody: { role: 'reader', type: 'anyone' }
          }).catch(function(e) { console.warn('Drive permission failed:', e.message); });
        }
      } catch (driveErr) {
        console.error('Drive upload failed (continuing without photo):', driveErr.message);
        photoUrl        = '';
        photoPreviewUrl = '';
      }
    }

    const timeStr = now.toTimeString().slice(0, 8);

    // Columns must match sheet headers exactly
    const newRow = [
      id,              // A (0)  — ID
      isoTs,           // B (1)  — Timestamp
      userEmail,       // C (2)  — Email
      userName,        // D (3)  — Name
      timeStr,         // E (4)  — Time
      photoPreviewUrl, // F (5)  — PhotoPreview (drive.google.com/uc?id= for =IMAGE())
      photoUrl,        // G (6)  — PhotoURL (Drive view link)
      location,        // H (7)  — Location / GPS
      '',              // I (8)
      notes,           // J (9)  — Notes
      '',              // K (10)
      activityType     // L (11) — ActivityType
    ];

    const sheetsAuth = getAuth(['https://www.googleapis.com/auth/spreadsheets']);
    const sheets     = google.sheets({ version: 'v4', auth: sheetsAuth });

    await sheets.spreadsheets.values.append({
      spreadsheetId:    sid,
      range:            "'" + CHECKINS_SHEET + "'!A:L",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody:      { values: [newRow] }
    });

    return res.status(200).json({
      success:  true,
      id:       id,
      message:  'Check-in saved successfully',
      photoUrl: photoUrl
    });

  } catch (err) {
    console.error('api/checkin error:', err);
    return res.status(500).json({ success: false, error: err.message || 'Internal server error' });
  }
}

handler.config = {
  api: { bodyParser: { sizeLimit: '10mb' } }
};

module.exports = handler;

function getAuth(scopes) {
  const keyJson     = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  const credentials = JSON.parse(keyJson);
  return new google.auth.GoogleAuth({ credentials, scopes });
}

function generateId() {
  return 'CI-' + Math.random().toString(36).substring(2, 10).toUpperCase();
}

function getUserDisplayName(email) {
  if (!email) return 'RCC User';
  return email.split('@')[0].split(/[._]/)
    .map(function(s) { return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); })
    .join(' ');
}
