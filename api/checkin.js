/**
 * Vercel Serverless Function: /api/checkin
 * Appends a field check-in row to the sheet and optionally uploads
 * the photo to Google Drive — all via service account.
 *
 * Required environment variables in Vercel:
 *   GOOGLE_SA_KEY    — base64-encoded service account JSON key
 *   SPREADSHEET_ID   — Google Sheet ID
 * Optional:
 *   CHECKINS_SPREADSHEET_ID — if check-ins are in a different sheet
 *   DRIVE_FOLDER_ID         — Drive folder ID for photo uploads
 *
 * Body (JSON): {
 *   userEmail, userName,
 *   activityType, location, notes,
 *   photoBase64?, photoName?
 * }
 */

'use strict';

const { google } = require('googleapis');
const { Readable } = require('stream');

const CHECKINS_SHEET = 'RCC_Field_Checkins';

// ── IMPORTANT: Increase Vercel body-size limit so large base64 photos don't
//    cause Vercel to reject the request with an HTML error page (the "A server
//    error" the client sees as invalid JSON).
//    This config MUST be a property on the exported handler function.
async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST')
    return res.status(405).json({ success: false, error: 'Method not allowed' });

  if (!process.env.GOOGLE_SA_KEY)
    return res.status(500).json({ success: false, error: 'GOOGLE_SA_KEY not set' });
  if (!process.env.SPREADSHEET_ID)
    return res.status(500).json({ success: false, error: 'SPREADSHEET_ID not set' });

  try {
    // ── Parse body ──────────────────────────────────────────────────────────
    let payload;
    try {
      payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    } catch (parseErr) {
      return res.status(400).json({ success: false, error: 'Invalid JSON body: ' + parseErr.message });
    }
    if (!payload || typeof payload !== 'object') {
      return res.status(400).json({ success: false, error: 'Empty or invalid request body' });
    }

    const userEmail    = payload.userEmail    || 'unknown@rcc';
    const userName     = (payload.userName && payload.userName.trim()) || getUserDisplayName(userEmail);
    const activityType = payload.activityType || '';
    const location     = payload.location     || '';
    const notes        = payload.notes        || '';
    const photoBase64  = payload.photoBase64  || '';
    const photoName    = payload.photoName    || 'checkin.jpg';

    const id  = generateId();
    const now = new Date();
    const isoTs = now.toISOString();
    const sid   = process.env.CHECKINS_SPREADSHEET_ID || process.env.SPREADSHEET_ID;

    // ── Optional: upload photo to Drive (non-fatal if it fails) ─────────────
    let photoUrl        = '';
    let photoPreviewUrl = '';

    if (photoBase64 && photoBase64.length > 100) {
      try {
        const driveAuth = getAuth([
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/spreadsheets'
        ]);
        const drive = google.drive({ version: 'v3', auth: driveAuth });

        // Strip data-URI prefix (data:image/jpeg;base64,...)
        const base64Data = photoBase64.replace(/^data:[^;]+;base64,/, '');
        const mimeType   = photoBase64.startsWith('data:image/png') ? 'image/png' : 'image/jpeg';
        const fileName   = id + '_' + now.getTime() + '.jpg';

        const fileBuffer = Buffer.from(base64Data, 'base64');
        const stream     = Readable.from(fileBuffer);

        const folderId = process.env.DRIVE_FOLDER_ID || null;
        const driveRes = await drive.files.create({
          requestBody: {
            name:     fileName,
            mimeType: mimeType,
            ...(folderId ? { parents: [folderId] } : {})
          },
          media:  { mimeType: mimeType, body: stream },
          fields: 'id,webViewLink'
        });

        const fileId = driveRes.data.id || '';
        photoUrl        = driveRes.data.webViewLink || '';
        photoPreviewUrl = fileId ? 'https://drive.google.com/uc?id=' + fileId : '';

        // Make the file readable by anyone with the link (non-fatal)
        if (fileId) {
          await drive.permissions.create({
            fileId:      fileId,
            requestBody: { role: 'reader', type: 'anyone' }
          }).catch(permErr => console.warn('Drive permission set failed:', permErr.message));
        }
      } catch (driveErr) {
        // Non-fatal — save the check-in row without a photo link
        console.error('Drive upload failed (continuing without photo):', driveErr.message);
        photoUrl        = '';
        photoPreviewUrl = '';
      }
    }

    // ── Append row to RCC_Field_Checkins ─────────────────────────────────────
    // Column layout (A–M, indices 0–12) matches what api/data.js reads back:
    //   row[0]=id  row[1]=timestamp  row[2]=email  row[3]=name
    //   row[4]=timeStr  row[5]=''  row[6]=''
    //   row[7]=location  row[8]=photoUrl  row[9]=photoPreviewUrl
    //   row[10]=''  row[11]=activityType  row[12]=notes
    const timeStr = now.toTimeString().slice(0, 8); // HH:MM:SS

    const newRow = [
      id,             // A (0)  — unique ID
      isoTs,          // B (1)  — full ISO timestamp
      userEmail,      // C (2)  — RCC email
      userName,       // D (3)  — RCC display name
      timeStr,        // E (4)  — time of day
      '',             // F (5)
      '',             // G (6)
  