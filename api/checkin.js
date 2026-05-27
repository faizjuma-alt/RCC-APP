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
    const payload = typeof req.body === 'string' ? JSON.parse(req.body) :