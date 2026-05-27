/**
 * Vercel Serverless Function: /api/data
 * Reads from the Jumia Google Sheet using a service account.
 * No Apps Script — no Workspace restrictions.
 *
 * Required environment variables in Vercel:
 *   GOOGLE_SA_KEY    — base64-encoded service account JSON key
 *   SPREADSHEET_ID   — Google Sheet ID (from the sheet URL)
 *
 * Supported query params:
 *   ?action=dashboard&email=...&name=...
 *   ?action=teamlist
 */

const { google } = require('googleapis');

// ── Sheet names (must match your actual tab names) ──────────────────────────
const SHEETS = {
  TEAM_NMV:      'Team + RCC NMV',
  ORDER_POINTS:  'Order point',
  AGENTS:        'New active agents',
  FIELD_CHECKINS: 'RCC_Field_Checkins'
};

// ── Targets (edit here if they change) ──────────────────────────────────────
const PERSONAL_TARGET = 5700;
const TEAM_TARGET     = 10000;

// ── Auth helper ──────────────────────────────────────────────────────────────
function getAuth() {
  const keyJson = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  const credentials = JSON.parse(keyJson);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets.readonly'
    ]
  });
}

// ── Handler ────────────────────────────────────────────