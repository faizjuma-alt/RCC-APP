/**
 * Vercel Serverless Function: /api/auth
 *
 * Actions:
 *   POST ?action=request  — whitelist check + send magic link email via Gmail SMTP
 *   GET  ?action=verify   — verify magic token, issue 30-day session JWT
 *   GET  ?action=me       — validate session JWT, return user info
 *
 * Required env vars:
 *   GOOGLE_SA_KEY   — base64 service account JSON (to read whitelist from sheet)
 *   SPREADSHEET_ID  — Google Sheet ID
 *   JWT_SECRET      — random 32+ char secret for signing tokens
 *   GMAIL_USER      — Gmail address to send from (e.g. jafarfaiz8@gmail.com)
 *   GMAIL_PASS      — Gmail App Password (16-char, no spaces)
 *                     Generate at: myaccount.google.com → Security → 2-Step Verification → App passwords
 *
 * Optional env vars:
 *   APP_URL         — base URL for magic links (default: https://rcc-app-one.vercel.app)
 */

'use strict';

const { google }  = require('googleapis');
const jwt         = require('jsonwebtoken');
const nodemailer  = require('nodemailer');

const DATA_TAB = 'Current month';
const APP_URL  = process.env.APP_URL  || 'https://rcc-app-one.vercel.app';

// ── Helpers ───────────────────────────────────────────────────────────────────

function getSheetAuth() {
  const key = Buffer.from(process.env.GOOGLE_SA_KEY, 'base64').toString('utf8');
  return new google.auth.GoogleAuth({
    credentials: JSON.parse(key),
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
  });
}

function signToken(payload, expiresIn) {
  return jwt.sign(payload, process.env.JWT_SECRET, { expiresIn });
}

function verifyToken(token) {
  return jwt.verify(token, process.env.JWT_SECRET);
}

function extractBearer(req) {
  const auth = req.headers.authorization || '';
  return auth.startsWith('Bearer ') ? auth.slice(7).trim() : null;
}

// ── Sheet whitelist lookup ────────────────────────────────────────────────────

async function findUserInSheet(email) {
  const auth   = getSheetAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const cleanEmail = email.toLowerCase().trim();

  // 1. Check main RCC data tab first
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId:     process.env.SPREADSHEET_ID,
    range:             `'${DATA_TAB}'!A1:AZ`,
    valueRenderOption: 'UNFORMATTED_VALUE'
  });

  const allRows = res.data.values || [];

  let headerIdx = -1;
  let emailCol  = -1;
  let nameCol   = -1;

  for (let i = 0; i < Math.min(allRows.length, 6); i++) {
    const row = allRows[i].map(c => String(c || '').toLowerCase().trim());
    const eIdx = row.findIndex(c => c.includes('email'));
    const nIdx = row.findIndex(c => c.includes('name'));
    if (eIdx >= 0 && nIdx >= 0) {
      headerIdx = i;
      emailCol  = eIdx;
      nameCol   = nIdx;
      break;
    }
  }

  if (headerIdx >= 0) {
    const dataRows = allRows.slice(headerIdx + 1);
    const userRow  = dataRows.find(row =>
      String(row[emailCol] || '').toLowerCase().trim() === cleanEmail
    );
    if (userRow) {
      return {
        email: cleanEmail,
        name:  String(userRow[nameCol] || '').trim() || formatName(email)
      };
    }
  }

  // 2. Not found as RCC — check Managers tab so managers can log in too
  try {
    const mgrRes = await sheets.spreadsheets.values.get({
      spreadsheetId:     process.env.SPREADSHEET_ID,
      range:             `'Managers'!A1:B`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    const mgrRows = mgrRes.data.values || [];
    // Column A = email, Column B = name (optional header row skipped automatically)
    const mgrRow = mgrRows.find(row =>
      row[0] && String(row[0]).toLowerCase().trim() === cleanEmail
    );
    if (mgrRow) {
      return {
        email: cleanEmail,
        name:  mgrRow[1] ? String(mgrRow[1]).trim() : formatName(email)
      };
    }
  } catch (e) {
    // Managers tab may not exist yet — that's fine
    console.warn('Managers tab not found during auth lookup:', e.message);
  }

  return null; // email not found in either tab → blocked
}

function formatName(email) {
  return email.split('@')[0].split(/[._]/)
    .map(s => s.charAt(0).toUpperCase() + s.slice(1).toLowerCase())
    .join(' ');
}

// ── Email template ────────────────────────────────────────────────────────────

function buildMagicLinkEmail(name, magicUrl) {
  const firstName = name.split(' ')[0];
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Your RCC Dashboard login link</title>
</head>
<body style="margin:0;padding:0;background:#F9FAFB;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#F9FAFB;padding:40px 20px;">
    <tr>
      <td align="center">
        <table width="500" cellpadding="0" cellspacing="0" style="background:white;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">
          <!-- Header -->
          <tr>
            <td style="background:#F68B1E;padding:28px 32px;">
              <p style="margin:0;color:white;font-size:20px;font-weight:700;letter-spacing:0.3px;">RCC Performance Dashboard</p>
              <p style="margin:4px 0 0;color:rgba(255,255,255,0.85);font-size:13px;">Field Intelligence · Jumia</p>
            </td>
          </tr>
          <!-- Body -->
          <tr>
            <td style="padding:36px 32px;">
              <p style="margin:0 0 8px;font-size:22px;font-weight:700;color:#1F2937;">Hi ${firstName} 👋</p>
              <p style="margin:0 0 28px;font-size:15px;color:#6B7280;line-height:1.6;">
                Click the button below to sign in to your dashboard. This link expires in <strong>1 hour</strong> and can only be used once.
              </p>
              <table cellpadding="0" cellspacing="0">
                <tr>
                  <td style="background:#F68B1E;border-radius:8px;">
                    <a href="${magicUrl}" style="display:inline-block;padding:14px 32px;color:white;font-size:15px;font-weight:600;text-decoration:none;letter-spacing:0.2px;">Sign in to Dashboard →</a>
                  </td>
                </tr>
              </table>
              <p style="margin:28px 0 0;font-size:13px;color:#9CA3AF;line-height:1.6;">
                If you didn't request this, you can safely ignore this email.<br>
                This link will expire automatically.
              </p>
            </td>
          </tr>
          <!-- Footer -->
          <tr>
            <td style="padding:20px 32px;border-top:1px solid #F3F4F6;">
              <p style="margin:0;font-size:12px;color:#9CA3AF;">
                RCC Performance Dashboard · Built by Yankee Solutions © 2026
              </p>
            </td>
          </tr>
        </table>
        <!-- Fallback link -->
        <p style="margin:20px 0 0;font-size:12px;color:#9CA3AF;">
          Link not working? Copy and paste this URL into your browser:<br>
          <a href="${magicUrl}" style="color:#F68B1E;word-break:break-all;">${magicUrl}</a>
        </p>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

// ── Handler ───────────────────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (!process.env.JWT_SECRET)
    return res.status(500).json({ success: false, error: 'JWT_SECRET not configured' });

  const action = req.query.action || '';

  // ── POST /api/auth?action=request ─────────────────────────────────────────
  if (action === 'request') {
    if (req.method !== 'POST')
      return res.status(405).json({ success: false, error: 'Method not allowed' });

    let body;
    try {
      body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    } catch {
      return res.status(400).json({ success: false, error: 'Invalid JSON' });
    }

    const email = (body && body.email || '').toLowerCase().trim();
    if (!email || !email.includes('@'))
      return res.status(400).json({ success: false, error: 'Valid email required' });

    // Whitelist check
    let user;
    try {
      user = await findUserInSheet(email);
    } catch (err) {
      console.error('Sheet lookup error:', err.message);
      return res.status(500).json({ success: false, error: 'Could not verify email. Try again.' });
    }

    if (!user) {
      return res.status(403).json({
        success: false,
        error:   'Email not registered. Contact your manager to be added.'
      });
    }

    // Generate 1-hour magic token
    const magicToken = signToken({ email: user.email, name: user.name, type: 'magic' }, '1h');
    const magicUrl   = `${APP_URL}/?token=${magicToken}`;

    // Send email via Gmail SMTP (App Password)
    if (!process.env.GMAIL_USER || !process.env.GMAIL_PASS) {
      // Dev mode: log link to console, still return success
      console.log('\n--- MAGIC LINK (no Gmail credentials set) ---');
      console.log(magicUrl);
      console.log('---------------------------------------------\n');
      return res.status(200).json({ success: true, message: 'Check your email for the login link.' });
    }

    try {
      const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
          user: process.env.GMAIL_USER,
          pass: process.env.GMAIL_PASS   // 16-char App Password, no spaces
        }
      });
      await transporter.sendMail({
        from:    `RCC Dashboard <${process.env.GMAIL_USER}>`,
        to:      user.email,
        subject: 'Your login link for RCC Dashboard',
        html:    buildMagicLinkEmail(user.name, magicUrl)
      });
    } catch (err) {
      console.error('Gmail send error:', err.message);
      return res.status(500).json({ success: false, error: 'Could not send email. Try again.' });
    }

    return res.status(200).json({ success: true, message: 'Check your email for the login link.' });
  }

  // ── GET /api/auth?action=verify&token=xxx ─────────────────────────────────
  if (action === 'verify') {
    if (req.method !== 'GET')
      return res.status(405).json({ success: false, error: 'Method not allowed' });

    const token = req.query.token || '';
    if (!token)
      return res.status(400).json({ success: false, error: 'Token required' });

    let payload;
    try {
      payload = verifyToken(token);
    } catch (err) {
      const msg = err.name === 'TokenExpiredError'
        ? 'Link has expired. Please request a new one.'
        : 'Invalid link. Please request a new one.';
      return res.status(401).json({ success: false, error: msg });
    }

    if (payload.type !== 'magic')
      return res.status(401).json({ success: false, error: 'Invalid token type' });

    // Issue 30-day session token
    const sessionToken = signToken(
      { email: payload.email, name: payload.name, type: 'session' },
      '30d'
    );

    return res.status(200).json({
      success:      true,
      sessionToken,
      user: { email: payload.email, name: payload.name }
    });
  }

  // ── GET /api/auth?action=me ───────────────────────────────────────────────
  if (action === 'me') {
    if (req.method !== 'GET')
      return res.status(405).json({ success: false, error: 'Method not allowed' });

    const token = extractBearer(req);
    if (!token)
      return res.status(401).json({ success: false, error: 'No token provided' });

    let payload;
    try {
      payload = verifyToken(token);
    } catch {
      return res.status(401).json({ success: false, error: 'Invalid or expired session' });
    }

    if (payload.type !== 'session')
      return res.status(401).json({ success: false, error: 'Invalid token type' });

    return res.status(200).json({
      success: true,
      user:    { email: payload.email, name: payload.name }
    });
  }

  return res.status(400).json({ success: false, error: `Unknown action: ${action}` });
};
