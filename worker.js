/**
 * MRXplorer Diagnostic — Cloudflare Worker
 *
 * Routes:
 *   POST /session  — log partial session state (fire-and-forget from frontend)
 *   POST /submit   — write completed submission to Google Sheets
 *   POST /generate — call Anthropic API and return response
 *
 * Required secrets (set via `wrangler secret put` or Cloudflare dashboard):
 *   ANTHROPIC_API_KEY         — Anthropic API key
 *   GOOGLE_SCRIPT_URL         — Google Apps Script web app URL (see README)
 *
 * The Google Apps Script handles all Sheets writes. To set it up:
 *   1. Open your Google Sheet
 *   2. Extensions → Apps Script
 *   3. Paste the doPost() function from the bottom of this file into Code.gs
 *   4. Deploy → New deployment → Web app
 *      Execute as: Me | Who has access: Anyone
 *   5. Copy the web app URL and save it as the GOOGLE_SCRIPT_URL secret
 */

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

const DEFAULT_MODEL = 'claude-haiku-4-5-20251001';

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: CORS_HEADERS });
    }

    if (request.method !== 'POST') {
      return new Response('Method not allowed', { status: 405 });
    }

    const url = new URL(request.url);

    try {
      if (url.pathname === '/generate') {
        return await handleGenerate(request, env);
      }
      if (url.pathname === '/submit') {
        return await handleSubmit(request, env);
      }
      if (url.pathname === '/session') {
        return await handleSession(request, env);
      }
      return new Response('Not found', { status: 404 });
    } catch (err) {
      return json({ error: err.message }, 500);
    }
  },
};

// ── /generate ────────────────────────────────────────────────────────────────

async function handleGenerate(request, env) {
  const body = await request.json();
  const { system, messages } = body;

  // Use model from request body; fall back to a known-good current default.
  const model = body.model || DEFAULT_MODEL;

  const anthropicRes = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model,
      max_tokens: 4096,
      system,
      messages,
    }),
  });

  if (!anthropicRes.ok) {
    const errText = await anthropicRes.text();
    return json({ error: `Anthropic error ${anthropicRes.status}: ${errText}` }, 500);
  }

  const data = await anthropicRes.json();
  return json(data);
}

// ── /submit ──────────────────────────────────────────────────────────────────

async function handleSubmit(request, env) {
  const body = await request.json();

  if (env.GOOGLE_SCRIPT_URL) {
    // Forward to Apps Script; don't await — let the frontend continue.
    fetch(env.GOOGLE_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        type: 'submit',
        timestamp: new Date().toISOString(),
        ...body,
      }),
    }).catch(() => {});
  }

  return json({ ok: true });
}

// ── /session ─────────────────────────────────────────────────────────────────

async function handleSession(request, env) {
  const body = await request.json();

  if (env.GOOGLE_SCRIPT_URL) {
    fetch(env.GOOGLE_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        type: 'session',
        timestamp: new Date().toISOString(),
        ...body,
      }),
    }).catch(() => {});
  }

  return json({ ok: true });
}

// ── helpers ──────────────────────────────────────────────────────────────────

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' },
  });
}

/*
───────────────────────────────────────────────────────────────────────────────
GOOGLE APPS SCRIPT  (paste into Code.gs, deploy as web app)
───────────────────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.type === 'submit') {
      var sheet = ss.getSheetByName('Submissions') || ss.insertSheet('Submissions');
      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          'Timestamp', 'Session ID', 'Name', 'Email',
          'Workflow Count', 'Workflows (JSON)', 'AI Answers (JSON)',
          'Verdicts (JSON)', 'Summary'
        ]);
      }
      sheet.appendRow([
        data.timestamp,
        data.sessionId || '',
        data.name || '',
        data.email || '',
        data.workflowCount || '',
        JSON.stringify(data.workflows || []),
        JSON.stringify(data.aiAnswers || {}),
        JSON.stringify(data.verdicts || []),
        data.summary || ''
      ]);
    }

    if (data.type === 'session') {
      var log = ss.getSheetByName('Sessions') || ss.insertSheet('Sessions');
      if (log.getLastRow() === 0) {
        log.appendRow(['Timestamp', 'Session ID', 'Screen', 'Workflow Count', 'Last Question']);
      }
      log.appendRow([
        data.timestamp,
        data.sessionId || '',
        data.currentScreen || '',
        data.workflowCount || '',
        data.lastQuestion || ''
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

───────────────────────────────────────────────────────────────────────────────
*/
