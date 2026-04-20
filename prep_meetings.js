#!/usr/bin/env node
'use strict';

// ─── DEPENDENCIES ──────────────────────────────────────────────────────────────
// This script uses CoCo's bundled Node.js and keytar from the Google Workspace MCP.
// Run via:  ~/.snowflake/cortex/.mcp-servers/google-workspace/node prep_meetings.js
//
// If CoCo is installed in a non-default location, update the path below:
const COCO_MCP_DIR = '/Users/' + require('os').userInfo().username +
  '/.snowflake/cortex/.mcp-servers/google-workspace';
const keytar = require(COCO_MCP_DIR + '/node_modules/keytar');

const https    = require('https');
const fs       = require('fs');
const path     = require('path');
const { execSync, spawnSync } = require('child_process');

// ─── CONFIG — edit these before running ────────────────────────────────────────
const ME = 'your.name@snowflake.com';       // Your email — all drafts go here

const VIP_ATTENDEES = [                     // AE(s) you work with — their meetings get priority
  'ae.name@snowflake.com'
];

const MEMORIES_DIR = '/memories';           // CoCo memory directory (usually /memories)

const SNOW_CLI        = '/Applications/SnowflakeCLI.app/Contents/MacOS/snow';
const SNOW_CONNECTION = 'YOUR_CONNECTION';  // `snow connection list` to find yours
const SNOW_WAREHOUSE  = 'YOUR_WAREHOUSE';   // Warehouse you have access to

const SKILL_DIR = path.dirname(__filename);

// ─── OAUTH TOKEN MANAGEMENT ────────────────────────────────────────────────────
// Reads the OAuth token stored by CoCo's Google Workspace MCP in macOS Keychain.
// Token is refreshed automatically if it's about to expire.
async function getAccessToken() {
  const raw = await keytar.getPassword('com.snowflake.cortex.gdrive', 'oauth_tokens');
  if (!raw) throw new Error(
    'No Google OAuth token found. Set up the Google Workspace MCP in Cortex Code first.'
  );
  const tok = JSON.parse(raw);

  if (new Date(tok.expiry) > new Date(Date.now() + 60000)) {
    return tok.access_token;
  }

  const body = new URLSearchParams({
    client_id:     tok.client_id,
    client_secret: tok.client_secret,
    refresh_token: tok.refresh_token,
    grant_type:    'refresh_token'
  }).toString();

  const resp = await apiPost('oauth2.googleapis.com', '/token', body,
    { 'Content-Type': 'application/x-www-form-urlencoded' });

  if (resp.error) throw new Error('Token refresh failed: ' + JSON.stringify(resp));

  tok.access_token = resp.access_token;
  tok.expiry = new Date(Date.now() + resp.expires_in * 1000).toISOString();
  await keytar.setPassword('com.snowflake.cortex.gdrive', 'oauth_tokens', JSON.stringify(tok));
  return tok.access_token;
}

// ─── HTTP HELPERS ──────────────────────────────────────────────────────────────
function apiGet(hostname, path, token) {
  return new Promise((resolve, reject) => {
    const opts = {
      hostname, path, method: 'GET',
      headers: token ? { Authorization: 'Bearer ' + token } : {}
    };
    const req = https.request(opts, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); }
        catch (e) { resolve({ status: res.statusCode, body: data }); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

function apiPost(hostname, path, body, headers, token) {
  return new Promise((resolve, reject) => {
    const h = Object.assign(
      { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
      headers || {},
      token ? { Authorization: 'Bearer ' + token } : {}
    );
    const req = https.request({ hostname, path, method: 'POST', headers: h }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch (e) { resolve(data); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ─── CALENDAR HELPERS ─────────────────────────────────────────────────────────
// Calendar access uses compiled Swift EventKit binaries (NOT the Google Calendar API).
// See README.md for why and how to compile the binaries.
const SKIP_CALENDARS = [
  'US Holidays', 'Siri Suggestions', 'Birthdays',
  'Scheduled Reminders', 'Home', 'iCloud', 'Work', 'Calendar'
];

function fetchCalendarEvents() {
  const bin = path.join(SKILL_DIR, 'read_cal_week');
  if (!fs.existsSync(bin)) {
    throw new Error(
      `read_cal_week binary not found at ${bin}.\n` +
      `Compile it: swiftc src/read_cal_week.swift -o read_cal_week`
    );
  }
  const result = spawnSync(bin, [], { encoding: 'utf8', timeout: 15000 });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error('read_cal_week failed: ' + result.stderr);
  return JSON.parse(result.stdout.trim());
}

function classifyEvent(event) {
  const attendees = event.attendees || [];
  if (!attendees.length) return 'internal';

  const emails    = attendees.map(a => (a.email || '').toLowerCase());
  const hasVip    = VIP_ATTENDEES.some(v => emails.includes(v));
  const hasExternal = emails.some(e =>
    e && !e.endsWith('@snowflake.com') && !e.includes('resource.calendar')
  );
  const isDeclined = (attendees.find(a => a.self) || {}).responseStatus === 'declined';

  if (isDeclined || event.status === 'cancelled') return 'skip';
  if (hasVip && hasExternal) return 'vip-external';
  if (hasVip)                return 'vip-internal';
  if (hasExternal)           return 'external';
  return 'internal';
}

function sortEvents(events) {
  const order = { 'vip-external': 0, 'vip-internal': 1, 'external': 2, 'internal': 3, 'skip': 4 };
  return events
    .map(e => ({ event: e, type: classifyEvent(e) }))
    .filter(({ type }) => type !== 'skip' && type !== 'internal')
    .sort((a, b) => order[a.type] - order[b.type]);
}

// ─── ATTENDEE HELPERS ──────────────────────────────────────────────────────────
function companyFromDomain(email) {
  const domain = email.split('@')[1] || '';
  return domain.split('.')[0].toLowerCase();
}

function inferName(attendee) {
  const raw = attendee.displayName || '';
  if (raw && raw !== attendee.email) return raw;
  const prefix = (attendee.email || '').split('@')[0];
  const parts  = prefix.split(/[._-]/).filter(Boolean);
  if (parts.length < 2) return '';
  return parts.map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join(' ');
}

function formatExtRow(a, sfTitle) {
  const name    = sfTitle && sfTitle.name  ? sfTitle.name  : inferName(a);
  const title   = sfTitle && sfTitle.title ? sfTitle.title : '';
  const company = companyFromDomain(a.email || '');
  return title
    ? `${name} | ${title} | ${a.email} | ${company}`
    : `${name} | ${a.email} | ${company}`;
}

function formatSFRow(a) {
  const name = inferName(a);
  return name ? `${name} | ${a.email}` : a.email;
}

// ─── SNOWHOUSE CONTACT LOOKUP ──────────────────────────────────────────────────
// Batch-queries FIVETRAN.SALESFORCE.CONTACT for job titles of all external attendees.
// Falls back silently if Snowhouse is unavailable (titles will show as blank).
function lookupContactTitles(emails) {
  if (!emails.length) return {};
  try {
    const inList = emails.map(e => `'${e.toLowerCase().replace(/'/g, "\\'")}'`).join(',');
    const sql =
      `SELECT LOWER(EMAIL) AS EMAIL, NAME, NULLIF(TITLE,'') AS TITLE ` +
      `FROM FIVETRAN.SALESFORCE.CONTACT ` +
      `WHERE LOWER(EMAIL) IN (${inList}) ` +
      `QUALIFY ROW_NUMBER() OVER (PARTITION BY LOWER(EMAIL) ORDER BY NULLIF(TITLE,'') DESC NULLS LAST) = 1`;
    const result = execSync(
      `"${SNOW_CLI}" sql --connection ${SNOW_CONNECTION} --warehouse ${SNOW_WAREHOUSE} ` +
      `--query ${JSON.stringify(sql)} --format json`,
      { timeout: 15000 }
    ).toString().trim();
    const rows = JSON.parse(result);
    const map = {};
    for (const r of rows) map[r.EMAIL] = { name: r.NAME, title: r.TITLE || '' };
    return map;
  } catch (e) {
    return {};
  }
}

// ─── DRIVE HELPERS ─────────────────────────────────────────────────────────────
// Searches Google Drive for files named like *master_notes*{company}*.
// Create notes files named:  YourInitials_CompanyName_Master_Notes
async function searchDriveNotes(companyName, token) {
  const query = `name contains 'master_notes' and name contains '${companyName}' and trashed = false`;
  const params = new URLSearchParams({
    q:      query,
    fields: 'files(id,name,mimeType,webViewLink)',
    pageSize: '5'
  });
  const res = await apiGet('www.googleapis.com', '/drive/v3/files?' + params.toString(), token);
  if (res.status !== 200) return null;
  return (res.body.files || [])[0] || null;
}

async function readDriveFile(file, token) {
  if (file.mimeType === 'application/vnd.google-apps.document') {
    const res = await apiGet(
      'www.googleapis.com',
      `/drive/v3/files/${file.id}/export?mimeType=text%2Fplain`,
      token
    );
    return typeof res.body === 'string' ? res.body : JSON.stringify(res.body);
  }
  const res = await apiGet('www.googleapis.com', `/drive/v3/files/${file.id}?alt=media`, token);
  return typeof res.body === 'string' ? res.body : JSON.stringify(res.body);
}

// ─── MEMORY FILE HELPERS ───────────────────────────────────────────────────────
function searchMemoryFiles(companyName) {
  try {
    const files   = fs.readdirSync(MEMORIES_DIR);
    const matches = files.filter(f =>
      f.toLowerCase().includes(companyName.toLowerCase()) ||
      companyName.toLowerCase().includes(f.replace('.md', '').split('_').join(' ').toLowerCase())
    );
    if (!matches.length) return null;
    const content = fs.readFileSync(path.join(MEMORIES_DIR, matches[0]), 'utf8');
    return { filename: matches[0], content };
  } catch (e) {
    return null;
  }
}

// ─── AGENDA BUILDER ────────────────────────────────────────────────────────────
function formatDateTime(event) {
  if (!event.start || !event.start.dateTime) return 'All Day';
  const start       = new Date(event.start.dateTime);
  const end         = new Date(event.end.dateTime);
  const durationMin = Math.round((end - start) / 60000);
  const opts = {
    weekday: 'short', month: 'short', day: 'numeric',
    hour: 'numeric', minute: '2-digit', timeZoneName: 'short'
  };
  return `${start.toLocaleString('en-CA', opts)} (${durationMin} min)`;
}

function getMeetingLink(event) {
  if (event.hangoutLink) return event.hangoutLink;
  const cd = event.conferenceData;
  if (cd && cd.entryPoints) {
    const video = cd.entryPoints.find(e => e.entryPointType === 'video');
    if (video) return video.uri;
  }
  return null;
}

function truncateNotes(text, maxChars = 600) {
  if (!text) return '';
  const clean = text.replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  if (clean.length <= maxChars) return clean;
  return clean.substring(0, maxChars) + '\n[... truncated]';
}

function buildAgenda(event, externalAttendees, driveFile, driveContent, memFile, contactTitles) {
  const title       = event.summary || 'Untitled Meeting';
  const dateStr     = formatDateTime(event);
  const meetingLink = getMeetingLink(event);

  const notesContent = driveContent || (memFile && memFile.content) || null;
  const sourceRef    = driveFile
    ? `Drive: ${driveFile.name} (${driveFile.webViewLink})`
    : memFile
      ? `Memory: /memories/${memFile.filename}`
      : 'No notes found';

  const snowflakeAttendees = (event.attendees || []).filter(a =>
    a.email && a.email.toLowerCase().endsWith('@snowflake.com') &&
    !a.email.toLowerCase().includes('resource.calendar')
  );

  const extRows = externalAttendees.map(a =>
    formatExtRow(a, (contactTitles || {})[a.email ? a.email.toLowerCase() : ''])
  ).join('\n  ');
  const sfRows = snowflakeAttendees.map(formatSFRow).join('\n  ');

  let useCases = 'See notes';
  let keyNotes = '';
  if (notesContent) {
    const lines  = notesContent.split('\n').filter(l => l.trim());
    const useIdx = lines.findIndex(l => /use.?case|focus|objective|goal/i.test(l));
    if (useIdx >= 0) useCases = lines.slice(useIdx, useIdx + 4).join('\n  ');
    keyNotes = truncateNotes(notesContent.substring(0, 500));
  }

  const plainText = `=== PRE-CALL PREP (auto-generated) ===
Meeting: ${title}
When: ${dateStr}${meetingLink ? '\nLink: ' + meetingLink : ''}

EXTERNAL ATTENDEES:
  ${extRows || '(none found)'}

SNOWFLAKE TEAM:
  ${sfRows || '(just you)'}

ACCOUNT CONTEXT:
  ${notesContent ? useCases : 'No notes found — check SFDC or search Drive manually.'}

KEY NOTES:
  ${notesContent ? keyNotes : '—'}

SUGGESTED AGENDA:
  1. Intro / context recap (5 min)
  2. Review use cases and current state (15 min)
  3. Technical deep dive / demo (10 min)
  4. Next steps + action items (5 min)

SOURCE: ${sourceRef}
Generated: ${new Date().toLocaleDateString('en-CA')}
=== END PREP ===`;

  const htmlBody = `
<div style="font-family:Arial,sans-serif;max-width:700px;padding:20px;">
  <h2 style="color:#29B5E8;border-bottom:2px solid #29B5E8;padding-bottom:8px;">
    PRE-CALL PREP: ${escHtml(title)}
  </h2>
  <p><strong>Date/Time:</strong> ${escHtml(dateStr)}</p>
  ${meetingLink ? `<p><strong>Meeting Link:</strong> <a href="${escHtml(meetingLink)}">${escHtml(meetingLink)}</a></p>` : ''}

  <h3 style="color:#444;">External Attendees</h3>
  <table style="border-collapse:collapse;width:100%;">
    <tr style="background:#f0f0f0;">
      <th style="border:1px solid #ddd;padding:6px;text-align:left;">Name</th>
      <th style="border:1px solid #ddd;padding:6px;text-align:left;">Title</th>
      <th style="border:1px solid #ddd;padding:6px;text-align:left;">Email</th>
      <th style="border:1px solid #ddd;padding:6px;text-align:left;">Company</th>
    </tr>
    ${externalAttendees.map(a => {
      const ct        = (contactTitles || {})[a.email ? a.email.toLowerCase() : ''];
      const name      = ct && ct.name  ? ct.name  : (inferName(a) || a.email || '—');
      const titleCell = ct && ct.title ? ct.title : '—';
      const company   = companyFromDomain(a.email || '');
      return `<tr>
        <td style="border:1px solid #ddd;padding:6px;">${escHtml(name)}</td>
        <td style="border:1px solid #ddd;padding:6px;">${escHtml(titleCell)}</td>
        <td style="border:1px solid #ddd;padding:6px;">${escHtml(a.email || '—')}</td>
        <td style="border:1px solid #ddd;padding:6px;">${escHtml(company)}</td>
      </tr>`;
    }).join('')}
  </table>

  <h3 style="color:#444;">Snowflake Team</h3>
  <p>${snowflakeAttendees.map(a => escHtml(inferName(a) || a.email)).join(', ') || '(just you)'}</p>

  <h3 style="color:#444;">Account Context</h3>
  ${notesContent
    ? `<pre style="background:#f9f9f9;padding:12px;border-left:4px solid #29B5E8;white-space:pre-wrap;">${escHtml(truncateNotes(notesContent))}</pre>`
    : '<p style="color:#999;"><em>No notes found in Drive or memory. Check SFDC or search Drive manually.</em></p>'
  }

  <h3 style="color:#444;">Suggested Agenda</h3>
  <ol>
    <li>Intro / context recap <em>(5 min)</em></li>
    <li>Review use cases and current state <em>(15 min)</em></li>
    <li>Technical deep dive / demo <em>(10 min)</em></li>
    <li>Next steps + action items <em>(5 min)</em></li>
  </ol>

  <hr style="border:none;border-top:1px solid #ddd;margin:20px 0;">
  <p style="color:#999;font-size:12px;">
    Source: ${escHtml(sourceRef)} &nbsp;|&nbsp; Generated by se-meeting-prep skill on ${new Date().toLocaleDateString('en-CA')}
  </p>
</div>`;

  return { plainText, htmlBody, title, dateStr };
}

function escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ─── GMAIL DRAFT ───────────────────────────────────────────────────────────────
function buildMimeRaw(subject, htmlBody) {
  const boundary = 'BOUNDARY_' + Date.now();
  const mime = [
    `To: ${ME}`,
    `From: ${ME}`,
    `Subject: ${subject}`,
    'MIME-Version: 1.0',
    `Content-Type: multipart/alternative; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    'Content-Type: text/html; charset=UTF-8',
    '',
    htmlBody,
    '',
    `--${boundary}--`
  ].join('\r\n');
  return Buffer.from(mime).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

async function createGmailDraft(subject, htmlBody, token) {
  const raw  = buildMimeRaw(subject, htmlBody);
  const body = JSON.stringify({ message: { raw } });
  return apiPost('gmail.googleapis.com', '/gmail/v1/users/me/drafts', body, {}, token);
}

// ─── CALENDAR UPDATE ───────────────────────────────────────────────────────────
// Uses the compiled update_cal_event Swift binary.
// Original event content (Zoom link etc.) is preserved at the TOP.
function updateCalendarDescriptionLocal(eventId, _calName, currentDesc, newPrepText) {
  const original = currentDesc
    ? currentDesc.replace(/=== PRE-CALL PREP[\s\S]*?=== END PREP ===/g, '').trim()
    : '';
  const newDesc = (original ? original + '\n\n' : '') + newPrepText;

  const bin    = path.join(SKILL_DIR, 'update_cal_event');
  const result = spawnSync(bin, [eventId], { input: newDesc, encoding: 'utf8', timeout: 15000 });
  if (result.error)      return { status: 500, error: result.error.message };
  if (result.status !== 0) return { status: 404, error: result.stderr };
  return { status: 200 };
}

// ─── MAIN ──────────────────────────────────────────────────────────────────────
async function main() {
  console.log('SE Meeting Prep - Week of', new Date().toLocaleDateString('en-CA'));
  console.log('='.repeat(60));

  const token = await getAccessToken();
  console.log('OAuth token loaded (Drive + Gmail)');

  const allEvents = fetchCalendarEvents();
  console.log(`Found ${allEvents.length} events this week`);

  const sorted = sortEvents(allEvents);
  console.log(`Processing ${sorted.length} external/VIP meetings\n`);

  if (!sorted.length) {
    console.log('No external meetings found this week. Nothing to prep!');
    return;
  }

  const allExternalEmails = [...new Set(
    sorted.flatMap(({ event }) =>
      (event.attendees || [])
        .filter(a =>
          a.email &&
          !a.email.toLowerCase().endsWith('@snowflake.com') &&
          !a.email.includes('resource.calendar')
        )
        .map(a => a.email.toLowerCase())
    )
  )];
  const contactTitles = lookupContactTitles(allExternalEmails);
  console.log(`Loaded SFDC titles for ${Object.keys(contactTitles).length}/${allExternalEmails.length} external contacts\n`);

  const results = [];

  for (const { event, type } of sorted) {
    const title = event.summary || 'Untitled Meeting';
    const label = type === 'vip-external' ? '[VIP + Customer]'
      : type === 'vip-internal' ? '[VIP internal]'
      : '[Customer]';

    console.log(`\nProcessing: ${title} ${label}`);

    const attendees = event.attendees || [];
    const externalAttendees = attendees.filter(a =>
      a.email &&
      !a.email.toLowerCase().endsWith('@snowflake.com') &&
      !a.email.includes('resource.calendar')
    );

    let driveFile    = null;
    let driveContent = null;
    let memFile      = null;

    if (externalAttendees.length > 0) {
      const company = companyFromDomain(externalAttendees[0].email);
      console.log(`  Company: ${company}`);

      driveFile = await searchDriveNotes(company, token);
      if (driveFile) {
        console.log(`  Drive notes found: ${driveFile.name}`);
        try { driveContent = await readDriveFile(driveFile, token); }
        catch (e) { console.log(`  Could not read Drive file: ${e.message}`); }
      } else {
        memFile = searchMemoryFiles(company);
        if (memFile) {
          console.log(`  Memory notes found: ${memFile.filename}`);
        } else {
          console.log(`  No notes found for: ${company}`);
        }
      }
    }

    const { plainText, htmlBody, dateStr } = buildAgenda(
      event, externalAttendees, driveFile, driveContent, memFile, contactTitles
    );

    const datePart = event.start && event.start.dateTime
      ? new Date(event.start.dateTime).toLocaleDateString('en-CA', {
          weekday: 'short', month: 'short', day: 'numeric'
        })
      : 'This Week';
    const subject = `Pre-call Prep - ${title} - ${datePart}`;

    let gmailOk   = false;
    let calOk     = false;
    let calSkipped = false;

    try {
      const draft = await createGmailDraft(subject, htmlBody, token);
      if (draft.id) {
        console.log(`  Gmail draft created: ${draft.id}`);
        gmailOk = true;
      } else {
        console.log(`  Gmail draft failed:`, JSON.stringify(draft));
      }
    } catch (e) {
      console.log(`  Gmail draft error: ${e.message}`);
    }

    try {
      const currentDesc = event.description || '';
      const res = updateCalendarDescriptionLocal(event.id, event._calName, currentDesc, plainText);
      if (res.status === 200) {
        console.log(`  Calendar description updated`);
        calOk = true;
      } else if (res.status === 404) {
        console.log(`  Calendar update skipped (not found in macOS Calendar)`);
        calSkipped = true;
      } else {
        console.log(`  Calendar update failed (${res.status}):`, res.error || '');
      }
    } catch (e) {
      console.log(`  Calendar update error: ${e.message}`);
    }

    const status = gmailOk && calOk     ? '[OK]'
      : gmailOk && calSkipped           ? '[OK] (Gmail only — not organizer)'
      : gmailOk                         ? '[OK] (Gmail only)'
      : '[FAIL]';

    results.push({ title, dateStr, type, status });
  }

  console.log('\n' + '='.repeat(60));
  console.log('SUMMARY');
  console.log('='.repeat(60));
  for (const r of results) {
    const typeLabel = r.type.includes('vip') ? '* ' : '  ';
    console.log(`${typeLabel}${r.status} ${r.title} (${r.dateStr})`);
  }
  console.log('\n* = VIP/AE meeting (prioritized)');
  console.log(`\nCheck Gmail Drafts and your calendar events for the prep content.`);
}

main().catch(err => {
  console.error('\nERROR:', err.message);
  process.exit(1);
});
