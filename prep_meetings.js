#!/usr/bin/env node
'use strict';

const keytar = require('/Users/cwalker/.snowflake/cortex/.mcp-servers/google-workspace/node_modules/keytar');
const https = require('https');
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { execSync, spawnSync, spawn } = require('child_process');
const crypto = require('crypto');

// ─── CONFIG ────────────────────────────────────────────────────────────────────
const ME = 'carson.walker@snowflake.com';
const VIP_ATTENDEES = ['ali.roshanzamir@snowflake.com', 'shaw.liu@snowflake.com'];
const MEMORIES_DIR = '/memories';
const SKILL_DIR = path.dirname(__filename);
const SNOW_CLI = '/Applications/SnowflakeCLI.app/Contents/MacOS/snow';
const SNOW_CONNECTION = 'SNOWHOUSE';
const SNOW_WAREHOUSE = 'SNOWADHOC';
const STATE_FILE = path.join(SKILL_DIR, 'prep_state.json');

// Drive folder IDs for account notes
const ALI_FOLDER_ID   = '1GkS8HPr4vlqtwnMvXkkqF0vHtyWdgbx6';
const SHAW_FOLDER_ID  = '1TBWK0SYOQTg6aar5BACMMZ9ZCOX80Ow7';

// ─── PREP STATE (incremental run tracking) ────────────────────────────────────
function computeChecksum(event, externalAttendees) {
  const key = [
    event.summary || '',
    (event.start && event.start.dateTime) || '',
    (event.end && event.end.dateTime) || '',
    (externalAttendees || []).map(a => a.email).sort().join(',')
  ].join('|');
  return crypto.createHash('md5').update(key).digest('hex');
}

function loadPrepState() {
  try {
    if (fs.existsSync(STATE_FILE)) {
      const data = JSON.parse(fs.readFileSync(STATE_FILE, 'utf8'));
      const cutoff = Date.now() - 21 * 24 * 60 * 60 * 1000;
      for (const id of Object.keys(data)) {
        if (new Date(data[id].processedAt).getTime() < cutoff) delete data[id];
      }
      return data;
    }
  } catch (e) {}
  return {};
}

function savePrepState(state) {
  try { fs.writeFileSync(STATE_FILE, JSON.stringify(state, null, 2)); } catch (e) {}
}

// ─── OAUTH TOKEN MANAGEMENT ────────────────────────────────────────────────────
async function getAccessToken() {
  const raw = await keytar.getPassword('com.snowflake.cortex.gdrive', 'oauth_tokens');
  if (!raw) throw new Error('No token found in keytar.');
  const tok = JSON.parse(raw);

  const expiry = new Date(tok.expiry);
  if (expiry > new Date(Date.now() + 60000)) return tok.access_token;

  const body = new URLSearchParams({
    client_id: tok.client_id,
    client_secret: tok.client_secret,
    refresh_token: tok.refresh_token,
    grant_type: 'refresh_token'
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

function apiPatch(hostname, path, body, token) {
  return new Promise((resolve, reject) => {
    const bodyStr = typeof body === 'string' ? body : JSON.stringify(body);
    const req = https.request({
      hostname, path, method: 'PATCH',
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(bodyStr)
      }
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); }
        catch (e) { resolve({ status: res.statusCode, body: data }); }
      });
    });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}

// ─── CALENDAR (macOS EventKit) ─────────────────────────────────────────────────
function fetchCalendarEvents(weekOffset) {
  const bin = path.join(SKILL_DIR, 'read_cal_week');
  const args = weekOffset > 0 ? ['--next-week'] : [];
  const result = spawnSync(bin, args, { encoding: 'utf8', timeout: 15000 });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error('read_cal_week failed: ' + result.stderr);
  const events = JSON.parse(result.stdout.trim());
  // Fallback: if EventKit returned no attendees, scan description for external emails
  // (handles sync lag and meetings where invites were added outside the guest list)
  const emailRe = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
  for (const ev of events) {
    if (!(ev.attendees || []).length && ev.description) {
      const found = (ev.description.match(emailRe) || [])
        .filter(e => !e.endsWith('@snowflake.com') && !e.includes('resource.calendar') && !e.includes('googleusercontent'));
      if (found.length) {
        ev.attendees = found.map(email => ({ email: email.toLowerCase(), displayName: email.split('@')[0], _fromDesc: true }));
      }
    }
  }
  return events;
}

function classifyEvent(event) {
  const attendees = event.attendees || [];
  if (!attendees.length) return 'internal';
  const emails = attendees.map(a => (a.email || '').toLowerCase());
  const hasVip = VIP_ATTENDEES.some(v => emails.includes(v));
  const hasExternal = emails.some(e => e && !e.endsWith('@snowflake.com') && !e.includes('resource.calendar'));
  const isDeclined = (attendees.find(a => a.self) || {}).responseStatus === 'declined';
  if (isDeclined || event.status === 'cancelled') return 'skip';
  if (hasVip && hasExternal) return 'vip-external';
  if (hasVip) return 'vip-internal';
  if (hasExternal) return 'external';
  return 'internal';
}

function sortEvents(events) {
  const order = { 'vip-external': 0, 'vip-internal': 1, 'external': 2, 'internal': 3, 'skip': 4 };
  return events
    .map(e => ({ event: e, type: classifyEvent(e) }))
    .filter(({ type }) => type !== 'skip' && type !== 'internal')
    .sort((a, b) => order[a.type] - order[b.type]);
}

// ─── COMPANY EXTRACTION ─────────────────────────────────────────────────────────
const STOP_WORDS = new Set(['inc', 'ltd', 'llc', 'corp', 'co', 'the', 'and', 'group', 'lp', 'de', 'by', 'its', 'for']);
const MEETING_KEYWORDS_RE = /^(QBR|Demo|Review|Discovery|Call|Sync|Meeting|Check-in|Checkin|Workshop|Debrief|Follow.?up|Intro|Introduction|Kickoff|Kick.off|Onboarding|Training|POC|Pilot|Update|Presentation|Chat|Lunch|Coffee|Happy|Hour|Quarterly|Monthly|Weekly|Bi-weekly|1on1|1:1|Session|Consult|Planning|Strategy)$/i;
const SNOWFLAKE_WORD_RE = /\bsnowflake\b/i;
const SF_TRAILING_KW_RE = /\s+\b(QBR|Demo|Review|Discovery|Call|Sync|Meeting|Workshop|Intro|Introduction|Kickoff|Onboarding|Training|POC|Pilot|Presentation|Quarterly|Monthly|Weekly|Renewal|Scoping|Team|Account|Lunch|Coffee|Happy|Hour)\b.*/i;

function extractCompanyFromTitle(title) {
  if (!title) return null;
  // Strip leading status words: "HOLD:", "HOLD -", "Cancelled:", etc.
  const t = title.replace(/^(?:hold|tbd|cancelled?|canceled?)\s*[:–—\-]\s*/i, '').trim();

  // "Snowflake/CompanyName [-topic]" or "Snowflake+CompanyName [-topic]"
  const sfSlashM = t.match(/^snowflake\s*[\/+]\s*([^\-—(]+)/i);
  if (sfSlashM) {
    let c = sfSlashM[1].trim().replace(SF_TRAILING_KW_RE, '').trim();
    if (c && c.length >= 2 && !SNOWFLAKE_WORD_RE.test(c)) return c;
  }

  // "Snowflake x Name - Company" → take what's after the dash
  const sfXM = t.match(/^snowflake\s+x\s+\S[\s\S]*?\s*[-—]\s*(.+)$/i);
  if (sfXM) {
    const c = sfXM[1].trim();
    if (c && !SNOWFLAKE_WORD_RE.test(c)) return c;
  }

  // "... with CompanyName [at end or before dash/pipe]"
  const withM = t.match(/\bwith\s+([A-Za-z0-9][^-|—(]{2,50}?)(?:\s*[-|—(]|\s*$)/i);
  if (withM) {
    const c = withM[1].trim();
    if (c && !SNOWFLAKE_WORD_RE.test(c)) return c;
  }

  // "CompanyName - topic" or "CompanyName | topic"
  const sepM = t.match(/^([A-Za-z0-9 &'.]+?)\s*[-|—]\s*/);
  if (sepM) {
    const c = sepM[1].trim();
    if (c && !MEETING_KEYWORDS_RE.test(c) && !SNOWFLAKE_WORD_RE.test(c)) return c;
  }

  // "CompanyName <meeting keyword>"
  const kwM = t.match(/^([A-Za-z0-9][A-Za-z0-9 &'.]{1,40}?)\s+(?:QBR|Demo|Review|Discovery|Call|Sync|Meeting|Workshop|Debrief|Follow|Intro|Kickoff|Onboarding|Training|POC|Pilot|Presentation|Chat|Quarterly|Monthly|Weekly)\b/i);
  if (kwM) {
    const c = kwM[1].trim();
    if (c && !SNOWFLAKE_WORD_RE.test(c)) return c;
  }

  return null;
}

function companyFromDomain(email) {
  const domain = (email || '').split('@')[1] || '';
  return domain.split('.')[0].toLowerCase();
}

function domainToKeywords(domain) {
  if (!domain) return [];
  const kws = [domain];
  if (domain.length > 10) {
    for (let len = 4; len <= 7; len++) {
      const p = domain.slice(0, len);
      if (!kws.includes(p)) kws.push(p);
    }
  }
  return kws;
}

function companyToKeywords(name) {
  if (!name) return [];
  const words = name.trim().split(/[\s\-_.,]+/).filter(w => w.length >= 3 && !STOP_WORDS.has(w.toLowerCase()));
  const merged = name.toLowerCase().replace(/[\s\-_.,]+/g, '');
  return [...new Set([...words, merged])].filter(Boolean);
}

function getCompanyInfo(event, externalAttendees) {
  const titleCompany = extractCompanyFromTitle(event.summary || '');
  const domainRaw = externalAttendees.length > 0 ? companyFromDomain(externalAttendees[0].email) : null;

  const titleKws = titleCompany ? companyToKeywords(titleCompany) : [];
  const domainKws = domainRaw ? domainToKeywords(domainRaw) : [];
  const allKeywords = [...new Set([...titleKws, ...domainKws])];

  const displayName = titleCompany || domainRaw || 'Unknown';

  let assumption = null;
  if (titleCompany && domainRaw && !titleCompany.toLowerCase().replace(/\s/g,'').includes(domainRaw.replace(/\s/g,''))) {
    assumption = `[ASSUMPTION: Account matched as "${titleCompany}" from meeting title; attendee domain is "${domainRaw}"]`;
  } else if (titleCompany && !domainRaw) {
    assumption = `[ASSUMPTION: Account "${titleCompany}" inferred from meeting title — no external attendee email found]`;
  } else if (!titleCompany && domainRaw) {
    assumption = `[ASSUMPTION: Account matched by email domain "${domainRaw}" — no company name found in meeting title]`;
  }

  return { displayName, keywords: allKeywords, assumption };
}

// ─── DRIVE: ACCOUNT FOLDER SEARCH ──────────────────────────────────────────────
async function findAccountFolder(keywords, token) {
  for (const kw of keywords) {
    if (!kw || kw.length < 3) continue;
    const kwSafe = kw.replace(/'/g, "\\'");
    for (const parentId of [ALI_FOLDER_ID, SHAW_FOLDER_ID]) {
      const q = `'${parentId}' in parents and mimeType = 'application/vnd.google-apps.folder' and name contains '${kwSafe}' and trashed = false`;
      const params = new URLSearchParams({ q, fields: 'files(id,name,webViewLink)', pageSize: '5' });
      const res = await apiGet('www.googleapis.com', '/drive/v3/files?' + params.toString(), token);
      if (res.status === 200) {
        const folders = res.body.files || [];
        if (folders.length) return { folder: folders[0], keyword: kw };
      }
    }
  }
  return null;
}

async function getNotesFromFolder(folderId, token) {
  const q = `'${folderId}' in parents and trashed = false`;
  const params = new URLSearchParams({ q, fields: 'files(id,name,mimeType,webViewLink)', pageSize: '20', orderBy: 'modifiedTime desc' });
  const res = await apiGet('www.googleapis.com', '/drive/v3/files?' + params.toString(), token);
  if (res.status !== 200) return null;
  const files = res.body.files || [];
  if (!files.length) return null;
  return files.find(f => /master.?notes/i.test(f.name) && f.mimeType === 'application/vnd.google-apps.document')
    || files.find(f => /master.?notes/i.test(f.name))
    || files.find(f => f.mimeType === 'application/vnd.google-apps.document')
    || files[0];
}

async function broadDriveSearch(keywords, token) {
  for (const kw of keywords) {
    if (!kw || kw.length < 3) continue;
    const kwSafe = kw.replace(/'/g, "\\'");
    for (const nameContains of ['master_notes', 'masternotes', 'master']) {
      const q = `name contains '${nameContains}' and name contains '${kwSafe}' and trashed = false`;
      const params = new URLSearchParams({ q, fields: 'files(id,name,mimeType,webViewLink)', pageSize: '5' });
      const res = await apiGet('www.googleapis.com', '/drive/v3/files?' + params.toString(), token);
      if (res.status === 200) {
        const files = res.body.files || [];
        if (files.length) return files[0];
      }
    }
  }
  return null;
}

async function findDriveNotesForCompany(keywords, token) {
  // Strategy 1: matching account subfolder in Ali/Shaw folders
  const folderMatch = await findAccountFolder(keywords, token);
  if (folderMatch) {
    const notesFile = await getNotesFromFolder(folderMatch.folder.id, token);
    return {
      file: notesFile,
      folderName: folderMatch.folder.name,
      folderLink: folderMatch.folder.webViewLink,
      method: notesFile ? 'folder' : 'folder-empty'
    };
  }
  // Strategy 2: broad file search
  const file = await broadDriveSearch(keywords, token);
  if (file) return { file, folderName: null, folderLink: null, method: 'search' };
  return null;
}

async function readDriveFile(file, token) {
  if (file.mimeType && file.mimeType.startsWith('application/vnd.google-apps.')) {
    const exportMime = file.mimeType === 'application/vnd.google-apps.spreadsheet'
      ? 'text%2Fcsv'
      : 'text%2Fplain';
    const res = await apiGet(
      'www.googleapis.com',
      `/drive/v3/files/${file.id}/export?mimeType=${exportMime}`,
      token
    );
    if (res.status >= 400) throw new Error(`Drive export error ${res.status}`);
    return typeof res.body === 'string' ? res.body : JSON.stringify(res.body);
  }
  const res = await apiGet('www.googleapis.com', `/drive/v3/files/${file.id}?alt=media`, token);
  if (res.status >= 400) throw new Error(`Drive read error ${res.status}`);
  return typeof res.body === 'string' ? res.body : JSON.stringify(res.body);
}

// ─── MEMORY FILE FALLBACK ──────────────────────────────────────────────────────
function searchMemoryFiles(keywords) {
  try {
    const files = fs.readdirSync(MEMORIES_DIR);
    for (const kw of keywords) {
      const matches = files.filter(f => f.toLowerCase().includes(kw.toLowerCase()));
      if (matches.length) {
        const content = fs.readFileSync(path.join(MEMORIES_DIR, matches[0]), 'utf8');
        return { filename: matches[0], content };
      }
    }
  } catch (e) {}
  return null;
}

// ─── SNOWHOUSE CONTACT LOOKUP ──────────────────────────────────────────────────
function lookupContactTitles(emails) {
  if (!emails.length) return {};
  try {
    const inList = emails.map(e => `'${e.toLowerCase().replace(/'/g, "\\'")}'`).join(',');
    const sql = `SELECT LOWER(EMAIL) AS EMAIL, NAME, NULLIF(TITLE,'') AS TITLE ` +
      `FROM FIVETRAN.SALESFORCE.CONTACT ` +
      `WHERE LOWER(EMAIL) IN (${inList}) ` +
      `QUALIFY ROW_NUMBER() OVER (PARTITION BY LOWER(EMAIL) ORDER BY NULLIF(TITLE,'') DESC NULLS LAST) = 1`;
    const result = execSync(
      `"${SNOW_CLI}" sql --connection ${SNOW_CONNECTION} --warehouse ${SNOW_WAREHOUSE} --query ${JSON.stringify(sql)} --format json`,
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

// ─── NOTES PARSING ─────────────────────────────────────────────────────────────
function parseNotesContent(text) {
  if (!text) return { useCases: null, recentNotes: null, raw: null };
  const lines = text.split('\n').map(l => l.trimEnd());
  const clean = lines.join('\n').replace(/\n{4,}/g, '\n\n\n').trim();

  // Extract use cases / focus section
  let useCases = null;
  const ucIdx = lines.findIndex(l => /use.?case|focus|objective|goal|priority|use case/i.test(l));
  if (ucIdx >= 0) {
    const section = lines.slice(ucIdx, ucIdx + 8).filter(l => l.trim()).join('\n');
    useCases = section.trim();
  }

  // Extract recent notes — look for dated entries or the last meaningful block
  let recentNotes = null;
  const dateLineIdx = lines.findLastIndex ? lines.findLastIndex(l => /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{4}[-\/]\d{2}|\d{1,2}\/\d{1,2})/i.test(l))
    : [...lines].reverse().findIndex(l => /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{4}[-\/]\d{2}|\d{1,2}\/\d{1,2})/i.test(l));
  const realIdx = lines.findLastIndex ? dateLineIdx : (dateLineIdx >= 0 ? lines.length - 1 - dateLineIdx : -1);

  if (realIdx >= 0) {
    recentNotes = lines.slice(realIdx).filter(l => l.trim()).slice(0, 15).join('\n').trim();
  }

  // Fallback: first 1200 chars of clean content
  const rawSnippet = clean.length > 1200 ? clean.substring(0, 1200) + '\n[... truncated]' : clean;

  return { useCases, recentNotes, raw: rawSnippet };
}

// ─── CORTEX AI NOTES SUMMARIZER (ASYNC) ──────────────────────────────────────
function cortexSummarizeAsync(notesText) {
  return new Promise((resolve) => {
    if (!notesText || notesText.length < 200) return resolve(null);
    const safe = notesText.replace(/\r/g, '').replace(/&/g, ' and ');
    const truncated = safe.length > 5000 ? safe.substring(0, 5000) + '\n[...truncated]' : safe;

    const prompt = `You are an SE meeting prep assistant. Analyze these account notes and produce a clean summary.

Extract TWO sections:

1. RECENT SE NOTES: Find the most recent dated meeting entries. PRIORITIZE 2025 and 2026 content — ignore 2024 or older unless nothing more recent exists. Write a SYNTHESIZED summary of 3-5 bullets — do NOT list each raw note line as its own bullet. Instead, group related ideas and capture the key themes: what the customer is trying to accomplish, key decisions or blockers, Snowflake features/products discussed, and any clear next steps or action items. Fix ALL spelling mistakes.

2. USE CASES: For each use case, initiative, or workload mentioned, write one line: "• [Name]: [1-sentence description of what they are doing or trying to achieve]". If SE Comments or SE observations exist for that use case, add on the next line: "  → SE: [cleaned up SE observation]". Also include a brief Use Case Summary if one exists in the notes.

3. ACTION ITEMS: List any open action items, follow-ups, or prep tasks from the most recent notes that are still relevant for an upcoming meeting. If none are mentioned, write "• None". Max 5 bullets.

IMPORTANT RULES:
- Fix ALL spelling errors and typos in your output
- Be concise — no filler text
- Only use facts from the notes — do not invent anything
- Keep Snowflake product/feature names (Cortex, Cortex Analyst, Cortex Search, etc.)
- If a use case has a "Summary" or "SE Comments" subsection, include those

NOTES:
${truncated}

Respond with NO preamble in exactly this format:
RECENT SE NOTES:
• [point]

USE CASES:
• [name]: [description]
  → SE: [comment if present]

ACTION ITEMS:
• [item or "None"]`;

    let resolved = false;
    const done = (result) => { if (!resolved) { resolved = true; resolve(result); } };

    const sqlEscaped = prompt.replace(/'/g, "''");
    const sql = `SELECT SNOWFLAKE.CORTEX.COMPLETE('mistral-large2', '${sqlEscaped}')::string AS response`;

    const child = spawn(
      SNOW_CLI,
      ['sql', '--connection', SNOW_CONNECTION, '--warehouse', SNOW_WAREHOUSE, '--query', sql, '--format', 'json'],
      { stdio: ['ignore', 'pipe', 'pipe'] }
    );

    let stdout = '';
    child.stdout.on('data', d => { stdout += d; });
    child.on('close', (code) => {
      if (code !== 0) return done(null);
      try {
        const rows = JSON.parse(stdout.trim());
        if (rows && rows.length > 0 && rows[0].RESPONSE) {
          const text = rows[0].RESPONSE.trim();
          const recentMatch = text.match(/RECENT SE NOTES:\s*([\s\S]*?)(?=\s*USE CASES:|$)/i);
          const ucMatch = text.match(/USE CASES:\s*([\s\S]*?)(?=\s*ACTION ITEMS:|$)/i);
          const aiMatch = text.match(/ACTION ITEMS:\s*([\s\S]*?)$/i);
          done({
            recentNotes: recentMatch ? recentMatch[1].trim() : null,
            useCases: ucMatch ? ucMatch[1].trim() : null,
            actionItems: aiMatch ? aiMatch[1].trim() : null
          });
        } else { done(null); }
      } catch (e) { done(null); }
    });
    child.on('error', () => done(null));

    const timer = setTimeout(() => { child.kill('SIGKILL'); done(null); }, 120000);
    timer.unref();
  });
}

// ─── DYNAMIC AGENDA BUILDER ───────────────────────────────────────────────────
function detectMeetingType(title) {
  const t = (title || '').toLowerCase();
  if (/\bqbr\b|quarterly.?business.?review|exec(utive)?.?review|ebr\b/.test(t)) return 'qbr';
  if (/\bdemo\b|demonstration/.test(t)) return 'demo';
  if (/\bdiscovery\b/.test(t)) return 'discovery';
  if (/\bintro\b|introduction|kick.?off/.test(t)) return 'intro';
  if (/\bpoc\b|proof.?of.?concept|pilot/.test(t)) return 'poc';
  if (/\brenewal\b|scoping|commercial|pricing|negotiat/.test(t)) return 'renewal';
  if (/\bworkshop\b|training|enablement|hands.?on/.test(t)) return 'workshop';
  if (/\bcheck.?in\b|sync\b|weekly\b|monthly\b|bi.?weekly\b|\b1.?on.?1\b|standup/.test(t)) return 'sync';
  if (/\barchitect(ure)?\b|deep.?dive|technical.?review/.test(t)) return 'technical';
  if (/\bfollow.?up\b/.test(t)) return 'followup';
  return 'general';
}

function detectAudience(externalAttendees, contactTitles) {
  const titles = externalAttendees
    .map(a => ((contactTitles[a.email ? a.email.toLowerCase() : ''] || {}).title || '').toLowerCase())
    .filter(Boolean);
  const isExec = titles.some(t => /\bceo\b|\bcto\b|\bcdo\b|\bcoo\b|\bcfo\b|\bciso\b|chief\s|\bpresident\b|\bsvp\b|\bevp\b|\bvp\b|vice.?pres/.test(t));
  const isTechnical = titles.some(t => /engineer|architect|developer|data\s+sci|analyst|devops|platform|dba|infrastructure/.test(t));
  const isDirector = titles.some(t => /\bdirector\b|head\s+of|practice\s+lead/.test(t));
  if (isExec) return 'executive';
  if (isTechnical && !isDirector) return 'technical';
  return 'general';
}

function buildDynamicAgenda(event, topicFromTitle, externalAttendees, contactTitles, aiParsed, notesContent, durationMin) {
  const title = event.summary || '';
  const type = detectMeetingType(title);
  const audience = detectAudience(externalAttendees, contactTitles);
  const hasNotes = !!notesContent;
  const hasActionItems = aiParsed && aiParsed.actionItems && !/^•\s*none\.?$/i.test(aiParsed.actionItems.trim());

  // Pull first named use case for agenda specificity
  const firstUseCase = (() => {
    if (!aiParsed || !aiParsed.useCases) return null;
    const m = aiParsed.useCases.match(/^•\s*([^:\n]+):/m);
    return m ? m[1].trim() : null;
  })();

  const isShort = durationMin > 0 && durationMin <= 30;
  const isLong  = durationMin > 0 && durationMin >= 75;

  const t = (min) => isShort ? `${Math.round(min / 2)} min` : isLong ? `${Math.round(min * 1.5)} min` : `${min} min`;

  // Action item review item (injected when prior open items exist)
  const actionItemStep = hasActionItems
    ? `Review open action items from last meeting (${t(5)})`
    : null;

  // Use-case specific demo/technical step
  const useCaseStep = firstUseCase
    ? `${firstUseCase} — demo / status update (${t(15)})`
    : null;

  let steps = [];

  switch (type) {
    case 'qbr':
      if (audience === 'executive') {
        steps = [
          `Business outcomes review — wins, metrics, progress (${t(10)})`,
          useCaseStep || `Strategic use case update (${t(15)})`,
          `Snowflake roadmap & upcoming investments (${t(10)})`,
          `Partnership priorities + executive asks (${t(10)})`,
          `Next steps + owners (${t(5)})`
        ];
      } else {
        steps = [
          `Recap: what's gone well, what hasn't (${t(10)})`,
          useCaseStep || `Current use case status & blockers (${t(15)})`,
          `Roadmap / upcoming features relevant to your use cases (${t(10)})`,
          actionItemStep,
          `Next steps + owners (${t(5)})`
        ];
      }
      break;

    case 'demo':
      if (audience === 'executive') {
        steps = [
          `Business problem framing — confirm what we're solving (${t(5)})`,
          useCaseStep || `Demo: ${topicFromTitle} (${t(20)})`,
          `Business case / value summary (${t(5)})`,
          `Next steps: POC, pricing, or deeper technical session (${t(5)})`
        ];
      } else {
        steps = [
          `Architecture / environment overview (${t(5)})`,
          useCaseStep || `Live demo: ${topicFromTitle} (${t(20)})`,
          `Technical Q&A + integration questions (${t(10)})`,
          `POC scoping / next steps (${t(5)})`
        ];
      }
      break;

    case 'discovery':
      steps = [
        `Introductions + context (who's who, goals for today) (${t(5)})`,
        `Current state: data stack, key pain points, team structure (${t(15)})`,
        `Top priorities for the next 6–12 months (${t(10)})`,
        `Snowflake fit discussion (${t(5)})`,
        `Agree on next steps (${t(5)})`
      ];
      break;

    case 'intro':
      steps = [
        `Introductions — their role, our team, agenda for today (${t(5)})`,
        `Their current landscape + biggest data challenges (${t(10)})`,
        `Snowflake overview — relevant to their space (${t(10)})`,
        `Identify best next step: discovery, demo, or POC (${t(5)})`
      ];
      break;

    case 'poc':
      steps = [
        `POC status: what's working, what's blocked (${t(10)})`,
        useCaseStep || `Technical walkthrough / testing: ${topicFromTitle} (${t(20)})`,
        `Success criteria check — are we on track? (${t(10)})`,
        `Blockers, open questions, owner assignments (${t(10)})`,
        `Next milestone + timeline (${t(5)})`
      ];
      break;

    case 'renewal':
      steps = [
        `Relationship recap — value delivered, key wins (${t(10)})`,
        `Current usage + expansion opportunities (${t(10)})`,
        `Commercial discussion: terms, timeline, approvals (${t(15)})`,
        `Confirm mutual next steps (${t(5)})`
      ];
      break;

    case 'workshop':
      steps = [
        `Objectives + agenda review (${t(5)})`,
        `Overview / concepts: ${topicFromTitle} (${t(10)})`,
        `Hands-on: ${firstUseCase || topicFromTitle} (${t(isLong ? 45 : 25)})`,
        `Q&A + troubleshooting (${t(10)})`,
        `Wrap-up: resources + next steps (${t(5)})`
      ];
      break;

    case 'sync':
      steps = [
        actionItemStep || `Quick wins / updates since last sync (${t(5)})`,
        `${topicFromTitle !== title ? topicFromTitle : 'Open topics / blockers'} (${t(isShort ? 10 : 15)})`,
        `Next steps + owners (${t(5)})`
      ];
      break;

    case 'technical':
      steps = [
        `Problem statement + current architecture (${t(10)})`,
        useCaseStep || `Deep dive: ${topicFromTitle} (${t(20)})`,
        `Integration points, edge cases, performance (${t(10)})`,
        `Recommended approach + trade-offs (${t(10)})`,
        `Action items + follow-up resources (${t(5)})`
      ];
      break;

    case 'followup':
      steps = [
        `Review open action items from last meeting (${t(10)})`,
        useCaseStep || `${topicFromTitle !== title ? topicFromTitle : 'Status update'} (${t(15)})`,
        `Open questions + blockers (${t(10)})`,
        `Agree on next steps + owners (${t(5)})`
      ];
      break;

    default: // general
      if (!hasNotes) {
        steps = [
          `Introductions + context recap (${t(5)})`,
          `Current priorities and pain points (${t(15)})`,
          `Snowflake capabilities relevant to their use cases (${t(10)})`,
          `Next steps (${t(5)})`
        ];
      } else if (audience === 'executive') {
        steps = [
          `Relationship update + wins since last meeting (${t(5)})`,
          useCaseStep || `Strategic use case status (${t(15)})`,
          actionItemStep,
          `Executive asks / roadmap alignment (${t(10)})`,
          `Next steps + owners (${t(5)})`
        ];
      } else {
        steps = [
          `Context recap + relationship update (${t(5)})`,
          useCaseStep || `${topicFromTitle !== title ? topicFromTitle : 'Review current use cases / status'} (${t(15)})`,
          actionItemStep,
          `Technical follow-up / demo / Q&A (${t(10)})`,
          `Next steps + action items (${t(5)})`
        ];
      }
  }

  return steps.filter(Boolean);
}

// ─── AGENDA BUILDER ────────────────────────────────────────────────────────────
function formatDateTime(event) {
  if (!event.start || !event.start.dateTime) return 'All Day';
  const start = new Date(event.start.dateTime);
  const end = new Date(event.end ? event.end.dateTime : event.start.dateTime);
  const durationMin = Math.round((end - start) / 60000);
  const opts = { weekday: 'short', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit', timeZoneName: 'short' };
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

function inferName(attendee) {
  const raw = attendee.displayName || '';
  if (raw && raw !== attendee.email) return raw;
  const prefix = (attendee.email || '').split('@')[0];
  const parts = prefix.split(/[._-]/).filter(Boolean);
  if (parts.length < 2) return '';
  return parts.map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join(' ');
}

function escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function buildAgenda(event, externalAttendees, companyInfo, noteResult, notesContent, contactTitles, aiParsed) {
  const title = event.summary || 'Untitled Meeting';
  const dateStr = formatDateTime(event);
  const meetingLink = getMeetingLink(event);

  // Original meeting description (agenda already in the invite)
  const originalDesc = (event.description || '')
    .replace(/=== PRE-CALL PREP[\s\S]*?=== END PREP ===/g, '').trim();

  const driveFile = noteResult && noteResult.file;
  const folderName = noteResult && noteResult.folderName;
  const folderLink = noteResult && noteResult.folderLink;
  const method = noteResult && noteResult.method;
  const parsed = parseNotesContent(notesContent);
  const useCases = (aiParsed && aiParsed.useCases) || parsed.useCases;
  const recentNotes = (aiParsed && aiParsed.recentNotes) || parsed.recentNotes;
  const raw = parsed.raw;

  const sourceRef = driveFile
    ? `${driveFile.name} — ${driveFile.webViewLink}`
    : folderName
      ? `Folder: ${folderName} (no notes file yet) — ${folderLink || ''}`
      : 'No notes found in Drive or memory';

  const snowflakeAttendees = (event.attendees || []).filter(a =>
    a.email && a.email.toLowerCase().endsWith('@snowflake.com') &&
    !a.email.toLowerCase().includes('resource.calendar')
  );

  // Derive meeting topic from title
  const topicFromTitle = (() => {
    const t = title.replace(/^[^-|—]+[-|—]\s*/, '').trim();
    return t && t !== title ? t : title;
  })();

  // Compute duration for agenda timing hints
  const durationMin = (() => {
    if (event.start && event.start.dateTime && event.end && event.end.dateTime) {
      return Math.round((new Date(event.end.dateTime) - new Date(event.start.dateTime)) / 60000);
    }
    return 60;
  })();

  const agendaSteps = buildDynamicAgenda(event, topicFromTitle, externalAttendees, contactTitles, aiParsed, notesContent, durationMin);

  const assumptions = [companyInfo.assumption].filter(Boolean);
  if (!notesContent && noteResult && noteResult.method === 'folder-empty') {
    assumptions.push(`[NOTE: Account folder "${folderName}" found in Drive but contains no notes files yet]`);
  }
  if (!noteResult && !notesContent) {
    assumptions.push(`[NOTE: No notes found for this account — prep is based on meeting title/attendees only]`);
  }

  // ─── PLAIN TEXT (calendar description) ──────────────────────────────────────
  const extRows = externalAttendees.map(a => {
    const ct = (contactTitles || {})[a.email ? a.email.toLowerCase() : ''];
    const name = (ct && ct.name) ? ct.name : (inferName(a) || a.email);
    const titleStr = (ct && ct.title) ? ct.title : '?';
    return `  ${name} | ${titleStr} | ${a.email} | ${companyInfo.displayName}`;
  }).join('\n') || '  (none found)';

  const sfRows = snowflakeAttendees.map(a => {
    const n = inferName(a) || a.email;
    return `  ${n} | ${a.email}`;
  }).join('\n') || '  (just you)';

  const plainText = [
    `=== PRE-CALL PREP (auto-generated) ===`,
    `Meeting : ${title}`,
    `Account : ${companyInfo.displayName}`,
    `Topic   : ${topicFromTitle}`,
    `When    : ${dateStr}`,
    meetingLink ? `Link    : ${meetingLink}` : null,
    ``,
    `EXTERNAL ATTENDEES:`,
    extRows,
    ``,
    `SNOWFLAKE TEAM:`,
    sfRows,
    ``,
    `ACCOUNT CONTEXT:`,
    useCases ? `  USE CASES / FOCUS:\n${useCases.split('\n').map(l => '    ' + l).join('\n')}` : `  USE CASES / FOCUS: See notes`,
    ``,
    recentNotes ? `  RECENT SE NOTES:\n${recentNotes.split('\n').map(l => '    ' + l).join('\n')}` : (raw ? `  NOTES EXCERPT:\n${raw.split('\n').slice(0,10).map(l => '    ' + l).join('\n')}` : `  NOTES: —`),
    ``,
    `SUGGESTED PREP AGENDA:`,
    ...agendaSteps.map((s, i) => `  ${i + 1}. ${s}`),
    (agendaSteps.length === 0) ? `  1. Context recap / update (5 min)` : null,
    ``,
    assumptions.length ? `ASSUMPTIONS / NOTES:\n${assumptions.map(a => '  ' + a).join('\n')}` : null,
    ``,
    `SOURCE: ${sourceRef}`,
    `Prep generated: ${new Date().toLocaleString('en-CA', { dateStyle: 'medium', timeStyle: 'short' })}`,  
    `=== END PREP ===`
  ].filter(l => l !== null).join('\n');

  // ─── HTML (Gmail draft) ──────────────────────────────────────────────────────
  const htmlBody = `
<div style="font-family:Arial,sans-serif;max-width:720px;padding:20px;color:#222;">
  <h2 style="color:#29B5E8;border-bottom:2px solid #29B5E8;padding-bottom:8px;margin-top:0;">
    PRE-CALL PREP: ${escHtml(title)}
  </h2>

  <table style="margin-bottom:12px;border-collapse:collapse;">
    <tr><td style="padding:2px 12px 2px 0;color:#555;white-space:nowrap;"><strong>Account</strong></td><td>${escHtml(companyInfo.displayName)}</td></tr>
    <tr><td style="padding:2px 12px 2px 0;color:#555;white-space:nowrap;"><strong>Topic</strong></td><td>${escHtml(topicFromTitle)}</td></tr>
    <tr><td style="padding:2px 12px 2px 0;color:#555;white-space:nowrap;"><strong>Date/Time</strong></td><td>${escHtml(dateStr)}</td></tr>
    ${meetingLink ? `<tr><td style="padding:2px 12px 2px 0;color:#555;"><strong>Link</strong></td><td><a href="${escHtml(meetingLink)}">${escHtml(meetingLink)}</a></td></tr>` : ''}
  </table>

  ${originalDesc ? `
  <h3 style="color:#444;margin-bottom:4px;">Meeting Request / Existing Agenda</h3>
  <pre style="background:#fffbea;padding:10px 14px;border-left:4px solid #f0c040;white-space:pre-wrap;font-size:13px;margin-top:0;">${escHtml(originalDesc)}</pre>` : ''}

  <h3 style="color:#444;margin-bottom:6px;">External Attendees</h3>
  <table style="border-collapse:collapse;width:100%;margin-bottom:16px;">
    <tr style="background:#f0f8ff;">
      <th style="border:1px solid #ddd;padding:6px 10px;text-align:left;">Name</th>
      <th style="border:1px solid #ddd;padding:6px 10px;text-align:left;">Title</th>
      <th style="border:1px solid #ddd;padding:6px 10px;text-align:left;">Email</th>
      <th style="border:1px solid #ddd;padding:6px 10px;text-align:left;">Company</th>
    </tr>
    ${externalAttendees.map(a => {
      const ct = (contactTitles || {})[a.email ? a.email.toLowerCase() : ''];
      const name = (ct && ct.name) ? ct.name : (inferName(a) || a.email || '—');
      const titleCell = (ct && ct.title) ? ct.title : '—';
      return `<tr>
        <td style="border:1px solid #ddd;padding:6px 10px;">${escHtml(name)}</td>
        <td style="border:1px solid #ddd;padding:6px 10px;">${escHtml(titleCell)}</td>
        <td style="border:1px solid #ddd;padding:6px 10px;">${escHtml(a.email || '—')}</td>
        <td style="border:1px solid #ddd;padding:6px 10px;">${escHtml(companyInfo.displayName)}</td>
      </tr>`;
    }).join('') || '<tr><td colspan="4" style="border:1px solid #ddd;padding:6px 10px;color:#999;">No external attendees found</td></tr>'}
  </table>

  <h3 style="color:#444;margin-bottom:4px;">Snowflake Team</h3>
  <p style="margin-top:4px;">${snowflakeAttendees.map(a => escHtml(inferName(a) || a.email)).join(', ') || '(just you)'}</p>

  ${notesContent ? `
  <h3 style="color:#444;margin-bottom:4px;">Account Context</h3>
  ${useCases ? `
  <p style="margin:4px 0 2px;"><strong>Use Cases / Focus:</strong></p>
  <pre style="background:#f0faf0;padding:10px 14px;border-left:4px solid #4CAF50;white-space:pre-wrap;font-size:13px;margin-top:0;">${escHtml(useCases)}</pre>` : ''}
  ${recentNotes ? `
  <p style="margin:8px 0 2px;"><strong>Recent SE Notes:</strong></p>
  <pre style="background:#f9f9f9;padding:10px 14px;border-left:4px solid #29B5E8;white-space:pre-wrap;font-size:13px;margin-top:0;">${escHtml(recentNotes)}</pre>` : ''}
  ${!useCases && !recentNotes && raw ? `
  <pre style="background:#f9f9f9;padding:10px 14px;border-left:4px solid #29B5E8;white-space:pre-wrap;font-size:13px;">${escHtml(raw)}</pre>` : ''}
  ` : `
  <h3 style="color:#444;margin-bottom:4px;">Account Context</h3>
  <p style="color:#999;font-style:italic;">${folderName ? `Account folder "${escHtml(folderName)}" found in Drive — no notes file yet. <a href="${escHtml(folderLink || '')}">Open folder</a>` : 'No notes found in Drive or memory. Check SFDC or search Drive manually.'}</p>
  `}

  <h3 style="color:#444;margin-bottom:4px;">Suggested Prep Agenda</h3>
  <ol style="margin-top:4px;">
    ${agendaSteps.map(s => `<li>${escHtml(s)}</li>`).join('\n    ')}
  </ol>

  ${assumptions.length ? `
  <div style="background:#fff8e1;border-left:4px solid #FFC107;padding:10px 14px;margin:16px 0;font-size:13px;">
    <strong>Assumptions / Notes:</strong><br>
    ${assumptions.map(a => escHtml(a)).join('<br>')}
  </div>` : ''}

  <hr style="border:none;border-top:1px solid #eee;margin:20px 0;">
  <p style="color:#aaa;font-size:11px;">
    Source: ${escHtml(sourceRef)}<br>
    Generated by se-meeting-prep skill &mdash; ${new Date().toLocaleDateString('en-CA')}
  </p>
</div>`;

  return { plainText, htmlBody, title, dateStr };
}

// ─── CONSOLIDATED SUMMARY EMAIL ────────────────────────────────────────────────
function buildConsolidatedEmail(weekLbl, meetingDataArr, aiParsedArr, contactTitles) {
  const hasNone = (s) => !s || /^•\s*none\.?$/i.test(s.trim());

  const cards = meetingDataArr.map((m, i) => {
    const { event, type, externalAttendees, companyInfo, noteResult, notesContent } = m;
    const aiParsed = aiParsedArr[i];
    const parsed = notesContent ? parseNotesContent(notesContent) : null;
    const recentNotes = (aiParsed && aiParsed.recentNotes) || (parsed && parsed.recentNotes);
    const useCases    = (aiParsed && aiParsed.useCases)    || (parsed && parsed.useCases);
    const actionItems = aiParsed && !hasNone(aiParsed.actionItems) ? aiParsed.actionItems : null;

    const isVip = type.includes('vip');
    const title = event.summary || 'Untitled Meeting';
    const timeStr = formatDateTime(event);
    const folderName = noteResult && noteResult.folderName;
    const folderLink = noteResult && noteResult.folderLink;

    const sfNames = (event.attendees || [])
      .filter(a => a.email && a.email.toLowerCase().endsWith('@snowflake.com') && !a.email.includes('resource.calendar'))
      .map(a => escHtml(inferName(a) || a.email)).join(', ') || 'you';

    const extNames = externalAttendees.map(a => {
      const ct = (contactTitles || {})[a.email ? a.email.toLowerCase() : ''];
      const name = (ct && ct.name) ? ct.name : (inferName(a) || a.email || '?');
      const titleStr = (ct && ct.title) ? ` <span style="color:#888;font-size:12px">(${escHtml(ct.title)})</span>` : '';
      return `${escHtml(name)}${titleStr}`;
    }).join(', ') || '<span style="color:#aaa">No external attendees</span>';

    const borderColor = isVip ? '#0070D2' : '#aaa';
    const label = isVip ? `<span style="color:#0070D2;font-weight:700">★ </span>` : '';

    const notesHtml = recentNotes
      ? `<div style="margin:6px 0 10px">
           <div style="font-size:11px;font-weight:700;color:#555;text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px">Recent Notes</div>
           <pre style="margin:0;padding:8px 12px;background:#f7f9fc;border-left:3px solid #29B5E8;white-space:pre-wrap;font-size:13px;font-family:inherit">${escHtml(recentNotes)}</pre>
         </div>`
      : notesContent ? '' : `<p style="color:#aaa;font-style:italic;font-size:13px;margin:6px 0">${folderName ? `Folder found ("${escHtml(folderName)}") — no notes file yet.` : 'No notes found.'}</p>`;

    const useCasesHtml = useCases
      ? `<div style="margin:0 0 10px">
           <div style="font-size:11px;font-weight:700;color:#555;text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px">Use Cases</div>
           <pre style="margin:0;padding:8px 12px;background:#f0faf0;border-left:3px solid #4CAF50;white-space:pre-wrap;font-size:13px;font-family:inherit">${escHtml(useCases)}</pre>
         </div>` : '';

    const actionHtml = actionItems
      ? `<div style="margin:0 0 6px">
           <div style="font-size:11px;font-weight:700;color:#d97706;text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px">⚡ Action Items</div>
           <pre style="margin:0;padding:8px 12px;background:#fffbea;border-left:3px solid #f0c040;white-space:pre-wrap;font-size:13px;font-family:inherit">${escHtml(actionItems)}</pre>
         </div>` : '';

    return `
<div style="border-left:4px solid ${borderColor};padding:12px 16px;margin-bottom:20px;background:#fafafa;border-radius:0 6px 6px 0">
  <div style="margin-bottom:6px">
    <span style="font-size:15px;font-weight:700;color:#111">${label}${escHtml(title)}</span>
  </div>
  <table style="border-collapse:collapse;font-size:12px;color:#666;margin-bottom:8px">
    <tr><td style="padding:1px 12px 1px 0;white-space:nowrap"><strong>Account</strong></td><td>${escHtml(companyInfo.displayName)}</td></tr>
    <tr><td style="padding:1px 12px 1px 0;white-space:nowrap"><strong>When</strong></td><td>${escHtml(timeStr)}</td></tr>
    <tr><td style="padding:1px 12px 1px 0;white-space:nowrap"><strong>With</strong></td><td>${extNames}</td></tr>
    <tr><td style="padding:1px 12px 1px 0;white-space:nowrap"><strong>SF Team</strong></td><td>${sfNames}</td></tr>
  </table>
  ${notesHtml}${useCasesHtml}${actionHtml}
</div>`;
  }).join('');

  return `<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;max-width:720px;padding:20px;color:#222">
  <h2 style="color:#0070D2;border-bottom:2px solid #0070D2;padding-bottom:8px;margin-top:0">Pre-call Prep Summary</h2>
  <p style="color:#666;margin:0 0 24px;font-size:13px">${escHtml(weekLbl)} &nbsp;·&nbsp; Generated ${new Date().toLocaleDateString('en-CA')}<br>
  <em style="font-size:12px">★ = Ali/Shaw meeting &nbsp;·&nbsp; Sorted highest to lowest priority &nbsp;·&nbsp; Full agendas written to each calendar event</em></p>
  ${cards}
</div>`;
}

// ─── GMAIL DRAFT ───────────────────────────────────────────────────────────────
function buildMimeRaw(subject, htmlBody) {
  const boundary = 'BOUNDARY_' + Date.now();
  const mime = [
    `To: ${ME}`,
    `From: Carson Walker <${ME}>`,
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
  return sendEmailViaSMTP(subject, htmlBody, token);
}

function sendEmailViaSMTP(subject, htmlBody, token) {
  return new Promise((resolve, reject) => {
    const tls = require('tls');
    const authStr = Buffer.from(`user=${ME}\x01auth=Bearer ${token}\x01\x01`).toString('base64');
    const boundary = 'BOUNDARY_' + Date.now();
    const message = [
      `From: Carson Walker <${ME}>`,
      `To: ${ME}`,
      `Subject: ${subject}`,
      'MIME-Version: 1.0',
      `Content-Type: multipart/alternative; boundary="${boundary}"`,
      '',
      `--${boundary}`,
      'Content-Type: text/html; charset=UTF-8',
      '',
      htmlBody,
      '',
      `--${boundary}--`,
      ''
    ].join('\r\n');

    const sock = tls.connect(465, 'smtp.gmail.com', { servername: 'smtp.gmail.com' }, () => {
      let buf = '';
      let step = 0;
      const send = (line) => sock.write(line + '\r\n');

      sock.on('data', (d) => {
        buf += d.toString();
        const lines = buf.split('\r\n');
        buf = lines.pop();
        for (const line of lines) {
          if (step === 0 && line.startsWith('220')) { send('EHLO localhost'); step = 1; }
          else if (step === 1 && line.startsWith('250') && !line.startsWith('250-')) { send('AUTH XOAUTH2 ' + authStr); step = 2; }
          else if (step === 2 && line.startsWith('235')) { send(`MAIL FROM:<${ME}>`); step = 3; }
          else if (step === 3 && line.startsWith('250')) { send(`RCPT TO:<${ME}>`); step = 4; }
          else if (step === 4 && line.startsWith('250')) { send('DATA'); step = 5; }
          else if (step === 5 && line.startsWith('354')) { sock.write(message + '\r\n.\r\n'); step = 6; }
          else if (step === 6 && line.startsWith('250')) { send('QUIT'); resolve({ id: 'smtp-sent', line }); step = 7; }
          else if (line.startsWith('4') || line.startsWith('5')) { reject(new Error('SMTP error: ' + line)); }
        }
      });
      sock.on('error', reject);
    });
    sock.on('error', reject);
  });
}

// ─── CALENDAR UPDATE (EventKit Swift binary) ───────────────────────────────────
function updateCalendarDescriptionLocal(eventId, _calName, currentDesc, newPrepText) {
  const original = currentDesc
    ? currentDesc.replace(/=== PRE-CALL PREP[\s\S]*?=== END PREP ===/g, '').trim()
    : '';
  const newDesc = (original ? original + '\n\n' : '') + newPrepText;
  const bin = path.join(SKILL_DIR, 'update_cal_event');
  const result = spawnSync(bin, [eventId], { input: newDesc, encoding: 'utf8', timeout: 15000 });
  if (result.error) return { status: 500, error: result.error.message };
  if (result.status !== 0) return { status: 404, error: result.stderr };
  return { status: 200 };
}

// ─── WEEK LABEL ────────────────────────────────────────────────────────────────
function weekLabel(weekOffset) {
  const now = new Date();
  const day = now.getDay();
  const daysToMon = day === 0 ? -6 : -(day - 1);
  const monday = new Date(now);
  monday.setDate(now.getDate() + daysToMon + (weekOffset || 0));
  return monday.toLocaleDateString('en-CA', { month: 'short', day: 'numeric' }) + ' week';
}

// ─── MAIN ──────────────────────────────────────────────────────────────────────
async function main(weekOffset) {
  const label = weekOffset > 0 ? 'NEXT WEEK' : 'THIS WEEK';
  console.log(`\nSE Meeting Prep - ${weekLabel(weekOffset)} (${label})`);
  console.log('='.repeat(60));

  const token = await getAccessToken();
  console.log('OAuth token loaded');

  const allEvents = fetchCalendarEvents(weekOffset);
  console.log(`Found ${allEvents.length} events`);

  const sorted = sortEvents(allEvents);
  console.log(`Processing ${sorted.length} external/VIP meetings\n`);

  if (!sorted.length) {
    console.log('No external meetings found. Nothing to prep!');
    return;
  }

  // Batch SFDC contact lookup
  const allExternalEmails = [...new Set(
    sorted.flatMap(({ event }) =>
      (event.attendees || [])
        .filter(a => a.email && !a.email.toLowerCase().endsWith('@snowflake.com') && !a.email.includes('resource.calendar'))
        .map(a => a.email.toLowerCase())
    )
  )];
  const contactTitles = lookupContactTitles(allExternalEmails);
  console.log(`SFDC titles loaded: ${Object.keys(contactTitles).length}/${allExternalEmails.length} contacts matched\n`);

  const prepState = loadPrepState();
  const results = [];

  // ── Phase 1: Gather all meeting data + classify new/changed/current ─────────
  console.log('Phase 1: Gathering meeting data...\n');
  const meetingDataArr = [];

  for (const { event, type } of sorted) {
    const title = event.summary || 'Untitled Meeting';
    const label2 = type.includes('vip') ? '[Ali/Shaw]' : '[Customer]';
    console.log(`Processing: ${title} ${label2}`);

    const externalAttendees = (event.attendees || []).filter(a =>
      a.email && !a.email.toLowerCase().endsWith('@snowflake.com') && !a.email.includes('resource.calendar')
    );

    const companyInfo = getCompanyInfo(event, externalAttendees);
    console.log(`  Account: ${companyInfo.displayName}${companyInfo.assumption ? ' *' : ''}`);
    if (companyInfo.assumption) console.log(`  ! ${companyInfo.assumption}`);

    // Determine if this meeting needs prep (new, changed, or no existing prep)
    const checksum = computeChecksum(event, externalAttendees);
    const stored = prepState[event.id];
    const hasPrep = (event.description || '').includes('=== PRE-CALL PREP');

    let meetingStatus;
    if (!stored || !hasPrep) {
      meetingStatus = 'new';
    } else if (stored.checksum !== checksum) {
      const reasons = [];
      if (stored.title !== (event.summary || '')) reasons.push('title renamed');
      if (stored.startTime !== ((event.start && event.start.dateTime) || '')) reasons.push('time moved');
      if (stored.attendees !== externalAttendees.map(a => a.email).sort().join(',')) reasons.push('attendees changed');
      meetingStatus = 'changed:' + (reasons.length ? reasons.join(', ') : 'details updated');
    } else {
      meetingStatus = 'current';
    }

    const needsUpdate = meetingStatus !== 'current';
    if (meetingStatus === 'current') {
      console.log(`  Status: up to date — skipping`);
    } else if (meetingStatus.startsWith('changed')) {
      console.log(`  Status: \x1b[33mupdated\x1b[0m (${meetingStatus.slice(8)}) — will re-prep`);
    } else {
      console.log(`  Status: \x1b[32mnew\x1b[0m — will prep`);
    }

    let noteResult = null;
    let notesContent = null;

    if (companyInfo.keywords.length > 0) {
      noteResult = await findDriveNotesForCompany(companyInfo.keywords, token);
      if (noteResult && noteResult.file) {
        console.log(`  Drive notes: ${noteResult.file.name} (via ${noteResult.method}${noteResult.folderName ? ` / folder: ${noteResult.folderName}` : ''})`);
        try { notesContent = await readDriveFile(noteResult.file, token); } catch (e) {
          console.log(`  Could not read Drive file: ${e.message}`);
        }
      } else if (noteResult && noteResult.folderName) {
        console.log(`  Account folder found: "${noteResult.folderName}" (no notes file)`);
      } else {
        const memFile = searchMemoryFiles(companyInfo.keywords);
        if (memFile) {
          console.log(`  Memory notes: ${memFile.filename}`);
          notesContent = memFile.content;
          noteResult = { file: null, folderName: memFile.filename, folderLink: null, method: 'memory' };
        } else {
          console.log(`  No notes found`);
        }
      }
    }

    meetingDataArr.push({ event, type, externalAttendees, companyInfo, noteResult, notesContent, needsUpdate, checksum, meetingStatus });
  }

  const toUpdateCount = meetingDataArr.filter(m => m.needsUpdate).length;
  const skippedCount = meetingDataArr.length - toUpdateCount;
  console.log(`\n${toUpdateCount} meeting(s) need prep, ${skippedCount} already up to date.`);

  if (toUpdateCount === 0) {
    console.log('Nothing to update — posting refreshed summary draft only.\n');
  }

  // ── Phase 2: Parallel Cortex summarization (new/changed meetings only) ──────
  const toSummarize = meetingDataArr.filter(m => m.needsUpdate && m.notesContent);
  console.log(`\nPhase 2: Cortex summarizing ${toSummarize.length} meeting(s) with notes (parallel, batches of 5)...`);
  const CORTEX_BATCH = 5;
  const aiParsedArr = new Array(meetingDataArr.length).fill(null);
  const noteIdxs = meetingDataArr.map((m, i) => (m.needsUpdate && m.notesContent) ? i : -1).filter(i => i >= 0);

  for (let b = 0; b < noteIdxs.length; b += CORTEX_BATCH) {
    const batchIdxs = noteIdxs.slice(b, b + CORTEX_BATCH);
    const batchNames = batchIdxs.map(i => meetingDataArr[i].companyInfo.displayName).join(', ');
    console.log(`  Batch ${Math.floor(b / CORTEX_BATCH) + 1}: ${batchNames}`);
    const batchResults = await Promise.all(
      batchIdxs.map(i => cortexSummarizeAsync(meetingDataArr[i].notesContent))
    );
    batchIdxs.forEach((i, j) => {
      aiParsedArr[i] = batchResults[j];
      console.log(`    ${meetingDataArr[i].companyInfo.displayName}: ${batchResults[j] ? 'done' : 'skipped (fallback to regex)'}`);
    });
  }

  // ── Phase 3: Update calendars (new/changed only), build consolidated draft ───
  console.log('\nPhase 3: Updating calendars and building summary draft...\n');

  for (let i = 0; i < meetingDataArr.length; i++) {
    const { event, type, externalAttendees, companyInfo, noteResult, notesContent, needsUpdate, checksum, meetingStatus } = meetingDataArr[i];
    const aiParsed = aiParsedArr[i];
    const title = event.summary || 'Untitled Meeting';
    const { plainText, htmlBody, dateStr } = buildAgenda(
      event, externalAttendees, companyInfo, noteResult, notesContent, contactTitles, aiParsed
    );

    if (!needsUpdate) {
      results.push({ title, dateStr, type, status: '[SKIPPED]', company: companyInfo.displayName, meetingStatus });
      continue;
    }

    let calOk = false, calSkipped = false;
    try {
      const currentDesc = event.description || '';
      const res = updateCalendarDescriptionLocal(event.id, event._calName, currentDesc, plainText);
      if (res.status === 200) { calOk = true; console.log(`  ${companyInfo.displayName}: Calendar updated`); }
      else if (res.status === 404) { calSkipped = true; console.log(`  ${companyInfo.displayName}: Calendar skipped (not found)`); }
      else console.log(`  ${companyInfo.displayName}: Calendar update failed (${res.status}):`, res.error || '');
    } catch (e) { console.log(`  ${companyInfo.displayName}: Calendar error: ${e.message}`); }

    // Save to state on success
    if (calOk || calSkipped) {
      prepState[event.id] = {
        checksum,
        processedAt: new Date().toISOString(),
        title: event.summary || '',
        startTime: (event.start && event.start.dateTime) || '',
        attendees: externalAttendees.map(a => a.email).sort().join(',')
      };
    }

    const statusLabel = meetingStatus === 'new' ? '[NEW]' : meetingStatus.startsWith('changed') ? '[UPDATED]' : '[OK]';
    const calStatus = calOk ? '' : calSkipped ? ' (Gmail only)' : ' (cal failed)';
    results.push({ title, dateStr, type, status: statusLabel + calStatus, company: companyInfo.displayName, meetingStatus });
  }

  savePrepState(prepState);

  // Post consolidated summary draft (only meetings at or after run time)
  const updatedCount = results.filter(r => r.status !== '[SKIPPED]').length;
  console.log(`\nPosting consolidated summary draft (${updatedCount} updated, ${skippedCount} unchanged)...`);
  const wLabel = weekLabel(weekOffset);
  const now = Date.now();
  const futureMeetingIdxs = meetingDataArr.map((m, i) => {
    const start = m.event.start && m.event.start.dateTime ? new Date(m.event.start.dateTime).getTime() : 0;
    return start >= now ? i : -1;
  }).filter(i => i >= 0);
  const futureMeetings = futureMeetingIdxs.map(i => meetingDataArr[i]);
  const futureAiParsed = futureMeetingIdxs.map(i => aiParsedArr[i]);
  const consolidatedHtml = buildConsolidatedEmail(wLabel, futureMeetings, futureAiParsed, contactTitles);
  const summarySubject = `Pre-call Prep Summary - ${wLabel}`;
  try {
    const draft = await createGmailDraft(summarySubject, consolidatedHtml, token);
    if (draft.id) console.log(`  Summary email sent to inbox: ${draft.id}`);
    else console.log('  Summary email failed:', JSON.stringify(draft).substring(0, 80));
  } catch (e) { console.log('  Summary draft error:', e.message); }

  console.log('\n' + '='.repeat(60));
  console.log('SUMMARY');
  console.log('='.repeat(60));
  for (const r of results) {
    const star = r.type.includes('vip') ? '* ' : '  ';
    const change = r.meetingStatus && r.meetingStatus.startsWith('changed:') ? ` (${r.meetingStatus.slice(8)})` : '';
    console.log(`${star}${r.status}${change} ${r.title} — ${r.company} (${r.dateStr})`);
  }
  console.log('\n* = Ali/Shaw meeting (prioritized)');
  console.log('[NEW] = first time prepped   [UPDATED] = meeting changed   [SKIPPED] = no changes since last run');
  console.log('Drafts: Gmail → Drafts folder');
  console.log('Calendar: Open each event to see the updated description');
}

// ─── ENTRY POINT ───────────────────────────────────────────────────────────────
(async () => {
  try {
    await main(0);

    // Ask about next week
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question('\nRun prep for NEXT WEEK too? (y/n): ', async (answer) => {
      rl.close();
      if (answer.trim().toLowerCase() === 'y') {
        try { await main(7); }
        catch (err) { console.error('\nERROR (next week):', err.message); }
      }
      console.log('\nDone.');
    });
  } catch (err) {
    console.error('\nERROR:', err.message);
    if (err.message.includes('Calendar')) {
      console.error('Fix: Ensure carson.walker@snowflake.com is in System Settings > Internet Accounts with Calendars enabled');
    }
    process.exit(1);
  }
})();
