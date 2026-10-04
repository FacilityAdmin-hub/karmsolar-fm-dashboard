// Scheduled sync: authenticates as the app itself (no user sign-in needed),
// downloads FM Tickets.xlsx from Microsoft Graph, extracts the Tickets table,
// and bakes it into ticket-board.html's EMBEDDED_TICKETS constant so every
// visitor sees current data with zero sign-in — same pattern as index.html.
//
// NOTE: this script used to also mirror tickets straight into index.html's
// TICKET_ISSUES constant. That was removed — sync.js (the "Sync dashboard from
// Microsoft Graph" workflow) now owns index.html exclusively and reads the
// EMBEDDED_TICKETS this script writes below to do that mirroring itself. Two
// workflows writing index.html independently raced: sync.js runs on a very
// frequent external trigger and resolves push conflicts with `-X ours`, so it
// was silently discarding whatever ticket data this script had just committed.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const USER_UPN = process.env.GRAPH_USER_UPN || 'mohi.mohsen@karmsolar.com';
const FILE_PATH = process.env.GRAPH_TICKETS_FILE_PATH || '/FM Tickets.xlsx';
const BOARD_HTML_PATH = process.env.BOARD_HTML_PATH || path.join(__dirname, '..', 'ticket-board.html');

const COLS = [
  'Ticket ID','Date Submitted','Title','Description','Category',
  'Priority','Site / Location','Requester Name','Requester Email',
  'Status','Assigned To','Date Updated','Resolution Notes'
];

const COMMENT_COLS = ['Ticket ID','Timestamp','Type','Author','Content','AttachmentURL'];

async function getAppToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Token request failed: ${res.status} ${text}`);
  }
  const json = await res.json();
  return json.access_token;
}

async function downloadWorkbook(token) {
  const encodedPath = encodeURI(FILE_PATH);
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(USER_UPN)}/drive/root:${encodedPath}:/content`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph file download failed: ${res.status} ${text}`);
  }
  const arrayBuf = await res.arrayBuffer();
  return Buffer.from(arrayBuf);
}

function extractTickets(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  const sname = wb.SheetNames.find(n => n.replace(/\s+/g, '').toLowerCase() === 'tickets') || wb.SheetNames[0];
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sname], { header: 1, raw: true, blankrows: false, defval: '' });
  if (!aoa.length) return [];
  const rows = [];
  for (let i = 1; i < aoa.length; i++) {
    const raw = aoa[i] || [];
    const id = String(raw[0] || '').trim();
    if (!id) continue;
    const obj = {};
    COLS.forEach((label, ci) => {
      let v = raw[ci];
      if (v == null) v = '';
      // dates may come through as Excel serials; stringify plainly, sync doesn't need to parse them further
      if ((ci === 1 || ci === 11) && (typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))))) { const ud = Math.floor(v - 25569) * 86400; const d = new Date(ud * 1000); const fd = v - Math.floor(v); const ts = Math.floor(86400 * fd + 0.5); const hh = Math.floor(ts/3600), mm = Math.floor(ts/60)%60; const pad = n => String(n).padStart(2,'0'); v = `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())} ${pad(hh)}:${pad(mm)}`; } obj[label] = v;
    });
    rows.push(obj);
  }
  return rows;
}

function patchEmbedded(htmlContent, tickets) {
  const marker = 'const EMBEDDED_TICKETS = ';
  const startIdx = htmlContent.indexOf(marker);
  if (startIdx < 0) throw new Error('Could not find "const EMBEDDED_TICKETS = " in ticket-board.html — has the file structure changed?');
  const jsonStart = startIdx + marker.length;
  const endMarker = '];\n';
  const endIdx = htmlContent.indexOf(endMarker, jsonStart);
  if (endIdx < 0) throw new Error('Could not find the end of the EMBEDDED_TICKETS array in ticket-board.html');
  const before = htmlContent.slice(0, jsonStart);
  const after = htmlContent.slice(endIdx + 1); // "]" consumed by JSON.stringify, keep ";\n" + rest
  const newJson = JSON.stringify(tickets);
  return before + newJson + after;
}

function patchEmbeddedComments(htmlContent, commentsByTicket) {
  const marker = 'const EMBEDDED_COMMENTS = ';
  const startIdx = htmlContent.indexOf(marker);
  if (startIdx < 0) throw new Error('Could not find "const EMBEDDED_COMMENTS = " in ticket-board.html — has the file structure changed?');
  const jsonStart = startIdx + marker.length;
  const endMarker = '};\n';
  const endIdx = htmlContent.indexOf(endMarker, jsonStart);
  if (endIdx < 0) throw new Error('Could not find the end of the EMBEDDED_COMMENTS object in ticket-board.html');
  const before = htmlContent.slice(0, jsonStart);
  const after = htmlContent.slice(endIdx + 1); // "}" consumed by JSON.stringify, keep ";\n" + rest
  const newJson = JSON.stringify(commentsByTicket);
  return before + newJson + after;
}

function excelSerialToStamp(v) {
  const ud = Math.floor(v - 25569) * 86400;
  const d = new Date(ud * 1000);
  const fd = v - Math.floor(v);
  const ts = Math.floor(86400 * fd + 0.5);
  const hh = Math.floor(ts / 3600), mm = Math.floor(ts / 60) % 60;
  const pad = n => String(n).padStart(2, '0');
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ${pad(hh)}:${pad(mm)}`;
}

function extractComments(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  const sname = wb.SheetNames.find(n => n.replace(/\s+/g, '').toLowerCase() === 'comments');
  if (!sname) return {};
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sname], { header: 1, raw: true, blankrows: false, defval: '' });
  if (!aoa.length) return {};
  const grouped = {};
  for (let i = 1; i < aoa.length; i++) {
    const raw = aoa[i] || [];
    const ticketId = String(raw[0] || '').trim();
    if (!ticketId) continue;
    const obj = {};
    COMMENT_COLS.forEach((label, ci) => {
      let v = raw[ci];
      if (v == null) v = '';
      if (ci === 1 && typeof v === 'number') v = excelSerialToStamp(v);
      obj[label] = v;
    });
    if (!grouped[ticketId]) grouped[ticketId] = [];
    grouped[ticketId].push(obj);
  }
  Object.values(grouped).forEach(arr => arr.sort((a, b) => String(a['Timestamp']).localeCompare(String(b['Timestamp']))));
  return grouped;
}

function patchSourceLabel(htmlContent) {
  const re = /(<span id="syncLabel">)[^<]*(<\/span>)/;
  if (!re.test(htmlContent)) throw new Error('Could not find <span id="syncLabel"> in ticket-board.html — has the file structure changed?');
  const stamp = new Date().toLocaleString('en-US', { timeZone: 'UTC' }) + ' UTC';
  const label = `Auto-synced · ${stamp}`;
  return htmlContent.replace(re, `$1${label}$2`);
}

// ---- New-comment notifications -------------------------------------------------
// Comments are posted through the Cloudflare Worker straight into the "Comments"
// sheet of FM Tickets.xlsx. Each sync run compares that sheet against the comments
// already baked into ticket-board.html; anything not there yet is a NEW comment (or
// attachment) and is emailed to the FM team. Status-change rows are skipped because
// update-status.js already emails those. This also catches comments typed directly
// into Excel. Override recipients with a comma-separated NOTIFY_EMAILS env var.
const NOTIFY_EMAILS = (process.env.NOTIFY_EMAILS || 'mostafa.abuelmagd@karmsolar.com,hussien.magdy@karmsolar.com,mohi.mohsen@karmsolar.com')
  .split(',').map(s => s.trim()).filter(Boolean);
const MAX_COMMENT_EMAILS_PER_RUN = 10;

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function readExistingComments(htmlContent) {
  const marker = 'const EMBEDDED_COMMENTS = ';
  const startIdx = htmlContent.indexOf(marker);
  if (startIdx < 0) return null;
  const jsonStart = startIdx + marker.length;
  const endIdx = htmlContent.indexOf('};\n', jsonStart);
  if (endIdx < 0) return null;
  try { return JSON.parse(htmlContent.slice(jsonStart, endIdx + 1)); } catch (e) { return null; }
}

function commentKey(c) {
  return [c['Ticket ID'], c['Timestamp'], c['Author'], c['Type'], c['Content'], c['AttachmentURL']].join('|');
}

function findNewComments(oldComments, newComments) {
  // No readable baseline means we can't tell old from new — never risk a flood.
  if (!oldComments) return [];
  const seen = new Set();
  Object.values(oldComments).forEach(arr => (arr || []).forEach(c => seen.add(commentKey(c))));
  const fresh = [];
  Object.values(newComments).forEach(arr => (arr || []).forEach(c => {
    if (c['Type'] === 'Status') return;
    if (!seen.has(commentKey(c))) fresh.push(c);
  }));
  return fresh;
}

async function sendTeamEmail(token, subject, htmlBody) {
  if (!NOTIFY_EMAILS.length) return;
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(USER_UPN)}/sendMail`;
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: {
          subject,
          body: { contentType: 'HTML', content: htmlBody },
          toRecipients: NOTIFY_EMAILS.map(address => ({ emailAddress: { address } })),
        },
        saveToSentItems: true,
      }),
    });
    if (!res.ok) console.log('Comment email failed:', res.status, await res.text());
  } catch (err) {
    console.log('Comment email failed:', err.message);
  }
}

async function notifyNewComments(token, freshComments, tickets) {
  const titleById = {};
  tickets.forEach(t => { titleById[t['Ticket ID']] = t['Title']; });
  const batch = freshComments.slice(0, MAX_COMMENT_EMAILS_PER_RUN);
  for (const c of batch) {
    const id = c['Ticket ID'];
    const isAttachment = c['Type'] !== 'Comment' && c['AttachmentURL'];
    const what = isAttachment ? 'New attachment' : 'New comment';
    const body = isAttachment
      ? `<p><a href="${esc(c['AttachmentURL'])}">${esc(c['Content'] || 'View attachment')}</a></p>`
      : `<p>${esc(c['Content']).replace(/\n/g, '<br>')}</p>`;
    await sendTeamEmail(
      token,
      `${what} on ${id}${titleById[id] ? ' — ' + titleById[id] : ''}`,
      `<p><b>${esc(c['Author'] || 'Someone')}</b> added ${isAttachment ? 'an attachment' : 'a comment'} on ticket
<b>${esc(id)}</b>${titleById[id] ? ' (' + esc(titleById[id]) + ')' : ''} at ${esc(c['Timestamp'])}.</p>
${body}
<p><a href="https://facilityadmin-hub.github.io/karmsolar-fm-dashboard/ticket-board.html">Open the ticket board</a></p>`
    );
  }
  if (freshComments.length > batch.length) {
    await sendTeamEmail(
      token,
      `${freshComments.length - batch.length} more new ticket comments`,
      `<p>${freshComments.length - batch.length} additional new comments/attachments arrived in the same sync and were not emailed individually. See the
<a href="https://facilityadmin-hub.github.io/karmsolar-fm-dashboard/ticket-board.html">ticket board</a>.</p>`
    );
  }
}

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Missing AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET environment variables.');
  }
  console.log('Authenticating with Microsoft Graph (app-only)...');
  const token = await getAppToken();
  console.log('Got token. Downloading workbook from', FILE_PATH, 'for', USER_UPN, '...');
  const buf = await downloadWorkbook(token);
  console.log(`Downloaded ${(buf.length / 1024).toFixed(1)} KB. Extracting Tickets table...`);
  const tickets = extractTickets(buf);
  console.log('Extracted', tickets.length, 'tickets.');
  const comments = extractComments(buf);
  console.log('Extracted comments/attachments for', Object.keys(comments).length, 'tickets.');

  console.log('Reading', BOARD_HTML_PATH, '...');
  const html = fs.readFileSync(BOARD_HTML_PATH, 'utf-8');
  const previousComments = readExistingComments(html);
  let patched = patchEmbedded(html, tickets);
  patched = patchEmbeddedComments(patched, comments);
  patched = patchSourceLabel(patched);
  fs.writeFileSync(BOARD_HTML_PATH, patched, 'utf-8');
  console.log('ticket-board.html updated with fresh data and current label.');

  // Email the FM team about comments/attachments that weren't in the previous snapshot.
  // Runs after the board file is written; never fails the sync itself.
  try {
    const fresh = findNewComments(previousComments, comments);
    if (fresh.length) {
      console.log('Found', fresh.length, 'new comment(s)/attachment(s) — emailing the team...');
      await notifyNewComments(token, fresh, tickets);
    } else {
      console.log('No new comments.');
    }
  } catch (err) {
    console.log('New-comment notification step failed (non-fatal):', err.message);
  }
}

module.exports = { patchEmbedded, patchEmbeddedComments, extractComments, patchSourceLabel, extractTickets, getAppToken, downloadWorkbook, readExistingComments, findNewComments };

if (require.main === module) {
  main().catch(err => {
    console.error('Ticket sync failed:', err.message);
    process.exit(1);
  });
}
