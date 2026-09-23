// Scheduled sync: authenticates as the app itself (no user sign-in needed),
// downloads the live FM Dashboard.xlsx from Microsoft Graph, extracts it into
// the same DS shape the browser computes, and bakes it into index.html's
// EMBEDDED constant so every visitor sees current data with zero sign-in.
const fs = require('fs');
const path = require('path');
const { extractAll } = require('./extract.js');

const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const USER_UPN = process.env.GRAPH_USER_UPN || 'mohi.mohsen@karmsolar.com';
const FILE_PATH = process.env.GRAPH_FILE_PATH || '/Main dashboard/FM Dashboard.xlsx';
const INDEX_HTML_PATH = process.env.INDEX_HTML_PATH || path.join(__dirname, '..', 'index.html');
// index.html's HQ Issues & requests tab also mirrors the ticket board (a separate
// workbook/workflow — see sync-tickets.js). Both this workflow and ticket-sync.yml
// used to write index.html independently, which raced: this workflow runs on a very
// frequent external trigger and resolves push conflicts with `-X ours`, so whenever
// it collided with a ticket-board update it silently discarded the newer ticket data.
// Fix: index.html has exactly one writer (this script). It reads the ticket board's
// already-committed EMBEDDED_TICKETS straight out of the checked-out ticket-board.html
// (no extra Graph call needed) and mirrors it into TICKET_ISSUES itself, in the same
// commit as everything else.
const TICKET_BOARD_HTML_PATH = process.env.TICKET_BOARD_HTML_PATH || path.join(__dirname, '..', 'ticket-board.html');

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

function patchIndexHtml(htmlContent, ds) {
  const marker = 'const EMBEDDED = ';
  const startIdx = htmlContent.indexOf(marker);
  if (startIdx < 0) throw new Error('Could not find "const EMBEDDED = " in index.html — has the file structure changed?');
  const jsonStart = startIdx + marker.length;
  // The EMBEDDED object is a single line ending in "};\n" — find that exact terminator.
  const endMarker = '};\n';
  const endIdx = htmlContent.indexOf(endMarker, jsonStart);
  if (endIdx < 0) throw new Error('Could not find the end of the EMBEDDED object in index.html');
  const before = htmlContent.slice(0, jsonStart);
  // endIdx points at the "}" that closes the OLD json object — skip past just that one
  // character so we don't double up with the "}" that JSON.stringify(ds) already provides.
  const after = htmlContent.slice(endIdx + 1); // ";\n..." (statement terminator + rest of file)
  const newJson = JSON.stringify(ds);
  return before + newJson + after;
}

// The "Data source" label near the top of the page is a hardcoded <b id="srcName"> in the
// static HTML — it's NOT driven by EMBEDDED/DS, so it never reflects freshness on its own
// for a first-time visitor. Patch it too, every sync, so the visible label matches reality.
function patchSourceLabel(htmlContent) {
  const re = /(<b id="srcName">)[^<]*(<\/b>)/;
  if (!re.test(htmlContent)) throw new Error('Could not find <b id="srcName"> in index.html — has the file structure changed?');
  const stamp = new Date().toLocaleString('en-US', { timeZone: 'UTC' }) + ' UTC';
  const label = `Live workbook (auto-synced) · ${stamp}`;
  return htmlContent.replace(re, `$1${label}$2`);
}

// HQ Issues lives in a separate top-level `let ISSUES=[];` (not part of EMBEDDED), because
// it needs richer upsert/merge behavior than a plain data blob. restoreState() only ever
// overwrites ISSUES from a visitor's own saved localStorage (st.issues) — it never reads
// this declaration itself after boot — so baking a fresh array into the initial value here
// is exactly as safe as EMBEDDED is for DS: a returning visitor's local edits and status
// history always take priority once restoreState() runs; only first-time visitors (with no
// saved state at all) see this baked snapshot.
//
// This can't be a simple regex up to the next "];" — issue titles/notes are free user text
// and could themselves contain "];". Instead we scan forward from "let ISSUES=" tracking
// JSON string/escape state and bracket depth, exactly like patchIndexHtml does for EMBEDDED,
// so we find the actual matching close-bracket regardless of what's inside.
function patchIssues(htmlContent, issues) {
  const marker = 'let ISSUES=';
  const start = htmlContent.indexOf(marker);
  if (start < 0) throw new Error('Could not find "let ISSUES=" in index.html — has the file structure changed?');
  const arrStart = start + marker.length; // position of the opening '['
  if (htmlContent[arrStart] !== '[') throw new Error('Expected "[" right after "let ISSUES=" — has the file structure changed?');
  let depth = 0, inStr = false, esc = false, i = arrStart;
  for (; i < htmlContent.length; i++) {
    const ch = htmlContent[i];
    if (inStr) {
      if (esc) esc = false;
      else if (ch === '\\') esc = true;
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') inStr = true;
    else if (ch === '[') depth++;
    else if (ch === ']') { depth--; if (depth === 0) { i++; break; } }
  }
  // i is now just past the matching ']'; expect a ';' right after
  if (htmlContent[i] !== ';') throw new Error('Could not find closing "];" for ISSUES array — has the file structure changed?');
  const before = htmlContent.slice(0, arrStart);
  const after = htmlContent.slice(i); // starts with ';'
  return before + JSON.stringify(issues) + after;
}

// Maps the ticket board's 8-stage lifecycle onto the HQ Issues & requests
// tab's 4-stage model, so tickets slot into the same status buckets used
// for hand-added / sheet-imported requests there. Kept identical to the
// mapping in sync-tickets.js.
const HQ_STATUS_MAP = {
  'Backlog': 'Awaiting approval',
  'Open': 'Awaiting approval',
  'Waiting Approval': 'Awaiting approval',
  'In Progress': 'In progress',
  'In Review': 'In progress',
  'Blocked': 'In progress',
  'Resolved': 'Done',
  'Closed': 'Done',
};

// Reads the ticket board's already-committed EMBEDDED_TICKETS straight out of the
// checked-out ticket-board.html and reshapes it into the lightweight records
// index.html's mergeTicketIssues() expects. Returns null (never throws) if the file
// or marker isn't there yet, so a ticket-board hiccup never fails the dashboard sync.
function readTicketIssuesFromBoard(boardHtmlPath) {
  if (!fs.existsSync(boardHtmlPath)) return null;
  const html = fs.readFileSync(boardHtmlPath, 'utf-8');
  const marker = 'const EMBEDDED_TICKETS = ';
  const startIdx = html.indexOf(marker);
  if (startIdx < 0) return null;
  const jsonStart = startIdx + marker.length;
  const endMarker = '];\n';
  const endIdx = html.indexOf(endMarker, jsonStart);
  if (endIdx < 0) return null;
  let tickets;
  try {
    tickets = JSON.parse(html.slice(jsonStart, endIdx + 1));
  } catch (e) {
    return null;
  }
  return tickets.map(t => ({
    ref: t['Ticket ID'],
    title: t['Title'] || '',
    category: t['Category'] || 'General',
    loc: t['Site / Location'] || 'HQ',
    rawStatus: t['Status'] || '',
    status: HQ_STATUS_MAP[t['Status']] || 'Awaiting approval',
    priority: t['Priority'] || '',
    requesterName: t['Requester Name'] || '',
    requesterEmail: t['Requester Email'] || '',
    date: t['Date Submitted'] || '',
  }));
}

function patchTicketIssues(htmlContent, ticketIssues) {
  const marker = 'const TICKET_ISSUES = ';
  const startIdx = htmlContent.indexOf(marker);
  if (startIdx < 0) return null; // dashboard hasn't been updated to expect this yet — skip, don't fail the sync
  const jsonStart = startIdx + marker.length;
  const endMarker = '];\n';
  const endIdx = htmlContent.indexOf(endMarker, jsonStart);
  if (endIdx < 0) return null;
  const before = htmlContent.slice(0, jsonStart);
  const after = htmlContent.slice(endIdx + 1);
  const newJson = JSON.stringify(ticketIssues);
  return before + newJson + after;
}

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Missing AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET environment variables.');
  }
  console.log('Authenticating with Microsoft Graph (app-only)...');
  const token = await getAppToken();
  console.log('Got token. Downloading workbook from', FILE_PATH, 'for', USER_UPN, '...');
  const buf = await downloadWorkbook(token);
  console.log(`Downloaded ${(buf.length / 1024 / 1024).toFixed(1)} MB. Extracting...`);
  const { DS, issues, missing, sheetNames } = extractAll(buf);
  if (missing.length) console.warn('WARNING: sheets not found in workbook:', missing.join(', '));
  console.log('Extracted sheets:', Object.keys(DS).map(k => `${k}(${DS[k].rows.length})`).join(', '));
  console.log('Extracted HQ issues:', issues.length);

  console.log('Reading', INDEX_HTML_PATH, '...');
  const html = fs.readFileSync(INDEX_HTML_PATH, 'utf-8');
  let patched = patchIndexHtml(html, DS);
  patched = patchSourceLabel(patched);
  patched = patchIssues(patched, issues);

  try {
    const ticketIssues = readTicketIssuesFromBoard(TICKET_BOARD_HTML_PATH);
    if (ticketIssues == null) {
      console.log('No ticket-board.html / EMBEDDED_TICKETS found yet — skipping ticket mirror (non-fatal).');
    } else {
      const withTickets = patchTicketIssues(patched, ticketIssues);
      if (withTickets == null) {
        console.log('index.html has no TICKET_ISSUES marker — skipping ticket mirror (non-fatal).');
      } else {
        patched = withTickets;
        console.log('Mirrored', ticketIssues.length, 'tickets into TICKET_ISSUES (HQ Issues & requests tab).');
      }
    }
  } catch (err) {
    // Never let a ticket-mirroring problem fail the main dashboard sync.
    console.log('Ticket mirror step failed (non-fatal):', err.message);
  }

  fs.writeFileSync(INDEX_HTML_PATH, patched, 'utf-8');
  console.log('index.html updated with fresh data, issues, tickets, and current label.');
}

module.exports = { patchIndexHtml, patchSourceLabel, patchIssues, readTicketIssuesFromBoard, patchTicketIssues, getAppToken, downloadWorkbook };

if (require.main === module) {
  main().catch(err => {
    console.error('Sync failed:', err.message);
    process.exit(1);
  });
}
