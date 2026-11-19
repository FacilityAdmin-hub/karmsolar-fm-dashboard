// Turns a requester's email reply into a comment on their ticket.
//
// The ticket emails (confirmation, status change) go out from the FM mailbox
// (GRAPH_USER_UPN) with the ticket number in the subject, e.g. "Ticket submitted: TCK-261105-091500-ABC".
// When the requester hits Reply, the answer lands in that same mailbox as "RE: Ticket submitted: TCK-...".
// This script runs inside the every-minute sync (before sync-tickets.js), reads the recent Inbox,
// and for each reply that
//   - is a real reply (subject starts with RE:), and
//   - mentions a ticket number that exists, and
//   - was sent by that ticket's requester (the Requester Email on the ticket)
// appends a row to the "Comments" table of FM Tickets.xlsx:
//   Ticket ID | received time (UTC) | Comment | <requester name> | <reply text without the quoted thread> | mail:<id hash>
// The "mail:<hash>" marker in the AttachmentURL column is how a message is remembered as handled,
// so nothing is ever added twice (the board ignores that column for plain comments).
// The normal sync then publishes the comment on the board and emails the FM team about it.
//
// Needs the Entra app to have the Microsoft Graph APPLICATION permission Mail.Read (admin consent),
// the same app that already has Mail.Send. Without it this script logs one line and does nothing.
// It never throws: a problem here must not stop the ticket sync.
const crypto = require('crypto');

const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const USER_UPN = process.env.GRAPH_USER_UPN || 'mohi.mohsen@karmsolar.com';
const FILE_PATH = process.env.GRAPH_TICKETS_FILE_PATH || '/FM Tickets.xlsx';
const LOOKBACK_HOURS = Number(process.env.INBOUND_LOOKBACK_HOURS || 72);
const MAX_PER_RUN = Number(process.env.INBOUND_MAX_PER_RUN || 20);
const MAX_LEN = 4000;
const GRAPH = 'https://graph.microsoft.com/v1.0';

const TICKET_RE = /TCK-\d{6}-\d{6}-[A-Z0-9]{3}/i;
const REPLY_SUBJECT_RE = /^\s*(re|aw|sv|رد)\s*:/i;

// ---- text helpers (exported for tests) -----------------------------------------------------
function htmlToText(html) {
  return String(html || '')
    .replace(/<(style|script)[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|li|h\d)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&');
}

// Keep what the person typed; drop the quoted thread, signature separators and mobile footers.
function stripQuoted(text) {
  const lines = String(text || '').replace(/\r\n?/g, '\n').split('\n');
  const cutAt = (i) => lines.slice(0, i);
  let out = lines;
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i].trim();
    if (
      /^-{2,}\s*(original message|forwarded message)\s*-{2,}$/i.test(l) ||
      /^_{8,}$/.test(l) ||
      /^on .{5,200} wrote:$/i.test(l) ||
      /^(from|de|von|من)\s*:\s*\S/i.test(l) && /^(sent|date|to|subject|cc|أُرسل|إلى|الموضوع)\s*:/i.test((lines[i + 1] || '').trim()) ||
      /^(from|de|von|من)\s*:\s*.*@.*$/i.test(l) && i > 0 && lines[i - 1].trim() === '' ||
      /^sent from my /i.test(l) ||
      /^get outlook for /i.test(l) ||
      /^>/.test(l) && i > 0 && /wrote:$|:$/.test(lines[i - 1].trim())
    ) { out = cutAt(i); break; }
  }
  // Drop trailing quote lines ("> ...") and blank lines.
  out = out.filter(l => !/^\s*>/.test(l));
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

function ticketIdFrom(subject) {
  const m = TICKET_RE.exec(String(subject || ''));
  return m ? m[0].toUpperCase() : null;
}

function cellText(s) {
  s = String(s == null ? '' : s).replace(/\u0000/g, '').slice(0, MAX_LEN);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function stamp(iso) {
  const d = new Date(iso);
  if (isNaN(d)) return new Date().toISOString().slice(0, 16).replace('T', ' ');
  return d.toISOString().slice(0, 16).replace('T', ' ');
}

function markerFor(msg) {
  const key = msg.internetMessageId || msg.id || '';
  return 'mail:' + crypto.createHash('sha1').update(String(key)).digest('hex').slice(0, 16);
}

// Pure decision function: which of these messages should become comments?
// tickets: [{id, email, name}], knownMarkers: Set of 'mail:...' already stored.
function pickReplies(messages, tickets, knownMarkers) {
  const byId = new Map(tickets.map(t => [t.id.toUpperCase(), t]));
  const picked = [], skipped = [];
  for (const m of messages) {
    const marker = markerFor(m);
    if (knownMarkers.has(marker)) continue;                       // already handled
    if (!REPLY_SUBJECT_RE.test(m.subject || '')) continue;        // original notices, auto-replies, forwards
    const id = ticketIdFrom(m.subject);
    if (!id) continue;                                            // not about a ticket
    const t = byId.get(id);
    const from = String((m.from && m.from.emailAddress && m.from.emailAddress.address) || '').trim().toLowerCase();
    if (!t) { skipped.push({ id, why: 'unknown ticket' }); continue; }
    if (!t.email || from !== t.email.toLowerCase()) { skipped.push({ id, why: 'sender is not the requester' }); continue; }
    const raw = m.uniqueBody && m.uniqueBody.content ? m.uniqueBody : m.body;
    let text = raw ? (String(raw.contentType).toLowerCase() === 'html' ? htmlToText(raw.content) : String(raw.content || '')) : (m.bodyPreview || '');
    text = stripQuoted(text);
    if (!text) { skipped.push({ id, why: 'empty after removing quoted text' }); continue; }
    picked.push({
      ticketId: t.id,
      timestamp: stamp(m.receivedDateTime),
      author: t.name || (m.from && m.from.emailAddress && m.from.emailAddress.name) || from,
      content: cellText(text),
      marker,
    });
  }
  return { picked, skipped };
}

// ---- Graph ---------------------------------------------------------------------------------
async function getAppToken() {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
  });
  if (!res.ok) throw new Error(`Token request failed: ${res.status} ${await res.text()}`);
  return (await res.json()).access_token;
}

async function readInbox(token) {
  const since = new Date(Date.now() - LOOKBACK_HOURS * 3600 * 1000).toISOString();
  const base = `${GRAPH}/users/${encodeURIComponent(USER_UPN)}/mailFolders/inbox/messages`;
  const headers = { Authorization: `Bearer ${token}`, Prefer: 'outlook.body-content-type="text"' };
  const q = (select) => `${base}?$filter=${encodeURIComponent(`receivedDateTime ge ${since}`)}&$orderby=${encodeURIComponent('receivedDateTime asc')}&$top=100&$select=${select}`;
  let url = q('id,subject,from,receivedDateTime,internetMessageId,bodyPreview,uniqueBody,body');
  let res = await fetch(url, { headers });
  if (res.status === 400) {                                     // some tenants reject uniqueBody in list calls
    url = q('id,subject,from,receivedDateTime,internetMessageId,bodyPreview,body');
    res = await fetch(url, { headers });
  }
  if (res.status === 401 || res.status === 403) {
    console.log(`Inbound replies: Graph says ${res.status} reading ${USER_UPN}'s Inbox. The Entra app needs the Application permission Mail.Read (with admin consent). Skipping.`);
    return null;
  }
  if (!res.ok) throw new Error(`Inbox read failed: ${res.status} ${await res.text()}`);
  const all = [];
  for (let page = 0; page < 5; page++) {
    const j = await res.json();
    all.push(...(j.value || []));
    if (!j['@odata.nextLink']) break;
    res = await fetch(j['@odata.nextLink'], { headers });
    if (!res.ok) throw new Error(`Inbox paging failed: ${res.status}`);
  }
  return all;
}

async function resolveFile(token) {
  const res = await fetch(`${GRAPH}/users/${encodeURIComponent(USER_UPN)}/drive/root:${encodeURI(FILE_PATH)}`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Could not resolve file: ${res.status} ${await res.text()}`);
  const item = await res.json();
  return { driveId: item.parentReference.driveId, itemId: item.id };
}

async function tableValues(token, driveId, itemId, table) {
  const res = await fetch(`${GRAPH}/drives/${driveId}/items/${itemId}/workbook/tables('${table}')/range?$select=values`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Could not read ${table}: ${res.status} ${await res.text()}`);
  return (await res.json()).values || [];
}

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) { console.log('Inbound replies: Azure credentials missing, skipping.'); return; }
  const token = await getAppToken();
  const messages = await readInbox(token);
  if (!messages) return;
  const candidates = messages.filter(m => REPLY_SUBJECT_RE.test(m.subject || '') && ticketIdFrom(m.subject));
  if (!candidates.length) { console.log(`Inbound replies: ${messages.length} recent messages, none are ticket replies.`); return; }

  const { driveId, itemId } = await resolveFile(token);
  const ticketRows = await tableValues(token, driveId, itemId, 'Tickets');
  const commentRows = await tableValues(token, driveId, itemId, 'Comments');
  const clean = v => String(v == null ? '' : v).replace(/^'/, '').trim();
  const tickets = ticketRows.slice(1).map(r => ({ id: clean(r[0]), name: clean(r[7]), email: clean(r[8]) })).filter(t => t.id);
  const known = new Set(commentRows.slice(1).map(r => String(r[5] || '')).filter(s => s.startsWith('mail:')));

  const { picked, skipped } = pickReplies(candidates, tickets, known);
  for (const s of skipped) console.log(`Inbound replies: ignored a reply for ${s.id} (${s.why}).`);
  const batch = picked.slice(0, MAX_PER_RUN);
  for (const c of batch) {
    const res = await fetch(`${GRAPH}/drives/${driveId}/items/${itemId}/workbook/tables('Comments')/rows/add`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: [[c.ticketId, c.timestamp, 'Comment', c.author, c.content, c.marker]] }),
    });
    if (!res.ok) { console.log(`Inbound replies: could not add comment for ${c.ticketId}: ${res.status} ${await res.text()}`); break; }
    console.log(`Inbound replies: added reply from ${c.author} as a comment on ${c.ticketId}.`);
  }
  if (picked.length > batch.length) console.log(`Inbound replies: ${picked.length - batch.length} more waiting, they will be added on the next run.`);
}

if (require.main === module) {
  main().catch(err => console.log('Inbound replies failed (sync continues):', err.message));
}

module.exports = { htmlToText, stripQuoted, ticketIdFrom, pickReplies, markerFor, cellText };
