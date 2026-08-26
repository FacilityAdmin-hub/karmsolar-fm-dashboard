// Triggered when a new GitHub Issue is opened by the public ticket-request page.
// Authenticates as the app itself (app-only, same credentials as the read-side
// sync-tickets.js), parses the ticket fields out of the issue body, and appends
// a row to FM Tickets.xlsx. This runs as a trusted service identity, not as the
// individual employee, so it isn't subject to the tenant's cross-user sharing
// restriction that blocks per-employee delegated writes.
const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const USER_UPN = process.env.GRAPH_USER_UPN || 'mohi.mohsen@karmsolar.com';
const FILE_PATH = process.env.GRAPH_TICKETS_FILE_PATH || '/FM Tickets.xlsx';
const TABLE_NAME = 'Tickets';

const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const REPO = process.env.GITHUB_REPOSITORY; // "owner/repo", provided by Actions
const ISSUE_NUMBER = process.env.ISSUE_NUMBER;
const ISSUE_BODY = process.env.ISSUE_BODY;

async function getAppToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  if (!res.ok) throw new Error(`Token request failed: ${res.status} ${await res.text()}`);
  return (await res.json()).access_token;
}

function genTicketId() {
  const now = new Date();
  const p = n => String(n).padStart(2, '0');
  const stamp = `${String(now.getFullYear()).slice(2)}${p(now.getMonth() + 1)}${p(now.getDate())}-${p(now.getHours())}${p(now.getMinutes())}${p(now.getSeconds())}`;
  const rand = Math.random().toString(36).slice(2, 5).toUpperCase();
  return `TCK-${stamp}-${rand}`;
}

async function addRow(token, fields) {
  const encodedPath = encodeURI(FILE_PATH);
  // Resolve the file via the user's drive path (app-only, same pattern as sync-tickets.js)
  const itemUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(USER_UPN)}/drive/root:${encodedPath}`;
  const itemRes = await fetch(itemUrl, { headers: { Authorization: `Bearer ${token}` } });
  if (!itemRes.ok) throw new Error(`Could not resolve file: ${itemRes.status} ${await itemRes.text()}`);
  const item = await itemRes.json();
  const driveId = item.parentReference.driveId;
  const itemId = item.id;

  const now = new Date();
  const dateStr = now.toISOString().slice(0, 16).replace('T', ' ');
  const ticketId = genTicketId();

  const rowValues = [[
    ticketId,
    dateStr,
    fields.title || '',
    fields.description || '',
    fields.category || '',
    fields.priority || '',
    fields.site || '',
    fields.requesterName || '',
    fields.requesterEmail || '',
    'Open',
    '', '', ''
  ]];

  const addUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${TABLE_NAME}')/rows/add`;
  const addRes = await fetch(addUrl, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: rowValues })
  });
  if (!addRes.ok) throw new Error(`Row add failed: ${addRes.status} ${await addRes.text()}`);
  return ticketId;
}

async function closeIssue(ticketId, ok, errMsg) {
  if (!GITHUB_TOKEN || !REPO || !ISSUE_NUMBER) return;
  const commentUrl = `https://api.github.com/repos/${REPO}/issues/${ISSUE_NUMBER}/comments`;
  const commentBody = ok
    ? `Recorded as ${ticketId}.`
    : `Failed to record this ticket: ${errMsg}`;
  await fetch(commentUrl, {
    method: 'POST',
    headers: { Authorization: `Bearer ${GITHUB_TOKEN}`, 'Content-Type': 'application/json', Accept: 'application/vnd.github+json' },
    body: JSON.stringify({ body: commentBody })
  }).catch(() => {});

  const closeUrl = `https://api.github.com/repos/${REPO}/issues/${ISSUE_NUMBER}`;
  await fetch(closeUrl, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${GITHUB_TOKEN}`, 'Content-Type': 'application/json', Accept: 'application/vnd.github+json' },
    body: JSON.stringify({ state: 'closed', labels: [ok ? 'ticket-recorded' : 'ticket-failed'] })
  }).catch(() => {});
}

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Missing AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET environment variables.');
  }
  if (!ISSUE_BODY) throw new Error('Missing ISSUE_BODY environment variable.');

  let fields;
  try {
    fields = JSON.parse(ISSUE_BODY);
  } catch (e) {
    throw new Error('Issue body was not valid JSON: ' + e.message);
  }

  console.log('Authenticating with Microsoft Graph (app-only)...');
  const token = await getAppToken();
  console.log('Writing ticket row...');
  let ticketId;
  try {
    ticketId = await addRow(token, fields);
    console.log('Row added:', ticketId);
    await closeIssue(ticketId, true, null);
  } catch (err) {
    console.error('Failed to add row:', err.message);
    await closeIssue(null, false, err.message);
    throw err;
  }
}

main().catch(err => {
  console.error('write-ticket failed:', err.message);
  process.exit(1);
});
