// Triggered when an "UpdateStatus: <ticketId>" GitHub Issue is opened by the
// public ticket-board page. Authenticates app-only (same credentials as
// write-ticket.js / delete-ticket.js / sync-tickets.js), finds the matching
// row in FM Tickets.xlsx by Ticket ID, and updates just its Status cell
// (plus Date Updated) without touching any other column.
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

// Must match STATUSES in ticket-board.html.
const VALID_STATUSES = ['Backlog', 'Open', 'In Progress', 'In Review', 'Waiting Approval', 'Blocked', 'Resolved', 'Closed'];

// Column order in the Tickets table (13 columns, A..M). Status is column J (index 9),
// Date Updated is column L (index 11).
const NUM_COLS = 13;
const STATUS_COL = 9;
const DATE_UPDATED_COL = 11;

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

async function resolveFile(token) {
  const encodedPath = encodeURI(FILE_PATH);
  const itemUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(USER_UPN)}/drive/root:${encodedPath}`;
  const itemRes = await fetch(itemUrl, { headers: { Authorization: `Bearer ${token}` } });
  if (!itemRes.ok) throw new Error(`Could not resolve file: ${itemRes.status} ${await itemRes.text()}`);
  const item = await itemRes.json();
  return { driveId: item.parentReference.driveId, itemId: item.id };
}

async function findRowIndex(token, driveId, itemId, ticketId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${TABLE_NAME}')/rows?$select=index,values`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Could not list rows: ${res.status} ${await res.text()}`);
  const json = await res.json();
  const rows = json.value || [];
  for (const row of rows) {
    const firstCell = row.values && row.values[0] && row.values[0][0];
    if (String(firstCell || '').trim() === ticketId) return row.index;
  }
  return null;
}

async function updateStatusCell(token, driveId, itemId, index, status) {
  const now = new Date();
  const dateStr = now.toISOString().slice(0, 16).replace('T', ' ');
  const rowValues = new Array(NUM_COLS).fill(null);
  rowValues[STATUS_COL] = status;
  rowValues[DATE_UPDATED_COL] = dateStr;

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/tables('${TABLE_NAME}')/rows/itemAt(index=${index})`;
  const res = await fetch(url, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [rowValues] }),
  });
  if (!res.ok) throw new Error(`Row update failed: ${res.status} ${await res.text()}`);
}

async function closeIssue(ticketId, ok, errMsg) {
  if (!GITHUB_TOKEN || !REPO || !ISSUE_NUMBER) return;
  const commentUrl = `https://api.github.com/repos/${REPO}/issues/${ISSUE_NUMBER}/comments`;
  const commentBody = ok
    ? `Status updated for ${ticketId}.`
    : `Failed to update status for ${ticketId}: ${errMsg}`;
  await fetch(commentUrl, {
    method: 'POST',
    headers: { Authorization: `Bearer ${GITHUB_TOKEN}`, 'Content-Type': 'application/json', Accept: 'application/vnd.github+json' },
    body: JSON.stringify({ body: commentBody })
  }).catch(() => {});

  const closeUrl = `https://api.github.com/repos/${REPO}/issues/${ISSUE_NUMBER}`;
  await fetch(closeUrl, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${GITHUB_TOKEN}`, 'Content-Type': 'application/json', Accept: 'application/vnd.github+json' },
    body: JSON.stringify({ state: 'closed', labels: [ok ? 'ticket-status-updated' : 'ticket-status-update-failed'] })
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
  const ticketId = (fields.ticketId || '').trim();
  const status = (fields.status || '').trim();
  if (!ticketId) throw new Error('No ticketId in issue body.');
  if (!VALID_STATUSES.includes(status)) throw new Error(`Invalid status: ${status}`);

  console.log('Authenticating with Microsoft Graph (app-only)...');
  const token = await getAppToken();
  const { driveId, itemId } = await resolveFile(token);

  try {
    console.log('Looking up row for', ticketId, '...');
    const index = await findRowIndex(token, driveId, itemId, ticketId);
    if (index == null) throw new Error(`Ticket ${ticketId} not found in table.`);
    await updateStatusCell(token, driveId, itemId, index, status);
    console.log('Updated status for', ticketId, 'to', status);
    await closeIssue(ticketId, true, null);
  } catch (err) {
    console.error('Failed to update status:', err.message);
    await closeIssue(ticketId, false, err.message);
    throw err;
  }
}

main().catch(err => {
  console.error('update-status failed:', err.message);
  process.exit(1);
});
