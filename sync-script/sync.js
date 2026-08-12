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

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Missing AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET environment variables.');
  }
  console.log('Authenticating with Microsoft Graph (app-only)...');
  const token = await getAppToken();
  console.log('Got token. Downloading workbook from', FILE_PATH, 'for', USER_UPN, '...');
  const buf = await downloadWorkbook(token);
  console.log(`Downloaded ${(buf.length / 1024 / 1024).toFixed(1)} MB. Extracting...`);
  const { DS, missing, sheetNames } = extractAll(buf);
  if (missing.length) console.warn('WARNING: sheets not found in workbook:', missing.join(', '));
  console.log('Extracted sheets:', Object.keys(DS).map(k => `${k}(${DS[k].rows.length})`).join(', '));

  console.log('Reading', INDEX_HTML_PATH, '...');
  const html = fs.readFileSync(INDEX_HTML_PATH, 'utf-8');
  const patched = patchIndexHtml(html, DS);
  fs.writeFileSync(INDEX_HTML_PATH, patched, 'utf-8');
  console.log('index.html updated with fresh data.');
}

module.exports = { patchIndexHtml, getAppToken, downloadWorkbook };

if (require.main === module) {
  main().catch(err => {
    console.error('Sync failed:', err.message);
    process.exit(1);
  });
}
