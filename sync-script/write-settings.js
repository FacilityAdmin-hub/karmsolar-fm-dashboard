// Shared dashboard settings: applies the changes people make on the dashboard (added columns and their values,
// renamed titles, card layout, saved charts, request updates) to dashboard-settings.json on the data branch.
// The dashboard sends them through the ticket proxy (Cloudflare Worker), which opens a "Settings:" issue.
// Every run applies ALL open "Settings:" issues opened by the repository owner (the account the worker posts as),
// oldest first, so a run that GitHub cancels while it waits never loses a change.
//
// Usage: node write-settings.js apply   -> writes the file and processed.json (the issue numbers it applied)
//        node write-settings.js close   -> closes the issues listed in processed.json
const fs = require('fs');
const path = require('path');

const REPO = process.env.GITHUB_REPOSITORY;
const TOKEN = process.env.GITHUB_TOKEN;
const OWNER = (process.env.REPO_OWNER || (REPO || '').split('/')[0] || '').toLowerCase();
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, '..', 'data-branch');
const FILE = path.join(DATA_DIR, 'dashboard-settings.json');
const PROCESSED = path.join(__dirname, 'processed.json');

const SECTIONS = new Set(['customcols', 'labels', 'custom', 'pageviews', 'cardsizes', 'cardhidden', 'overview', 'ovhidden', 'ovorder', 'ifields', 'issues']);
const MAX_KEY = 200, MAX_VALUE = 60000, MAX_OPS = 200;

const gh = (url, opt = {}) => fetch(url, { ...opt, headers: { Authorization: `Bearer ${TOKEN}`, Accept: 'application/vnd.github+json', 'Content-Type': 'application/json', 'User-Agent': 'dashboard-settings', ...(opt.headers || {}) } });

/* titles are shown as page text: keep plain text plus <b> <i> <u> <br>, nothing that could run code */
function cleanLabel(v) {
  if (typeof v !== 'string') return '';
  return v.replace(/<(\/?)(b|i|u|br|strong|em)\b[^>]*>/gi, '\u0001$1$2\u0002')
    .replace(/<[^>]*>?/g, '')
    .replace(/\u0001(\/?)(b|i|u|br|strong|em)\u0002/gi, (m, s, t) => `<${s}${t.toLowerCase()}>`)
    .slice(0, 2000);
}

function load() {
  try { const d = JSON.parse(fs.readFileSync(FILE, 'utf-8')); if (d && typeof d === 'object' && d.sections) return d; } catch (e) {}
  return { v: 1, updated: null, sections: {} };
}

function applyOps(doc, ops) {
  let n = 0;
  for (const op of ops) {
    if (!op || typeof op !== 'object') continue;
    const s = String(op.s || ''), k = String(op.k == null ? '' : op.k);
    if (!SECTIONS.has(s) || !k || k.length > MAX_KEY) continue;
    let v = op.v;
    if (v !== null && v !== undefined) { const j = JSON.stringify(v); if (j === undefined || j.length > MAX_VALUE) continue; v = JSON.parse(j); }
    if (s === 'labels' && v != null) v = cleanLabel(v);
    const sec = doc.sections[s] = doc.sections[s] || {};
    if (v === null || v === undefined) { if (k in sec) { delete sec[k]; n++; } }
    else { sec[k] = v; n++; }
  }
  return n;
}

async function openSettingsIssues() {
  const out = [];
  for (let page = 1; page <= 5; page++) {
    const r = await gh(`https://api.github.com/repos/${REPO}/issues?state=open&per_page=100&page=${page}&sort=created&direction=asc`);
    if (!r.ok) throw new Error(`Could not list issues: ${r.status}`);
    const list = await r.json(); if (!list.length) break;
    list.forEach(i => { if (!i.pull_request && /^Settings:/.test(i.title || '') && String(i.user && i.user.login || '').toLowerCase() === OWNER) out.push(i); });
    if (list.length < 100) break;
  }
  return out.sort((a, b) => a.number - b.number);
}

async function applyAll() {
  const issues = await openSettingsIssues();
  const doc = load(); const done = []; let changes = 0;
  for (const i of issues) {
    let body = null; try { body = JSON.parse(i.body || ''); } catch (e) {}
    const ops = body && Array.isArray(body.ops) ? body.ops.slice(0, MAX_OPS) : [];
    changes += applyOps(doc, ops); done.push({ number: i.number, ok: ops.length > 0 });
  }
  if (done.length) { doc.updated = new Date().toISOString(); fs.writeFileSync(FILE, JSON.stringify(doc), 'utf-8'); }
  fs.writeFileSync(PROCESSED, JSON.stringify(done), 'utf-8');
  console.log(`Applied ${done.length} settings issue(s), ${changes} change(s).`);
}

async function closeAll() {
  let done = []; try { done = JSON.parse(fs.readFileSync(PROCESSED, 'utf-8')); } catch (e) {}
  for (const d of done) {
    await gh(`https://api.github.com/repos/${REPO}/issues/${d.number}`, { method: 'PATCH', body: JSON.stringify({ state: 'closed', labels: [d.ok ? 'settings-saved' : 'settings-ignored'] }) });
  }
  console.log(`Closed ${done.length} settings issue(s).`);
}

module.exports = { applyOps, cleanLabel, load };

if (require.main === module) {
  const cmd = process.argv[2];
  (cmd === 'close' ? closeAll() : applyAll()).catch(err => { console.error('Settings step failed:', err.message); process.exit(1); });
}
