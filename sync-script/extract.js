// Ported faithfully from the dashboard's client-side extraction logic (dash.html)
// so the backend produces byte-for-byte-equivalent DS objects to what the browser computes.
const XLSX = require('xlsx');

const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function num(x) {
  if (x == null) return 0;
  if (typeof x === 'number') return isFinite(x) ? x : 0;
  if (x instanceof Date) return 0;
  let s = String(x).replace(/EGP/gi, '').replace(/[,\s]/g, '').replace(/[^\d.\-]/g, '');
  let n = parseFloat(s);
  return isFinite(n) ? n : 0;
}
function isNumish(x) {
  if (x === '' || x == null) return false;
  if (x instanceof Date) return false;
  return isFinite(parseFloat(String(x).replace(/[,\s]/g, '')));
}
function toDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
  if (typeof v === 'number') { if (v < 20000 || v > 90000) return null; return new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000); }
  const s = String(v).trim();
  if (/^\d+(\.\d+)?$/.test(s)) { const n = +s; if (n > 20000 && n < 90000) return new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 86400000); return null; }
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) { let d = +m[1], mo = +m[2], y = +m[3]; if (y < 100) y += 2000; if (mo > 12 && d <= 12) { const t = d; d = mo; mo = t; } const dt = new Date(Date.UTC(y, mo - 1, d)); return isNaN(dt) ? null : dt; }
  m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  const p = Date.parse(s); return isNaN(p) ? null : new Date(p);
}

// ---------- SPEC: identical to dash.html's declarative sheet map ----------
const SPEC = {
  payments: { name: 'Payments', cols: [0,1,2,3,4,7,8,10,12,13], valid: c => isNumish(c[0]) },
  hk: { name: 'H.K Feedback', cols: [0,5,6,7,8,9,10,11,12,13,14,18,19], valid: c => isNumish(c[0]) && num(c[0]) > 0 },
  hkext: { name: 'Housekeeping performance Survey', cols: [0,1,5,6,7,8,9,10,11,12,13,14,15], valid: c => { const id = String(c[0] || '').trim(); return id !== '' && id !== '0'; } },
  kstd: { name: 'Karm standard monthly consumpt.', cols: [0,1,2,3,4,5,6,7,8], valid: c => String(c[0] || '').trim() !== '' },
  water: { name: 'WATER&ELECTRICAL CONSUMPTION', cols: [1,5,6,7,8,9,10,11,12,13,14], valid: c => num(c[0]) > 0 && toDate(c[1]) != null },
  adminconsumption: { name: 'ADMIN consumption STOCK', cols: Array.from({length:44},(_,i)=>i+1), valid: c => num(c[0]) > 0 && toDate(c[1]) != null },
  stock: { name: 'Admin current stock', cols: [0,2,3,4,6,8], valid: c => String(c[0] || '').trim() !== '' },
  contracts: { name: 'contracts', cols: [0,2,3,4,5,6,7,8,10,11,16,17], valid: c => String(c[0] || '').trim() !== '' },
  couriers: { name: 'courriers', cols: [0,1,2,5], valid: c => toDate(c[0]) != null },
  fna: { name: 'FNA Budget', cols: [0,1,2,3,4,5,6,7,8,9,10,11,12,13,29], valid: c => String(c[0] || '').trim() !== '' },
  sand: { name: 'Sand Admin Budget', cols: [0,1,2,3,4,5,6,7,8,9,10,11,12,13,28], valid: c => String(c[0] || '').trim() !== '' },
  capex: { name: 'CAPEX', cols: [0,1,2,3,5,6,7,8], valid: c => String(c[0] || '').trim() !== '' },
  actual: { name: 'Actual expenses Budget 2.0', cols: [0,1,2,3,4,5,6,7,8,9,10,11,12,13], valid: c => String(c[0] || '').trim() !== '' },
  pr: { name: 'PR ', cols: [0,1,2,3], valid: c => String(c[0] || '').trim() !== '' },
  assetsdata: { name: 'Assets Data', cols: [0,1,2,3,4], valid: c => String(c[0] || '').trim() !== '' },
  ppmlog: { name: 'PPM Log', headerRow: 2, cols: [0,1,2,3,4,5,6,7,8,9,10,11,12], valid: c => String(c[1] || '').trim() !== '' && !/example row/i.test(String(c[11] || '')) },
  pest: { name: 'Pest Control Tracker', cols: [0,1,2], valid: c => String(c[0] || '').trim() !== '' }
};

function extractSheet(wb, spec) {
  const target = spec.name.replace(/\s+/g, '').toLowerCase();
  const sname = wb.SheetNames.find(n => n.replace(/\s+/g, '').toLowerCase() === target)
    || wb.SheetNames.find(n => n.replace(/\s+/g, '').toLowerCase().startsWith(target.slice(0, 8)));
  if (!sname) return null;
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sname], { header: 1, raw: true, blankrows: false, defval: '' });
  const hr = spec.headerRow || 0;
  if (!aoa.length || aoa.length <= hr) return { headers: spec.cols.map(() => ''), rows: [] };
  const headers = spec.cols.map(c => String((aoa[hr] || [])[c] == null ? '' : (aoa[hr] || [])[c]).trim());
  const rows = []; let empt = 0;
  for (let i = hr + 1; i < aoa.length; i++) {
    const raw = aoa[i] || [];
    if (spec.valid(raw)) { rows.push(spec.cols.map(c => { const v = raw[c]; return v == null ? '' : v; })); empt = 0; }
    else { empt++; if (empt > 250 && rows.length > 0) break; }
  }
  return { headers, rows };
}

// ---------- HQ Issues: ported faithfully from dash.html ----------
function sheetRows(wb, name) { return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '' }); }
function findIssuesSheetRows(wb) {
  const z = n => String(n).replace(/\s+/g, '').toLowerCase();
  const sname = wb.SheetNames.find(n => z(n) === 'hqissues')
    || wb.SheetNames.find(n => z(n).includes('hqissue'))
    || wb.SheetNames.find(n => /issue|request|ticket|complaint|maintenance/i.test(n));
  if (!sname) return null;
  return { name: sname, rows: sheetRows(wb, sname) };
}
function _gcol(low, keys, anti) {
  for (const k of keys) { for (let i = 0; i < low.length; i++) { if (low[i].indexOf(k) >= 0 && !(anti || []).some(a => low[i].indexOf(a) >= 0)) return i; } }
  return -1;
}
function normStatus(s) { s = String(s || '').toLowerCase(); if (s.includes('done') || s.includes('closed') || s.includes('complete') || s.includes('finish')) return 'Done'; if (s.includes('procure')) return 'Awaiting procurement'; if (s.includes('approv')) return 'Awaiting approval'; if (s.includes('progress') || s.includes('wip') || s.includes('ongoing')) return 'In progress'; return 'Awaiting approval'; }
function normType(s) { s = String(s || '').toLowerCase(); if (s.includes('cap')) return 'CAPEX'; if (s.includes('op')) return 'OPEX'; return 'General'; }
function newId() { return 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5); }
function issueRef(issues, seq) { const y = new Date().getFullYear(); return 'REQ-' + y + '-' + String(seq).padStart(4, '0'); }

function _issueUpsert(rec, byKey, c, ISSUES, refSeq) {
  let refRaw = String(rec.ref || '').trim(); if (refRaw === '0') refRaw = '';
  let title = String(rec.title || '').trim(); if (title === '0') title = '';
  const loc = (String(rec.loc || '').trim()) || 'HQ';
  if (!title && !refRaw) return;
  const key = 'sheet:' + (refRaw ? refRaw.toLowerCase() : (title.toLowerCase() + '|' + loc.toLowerCase()));
  if (c.seen[key]) return; c.seen[key] = 1;
  const status = rec.status ? normStatus(rec.status) : 'Awaiting approval';
  const type = rec.type ? normType(rec.type) : 'General';
  const category = (String(rec.category || '').trim()) || 'General';
  const appr = String(rec.appr || '').trim();
  const created = rec.date ? (new Date(rec.date).getTime() || Date.now()) : Date.now();
  const fields = rec.fields || {};
  refSeq.n++;
  const it = {
    id: newId(), key, src: 'sheet', ref: refRaw || issueRef(ISSUES, refSeq.n), title: title || ('Request ' + (c.added + c.updated + 1)),
    type, category, loc, site: '', status, created,
    actions: [{ type: 'open', value: 'imported', ts: created }, { type: 'status', value: status, ts: created }], fields
  };
  if (appr) it.actions.push({ type: 'approval', value: appr, ts: created });
  ISSUES.push(it); c.added++;
}
function parseIssuesMapped(rows, map, ISSUES, refSeq) {
  if (!rows || !rows.length) return { added: 0 };
  const byKey = {};
  const c = { added: 0, updated: 0, seen: {} };
  const g = (row, k) => (map[k] != null && map[k] >= 0 ? String(row[map[k]] == null ? '' : row[map[k]]).trim() : '');
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || []; if (row.every(x => String(x || '').trim() === '')) continue;
    const fields = {}; if (map.fields) for (const f in map.fields) { const ci = map.fields[f]; if (ci >= 0) fields[f] = String(row[ci] == null ? '' : row[ci]).trim(); }
    _issueUpsert({ ref: g(row, 'ref'), title: g(row, 'title'), type: g(row, 'type'), category: g(row, 'category'), status: g(row, 'status'), loc: g(row, 'loc'), appr: g(row, 'appr'), date: g(row, 'date'), fields }, byKey, c, ISSUES, refSeq);
  }
  return { added: c.added };
}
function parseHQIssuesFromWorkbook(wb) {
  const found = findIssuesSheetRows(wb); if (!found) return { issues: [], fields: [] };
  const low = (found.rows[0] || []).map(h => String(h || '').toLowerCase());
  const g = (keys, anti) => _gcol(low, keys, anti);
  const ISSUE_FIELDS = [];
  const map = {
    ref: g(['ref', 'request id', 'req no', 'ticket', 'id'], ['by', 'date', 'depart', 'dept']),
    title: g(['issue description', 'description', 'problem', 'title', 'subject', 'task', 'request', 'complaint', 'details', 'item'], ['category', 'priority', 'date', 'by', 'depart', 'dept', 'time', 'status', 'type', 'photo', 'email', 'name', 'no.']),
    type: g(['capex', 'opex', 'expense type', 'type'], ['date', 'time']),
    category: g(['issue category', 'category', 'class', 'group'], []),
    status: g(['progress', 'status', 'state', 'stage'], []),
    loc: g(['location / floor', 'location', 'floor', 'site', 'branch'], []),
    appr: g(['approver', 'approved by', 'approval'], []),
    date: g(['start time', 'date', 'created', 'reported', 'opened', 'logged', 'submitted', 'timestamp'], []),
    fields: {}
  };
  const addF = (name, keys, anti) => { const i = g(keys, anti); if (i >= 0) { if (ISSUE_FIELDS.indexOf(name) < 0) ISSUE_FIELDS.push(name); map.fields[name] = i; } };
  addF('Priority', ['priority'], []);
  addF('Department', ['department', 'dept'], []);
  addF('Reporter', ['full name'], []); if (map.fields['Reporter'] == null) addF('Reporter', ['name'], ['full', 'file', 'user']);
  addF('Photo', ['photo', 'image', 'attachment'], []);
  const ISSUES = []; const refSeq = { n: 0 };
  parseIssuesMapped(found.rows, map, ISSUES, refSeq);
  return { issues: ISSUES, fields: ISSUE_FIELDS };
}

const zlib = require('zlib');

// ---------- ZIP-level fallback for sheets too large for standard SheetJS parsing ----------
// (ported from dash.html's extractBloatedWater; browser used DecompressionStream, we use zlib)
function _u16(b, o) { return b[o] | (b[o + 1] << 8); }
function _u32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] * 16777216)) >>> 0; }
function _zipCD(bytes) {
  let eo = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 66000); i--) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) { eo = i; break; }
  }
  if (eo < 0) return null;
  let cnt = _u16(bytes, eo + 10), p = _u32(bytes, eo + 16); const map = {};
  for (let i = 0; i < cnt && p + 46 <= bytes.length; i++) {
    if (!(bytes[p] === 0x50 && bytes[p + 1] === 0x4b && bytes[p + 2] === 0x01 && bytes[p + 3] === 0x02)) break;
    const method = _u16(bytes, p + 10), compSize = _u32(bytes, p + 20), nameLen = _u16(bytes, p + 28), extraLen = _u16(bytes, p + 30), commLen = _u16(bytes, p + 32), lo = _u32(bytes, p + 42);
    map[bytes.toString('utf8', p + 46, p + 46 + nameLen)] = { method, compSize, lo };
    p += 46 + nameLen + extraLen + commLen;
  }
  return map;
}
function _dataStart(bytes, lo) { return lo + 30 + _u16(bytes, lo + 26) + _u16(bytes, lo + 28); }
function _inflateEntry(bytes, e) {
  const start = _dataStart(bytes, e.lo);
  const sub = bytes.subarray(start, start + e.compSize);
  if (e.method === 0) return sub; // stored, no compression
  return zlib.inflateRawSync(sub);
}
function _rowCells(xml) {
  const cells = {}; const cre = /<c[^>]*?r="([A-Z]+)\d+"[^>]*?>(?:<v>([\s\S]*?)<\/v>)?/g; let m;
  while ((m = cre.exec(xml))) { if (m[2] != null) cells[m[1]] = m[2]; }
  return cells;
}
function extractBloatedWater(buf) {
  try {
    const bytes = Buffer.isBuffer(buf) ? buf : Buffer.from(buf);
    const cd = _zipCD(bytes); if (!cd) return null;
    const wbE = cd['xl/workbook.xml'], relE = cd['xl/_rels/workbook.xml.rels']; if (!wbE || !relE) return null;
    const wbXml = _inflateEntry(bytes, wbE).toString('utf8');
    const relXml = _inflateEntry(bytes, relE).toString('utf8');
    let sm = wbXml.match(/<sheet[^>]*name="([^"]*ELECTRIC[^"]*)"[^>]*r:id="([^"]+)"/i); let rid = sm ? sm[2] : null;
    if (!rid) { sm = wbXml.match(/<sheet[^>]*r:id="([^"]+)"[^>]*name="([^"]*ELECTRIC[^"]*)"/i); rid = sm ? sm[1] : null; }
    if (!rid) return null;
    const rr = new RegExp('Id="' + rid + '"[^>]*?Target="([^"]+)"').exec(relXml) || new RegExp('Target="([^"]+)"[^>]*?Id="' + rid + '"').exec(relXml);
    if (!rr) return null;
    let target = rr[1].replace(/^\//, ''); if (!/^xl\//.test(target)) target = 'xl/' + target;
    const e = cd[target]; if (!e) return null;
    const srcLet = ['B','F','G','H','I','J','K','L','M','N','O'], meterLet = ['F','H','J','L','N','O'];

    // Decompress into a Buffer (fine, no size limit) but process it in CHUNKS when
    // decoding/regex-scanning, since converting the whole thing to a JS string at once
    // hits V8's ~512MB string-length ceiling for these array-formula-bloated sheets.
    const start = _dataStart(bytes, e.lo);
    const sub = bytes.subarray(start, start + e.compSize);
    const decompressed = e.method === 0 ? sub : zlib.inflateRawSync(sub);

    const { StringDecoder } = require('string_decoder');
    const decoder = new StringDecoder('utf8');
    const CH = 4 * 1024 * 1024; // 4MB chunks
    let sbuf = '';
    const wrows = [];
    const rowre = /<row[^>]*?>([\s\S]*?)<\/row>/g;
    for (let i = 0; i < decompressed.length; i += CH) {
      const chunk = decompressed.subarray(i, Math.min(i + CH, decompressed.length));
      sbuf += decoder.write(chunk);
      let last = 0; let m; rowre.lastIndex = 0;
      while ((m = rowre.exec(sbuf))) {
        last = rowre.lastIndex;
        const c = _rowCells(m[1]);
        const dt = parseFloat(String(c.B || '').replace(/,/g, ''));
        if (!(dt > 20000)) continue;
        const meters = meterLet.map(L => parseFloat(String(c[L] || '0').replace(/,/g, '')) || 0);
        if (!meters.some(v => v !== 0)) continue;
        wrows.push(srcLet.map(L => { const v = c[L]; if (v == null || v === '') return ''; const n = parseFloat(String(v).replace(/,/g, '')); return isFinite(n) ? n : v; }));
      }
      sbuf = sbuf.slice(last);
    }
    sbuf += decoder.end();
    if (!wrows.length) return null;
    // Headers are stable for this known sheet layout; hardcoded since the bloated-sheet
    // path can't read the header row the same way (mirrors what the live sheet has shown historically).
    const headers = ['Start time','قراءة عداد الدور الثالث','باقي الفلوس في عداد الدور الثالث','قراءة عداد الدور الرابع','باقي الفلوس في عداد الدور الرابع','قراءة عداد الدور الخامس','باقي الفلوس في عداد الدور الخامس','قراءة عداد الدور السابع','باقي الفلوس في عداد الدور السابع','قراءة عداد المياه  1 الشهريه','قراءة عداد المياه 2 الشهريه'];
    return { headers, rows: wrows };
  } catch (e) { return null; }
}

// ---------- Sites Catering: bespoke pivot extraction (not a SPEC sheet) ----------
// Ported exactly from dash.html's applyWorkbook() inline logic. The sheet has one row
// per item with per-site cost columns spread across it (not a simple table), so this
// can't go through the generic extractSheet() path.
function extractCatering(wb) {
  try {
    const csn = (wb.SheetNames || []).find(n => /sites?\s*catering/i.test(n));
    if (!csn) return null;
    const cr = XLSX.utils.sheet_to_json(wb.Sheets[csn], { header: 1, raw: false, defval: '' });
    const hdr = cr[0] || [];
    const priceCols = [4, 6, 8, 10, 12, 14, 16, 18];
    const sites = priceCols.map(ci => ({ name: String(hdr[ci - 1] || hdr[ci] || ('Site ' + ci)).replace(/\s+/g, ' ').trim().slice(0, 26), total: 0 }));
    const items = [];
    for (let i = 1; i < cr.length; i++) {
      const row = cr[i] || [];
      const it = String(row[2] || '').trim();
      if (!it) continue;
      const qty = num(row[22]);
      const tot = num(row[25]);
      if (tot <= 0 && qty <= 0) continue;
      priceCols.forEach((ci, si) => { sites[si].total += num(row[ci]); });
      items.push([it, qty, tot]);
    }
    if (!items.length) return null;
    return {
      catering: { headers: ['Item', 'Qty', 'Total cost'], rows: items },
      cateringsites: { headers: ['Site', 'Total cost'], rows: sites.filter(s => s.total > 0).map(s => [s.name, Math.round(s.total)]) }
    };
  } catch (e) { return null; }
}

// "SYSTEM ONE حضور" is a Microsoft Forms response export: each row is one check-in
// session with up to 9 repeating (employee name, time) column pairs starting at
// column index 5 (F). Flattens to one record per actual (employee, time) entry.
function extractAttendance(wb) {
  try {
    const names = wb.SheetNames || [];
    // Try progressively looser matches: exact combo first, then just the Arabic word
    // (most distinctive part), then just "system one" as a last resort.
    const sn = names.find(n => /system\s*one/i.test(n) && /\u062d\u0636\u0648\u0631/.test(n))
      || names.find(n => /\u062d\u0636\u0648\u0631/.test(n))
      || names.find(n => /system\s*one/i.test(n));
    if (!sn) return null;
    // Read twice: raw:true for an unambiguous numeric date serial (Forms exports dates as
    // M/D/YYYY text, which toDate() would otherwise misread as day-first), raw:false for
    // the employee names and check-in time text (parseCheckin needs the display string).
    const crText = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: false, defval: '' });
    const crRaw = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' });
    console.error('[debug] attendance sheet name matched:', sn);
    console.error('[debug] attendance crText.length:', crText.length);
    console.error('[debug] attendance row0 (header):', JSON.stringify(crText[0]));
    console.error('[debug] attendance row1 text:', JSON.stringify(crText[1]));
    console.error('[debug] attendance row1 raw:', JSON.stringify(crRaw[1]));
    console.error('[debug] attendance row2 text:', JSON.stringify(crText[2]));
    console.error('[debug] attendance row2 raw:', JSON.stringify(crRaw[2]));
    const rows = [];
    for (let i = 1; i < crText.length; i++) {
      const row = crText[i] || [];
      const rawRow = crRaw[i] || [];
      const start = String(row[1] || '').trim();
      if (!start || start === '0') continue;
      const rawStart = rawRow[1];
      let startDate = rawStart instanceof Date ? rawStart : (typeof rawStart === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.floor(rawStart) * 86400000) : null);
      if (!startDate || isNaN(startDate)) {
        // Fallback: parse the display string directly as M/D/YYYY (Forms' native export
        // format), rather than relying on the raw numeric serial.
        const m = start.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
        if (m) { let mo = +m[1], d = +m[2], y = +m[3]; if (y < 100) y += 2000; startDate = new Date(Date.UTC(y, mo - 1, d)); }
      }
      if (i <= 3) console.error('[debug] row', i, 'start=', JSON.stringify(start), 'rawStart=', JSON.stringify(rawStart), 'parsedDate=', startDate);
      if (!startDate || isNaN(startDate) || startDate.getUTCFullYear() < 2000) continue; // skip the 1/0/1900 placeholder rows
      for (let ci = 5; ci + 1 < row.length; ci += 2) {
        const name = String(row[ci] || '').trim();
        const time = String(row[ci + 1] || '').trim();
        if (!name || name === '0') continue;
        rows.push([startDate.toISOString().slice(0, 10), name, time]);
      }
    }
    console.error('[debug] attendance rows.length after loop:', rows.length);
    if (!rows.length) return null;
    return { headers: ['Date', 'Employee', 'Check-in'], rows };
  } catch (e) { return null; }
}

// "HouseKeeping Daily checklist" is a Microsoft Forms response export: each row is one
// submission, columns from index 5 (F) onward are one per location (header = location
// name), and each cell is a semicolon-joined list of completed checklist items.
function extractHKChecklist(wb) {
  try {
    const sn = (wb.SheetNames || []).find(n => /house\s*keeping/i.test(n) && /daily/i.test(n) && /checklist/i.test(n));
    if (!sn) return null;
    const cr = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: false, defval: '' });
    const crRaw = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' });
    const hdr = cr[0] || [];
    const locCols = [];
    for (let ci = 5; ci < hdr.length; ci++) {
      const loc = String(hdr[ci] || '').trim();
      if (loc) locCols.push({ ci, loc });
    }
    const rows = [];
    for (let i = 1; i < cr.length; i++) {
      const row = cr[i] || [];
      const rawRow = crRaw[i] || [];
      const start = String(row[1] || '').trim();
      if (!start || start === '0') continue;
      const rawStart = rawRow[1];
      const startDate = rawStart instanceof Date ? rawStart : (typeof rawStart === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.floor(rawStart) * 86400000) : null);
      if (!startDate || isNaN(startDate) || startDate.getUTCFullYear() < 2000) continue; // skip the 1/0/1900 placeholder rows
      const datePart = startDate.toISOString().slice(0, 10);
      locCols.forEach(({ ci, loc }) => {
        const cell = String(row[ci] || '').trim();
        if (!cell || cell === '0') return;
        const items = cell.split(';').map(s => s.trim()).filter(Boolean);
        rows.push([datePart, loc, items.length, cell]);
      });
    }
    if (!rows.length) return null;
    return { headers: ['Date', 'Location', 'Items completed', 'Details'], rows };
  } catch (e) { return null; }
}

function extractAll(workbookBuffer) {
  const wb = XLSX.read(workbookBuffer, { type: 'buffer', cellDates: true });
  const DS = {};
  const missing = [];
  for (const key in SPEC) {
    const res = extractSheet(wb, SPEC[key]);
    if (res) DS[key] = res; else missing.push(SPEC[key].name);
  }
  // Fallback for the water sheet when it's too large for standard SheetJS parsing
  // (mirrors dash.html's extractBloatedWater, ported to Node's zlib instead of DecompressionStream)
  if ((!DS.water || !DS.water.rows.length)) {
    const bloated = extractBloatedWater(workbookBuffer);
    if (bloated && bloated.rows.length) DS.water = bloated;
  }
  const catering = extractCatering(wb);
  if (catering) { DS.catering = catering.catering; DS.cateringsites = catering.cateringsites; }
  console.error('[debug] sheet names:', JSON.stringify(wb.SheetNames));
  const attendance = extractAttendance(wb);
  console.error('[debug] attendance result:', attendance ? attendance.rows.length + ' rows' : 'null');
  if (attendance) DS.attendance = attendance;
  const hkchecklist = extractHKChecklist(wb);
  if (hkchecklist) DS.hkchecklist = hkchecklist;
  const { issues, fields } = parseHQIssuesFromWorkbook(wb);
  return { DS, issues, issueFields: fields, missing, sheetNames: wb.SheetNames };
}

module.exports = { extractAll };
