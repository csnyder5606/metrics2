// Ledger sync endpoint for csnyder5606.github.io/metrics2
// Paste into the ledger Google Sheet: Extensions > Apps Script. Deploy as a Web app
// (Execute as: Me, Who has access: Anyone). The app's sync link is <Web app URL>?key=<KEY>.
// KEY is the only thing protecting the data: keep it out of the public repo.
const KEY = 'PASTE_PRIVATE_KEY_HERE';
const SHEET = 'Entries';

function doGet(e) { return handle_(e, null); }
function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return out_({ ok: false, error: 'Bad JSON' }); }
  return handle_(e, body);
}

function handle_(e, body) {
  if (!e || !e.parameter || e.parameter.key !== KEY) return out_({ ok: false, error: 'Bad key' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    if (body && body.rows && typeof body.rows === 'object') applyRows_(sh, body.rows);
    return out_({ ok: true, entries: readAll_(sh) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) { sh.getRange(1, 1).setValue('date'); sh.setFrozenRows(1); }
  sh.getRange('A:A').setNumberFormat('@'); // keep dates as plain yyyy-mm-dd text
  return sh;
}

function dateStr_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v).trim();
}

function readAll_(sh) {
  const out = {};
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return out;
  const vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const head = vals[0].map(String);
  for (let r = 1; r < vals.length; r++) {
    const d = dateStr_(vals[r][0]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) continue;
    const rec = {};
    for (let c = 1; c < head.length; c++) {
      const v = vals[r][c];
      if (!head[c] || v === '' || v === null) continue;
      rec[head[c]] = v;
    }
    if (Object.keys(rec).length) out[d] = rec;
  }
  return out;
}

function applyRows_(sh, rows) {
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const newKeys = [];
  Object.keys(rows).forEach(d => {
    const rec = rows[d];
    if (rec) Object.keys(rec).forEach(k => { if (head.indexOf(k) < 0 && newKeys.indexOf(k) < 0) newKeys.push(k); });
  });
  if (newKeys.length) {
    sh.getRange(1, head.length + 1, 1, newKeys.length).setValues([newKeys]);
    head = head.concat(newKeys);
  }
  const lastRow = sh.getLastRow();
  const dates = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => dateStr_(r[0])) : [];
  const toDelete = [];
  Object.keys(rows).forEach(d => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    const rec = rows[d];
    const idx = dates.indexOf(d);
    if (!rec) { if (idx >= 0) toDelete.push(idx + 2); return; }
    const line = head.map((h, i) => (i === 0 ? d : cell_(rec[h])));
    const row = idx >= 0 ? idx + 2 : dates.length + 2;
    sh.getRange(row, 1, 1, head.length).setValues([line]);
    if (idx < 0) dates.push(d);
  });
  toDelete.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).sort(1);
}

// Strings get a leading apostrophe so Sheets stores them as typed (no auto dates/formulas).
function cell_(v) {
  if (v === undefined || v === null) return '';
  return typeof v === 'string' ? "'" + v : v;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
