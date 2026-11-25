// Ledger sync endpoint for csnyder5606.github.io/metrics2
// Paste into the ledger Google Sheet: Extensions > Apps Script. Deploy as a Web app
// (Execute as: Me, Who has access: Anyone). The app's sync link is <Web app URL>?key=<SYNC_KEY>.
//
// Secrets live in Project Settings > Script Properties, never in this file (the repo is public):
//   SYNC_KEY           the private key from your sync link (everything after ?key=)
//   ANTHROPIC_API_KEY  Claude API key, used only by "Summarize my notes"
//
// Tabs: Entries (one row per day), PHQ4 (one row per check-in, never edited),
//       Insights (one row per notes summary).
const PROPS = PropertiesService.getScriptProperties();
const SHEET = 'Entries';
const PHQ_SHEET = 'PHQ4';
const PHQ_HEAD = ['id', 'timestamp', 'date', 'q1', 'q2', 'q3', 'q4', 'anxiety', 'depression', 'total'];
const INS_SHEET = 'Insights';
const INS_HEAD = ['created', 'from', 'to', 'days', 'model', 'text', 'stop_reason', 'input_tokens', 'output_tokens'];
const MODEL = 'claude-sonnet-5-5';
const SLOTS = ['Morning', 'Midday', 'Evening'];
const FLAGS = [['irritability', 'irritability'], ['cravingConnection', 'craving connection'], ['feltConnected', 'felt connected'],
  ['depression', 'depression'], ['anxiety', 'anxiety'], ['cortisol', 'cortisol feelings'], ['moodEating', 'mood-related eating']];

function doGet(e) { return handle_(e, null); }
function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return out_({ ok: false, error: 'Bad JSON' }); }
  return handle_(e, body);
}

function handle_(e, body) {
  const key = PROPS.getProperty('SYNC_KEY');
  if (!key || !e || !e.parameter || e.parameter.key !== key) return out_({ ok: false, error: 'Bad key' });
  if (body && body.action === 'insights') return out_(insights_(body));
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    if (body && body.rows && typeof body.rows === 'object') applyRows_(sh, body.rows);
    if (body && Array.isArray(body.phq4)) addPhq4_(body.phq4);
    return out_({ ok: true, entries: readAll_(sh), phq4: readPhq4_(), insight: latestInsight_() });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Entries ----------
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

// ---------- PHQ-4 (append-only) ----------
function tab_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    sh.setFrozenRows(1);
    sh.getRange('A:C').setNumberFormat('@');
  } else if (sh.getLastColumn() < head.length) {
    sh.getRange(1, 1, 1, head.length).setValues([head]); // header grew in a script update
  }
  return sh;
}

function addPhq4_(takes) {
  const sh = tab_(PHQ_SHEET, PHQ_HEAD);
  const last = sh.getLastRow();
  const have = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
  takes.forEach(t => {
    if (!t || typeof t.id !== 'string' || have.indexOf(t.id) >= 0) return;
    const a = (t.answers || []).map(Number);
    if (a.length !== 4 || a.some(x => !(x >= 0 && x <= 3 && Math.floor(x) === x))) return;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(t.date))) return;
    sh.appendRow([t.id, String(t.ts || ''), t.date, a[0], a[1], a[2], a[3], a[0] + a[1], a[2] + a[3], a[0] + a[1] + a[2] + a[3]]);
    have.push(t.id);
  });
}

function readPhq4_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(PHQ_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, PHQ_HEAD.length).getValues().map(r => ({
    id: String(r[0]), ts: String(r[1]), date: dateStr_(r[2]),
    answers: [r[3], r[4], r[5], r[6]].map(Number),
  })).filter(t => t.id);
}

// ---------- Notes insights (Claude API) ----------
function latestInsight_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(INS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const r = sh.getRange(sh.getLastRow(), 1, 1, INS_HEAD.length).getValues()[0];
  return { created: String(r[0]), from: dateStr_(r[1]), to: dateStr_(r[2]), days: Number(r[3]), text: String(r[5]) };
}

function insights_(body) {
  const from = String(body.from || ''), to = String(body.to || ''), days = Number(body.days) || 0;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) return { ok: false, error: 'Bad range' };
  const apiKey = PROPS.getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'ANTHROPIC_API_KEY is missing from Script Properties' };

  const entries = readAll_(sheet_());
  const dates = Object.keys(entries).filter(d => d >= from && d <= to).sort();
  const lines = dates.map(d => dayLine_(d, entries[d])).filter(Boolean);
  if (lines.filter(l => l.indexOf('note:') >= 0 || l.indexOf('glimmers:') >= 0).length < 2) {
    return { ok: false, error: 'Not enough notes in this range yet (needs at least 2 days with notes).' };
  }

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: MODEL, max_tokens: 8000, system: SYSTEM_PROMPT_,
      messages: [{ role: 'user', content: `Journal entries from ${from} to ${to} (${dates.length} days logged):\n\n${lines.join('\n\n')}` + linksText_(body.links) }],
    }),
  });
  const code = res.getResponseCode();
  let data;
  try { data = JSON.parse(res.getContentText()); } catch (err) { data = null; }
  if (code !== 200 || !data || !Array.isArray(data.content)) {
    const msg = data && data.error && data.error.message ? data.error.message : 'HTTP ' + code;
    return { ok: false, error: 'Claude API error: ' + msg };
  }
  const text = data.content.filter(b => b.type === 'text').map(b => b.text).join('\n').trim();
  const usage = data.usage || {};
  if (data.stop_reason === 'max_tokens' || !text) {
    return { ok: false, error: 'The summary came back incomplete (' + (data.stop_reason || 'empty') + '). Try again or a shorter range.' };
  }
  const created = new Date().toISOString();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    tab_(INS_SHEET, INS_HEAD).appendRow([created, from, to, days, MODEL, "'" + text, String(data.stop_reason || ''), usage.input_tokens || '', usage.output_tokens || '']);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, insight: { created, from, to, days, text } };
}

// Mood links computed in the app: average daily mood on days with vs without something, plus which dates.
function linksText_(links) {
  if (!Array.isArray(links) || !links.length) return '\n\nMood links: none with enough days yet.';
  const ok = links.slice(0, 12).filter(l => l && typeof l.label === 'string' && Array.isArray(l.withDates) && Array.isArray(l.withoutDates));
  if (!ok.length) return '\n\nMood links: none with enough days yet.';
  const fmt = (ds) => ds.filter(d => /^\d{4}-\d{2}-\d{2}$/.test(String(d))).join(', ');
  return '\n\nMood links (average daily mood 1-10, with vs without):\n' + ok.map(l =>
    `- ${String(l.label).slice(0, 60)}: ${Number(l.withAvg)} with (${fmt(l.withDates)}) vs ${Number(l.withoutAvg)} without (${fmt(l.withoutDates)})`).join('\n');
}

function dayLine_(d, e) {
  const n = (k) => (e[k] === '' || e[k] === undefined || e[k] === null) ? null : e[k];
  const parts = [d];
  const moods = SLOTS.map(s => n('mood' + s) !== null ? `${s} ${n('mood' + s)}` : null).filter(Boolean);
  if (moods.length) parts.push('mood (1-10): ' + moods.join(', '));
  if (n('sleepHours') !== null) parts.push('sleep last night: ' + n('sleepHours') + ' hrs');
  if (Number(n('drankSmoked')) === 1) parts.push('drank or smoked' + (Number(n('toCope')) === 1 ? ' (to cope)' : ''));
  if (Number(n('urgeResisted')) > 0) parts.push('urges resisted: ' + n('urgeResisted'));
  if (Number(n('kidsDay')) === 1) parts.push('kids with me');
  const flags = FLAGS.map(([k, label]) => {
    const on = SLOTS.filter(s => Number(n(k + s)) === 1);
    return on.length ? `${label} (${on.join('/').toLowerCase()})` : null;
  }).filter(Boolean);
  if (flags.length) parts.push('flags: ' + flags.join(', '));
  if (n('note') !== null) parts.push('note: ' + String(n('note')).replace(/\s+/g, ' '));
  if (n('glimmers') !== null) parts.push('glimmers: ' + String(n('glimmers')).replace(/\s+/g, ' '));
  return parts.length > 1 ? parts.join(' | ') : null;
}

const SYSTEM_PROMPT_ = [
  'You read a personal daily mood journal and write a short summary of insights for the person who wrote it.',
  'They may share it with their therapist. Write in plain language, in the second person ("you").',
  '',
  'Use exactly these sections, each as a line starting with "## " followed by bullets starting with "- " (2-4 bullets, except one per link under Mood links vs. notes):',
  '## Recurring themes',
  '## Triggers',
  '## What helped',
  '## Glimmers and mood',
  '## Worth raising in therapy',
  '## Mood links vs. notes',
  '',
  'For "Mood links vs. notes": write one bullet per mood link given at the end of the entries (skip the section if there are none).',
  'Each bullet states the link in a few words, then exactly one of these labels:',
  '"Notes show another cause:" only when the notes on those dates name something else going on, and cite the dates;',
  '"Possible explanation (a guess):" when nothing in the notes says so directly;',
  '"Can\'t tell from the notes."',
  'Never state a cause without one of these labels. These comparisons cover only a few days.',
  '',
  'Rules:',
  '- Ground every point in the entries and cite dates like (Nov 21). Do not invent details.',
  '- Separate what the entries show from your interpretation; mark interpretation with "possibly" or "seems".',
  '- No diagnoses or clinical labels the person did not use themselves. No moralizing, no pep talk, no generic advice.',
  '- Keep it under 400 words. Be specific and direct.',
  '- If any entry mentions thoughts of self-harm or suicide, start with one line saying so and that the 988 Suicide & Crisis Lifeline (call or text 988) is available any time.',
].join('\n');

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
