// Word Racer: class "done today" list + progress shared between a child's phones.
// Paste this into the Google Sheet: Extensions > Apps Script, replacing everything there.
// Then: Deploy > Manage deployments > (pencil) Edit > Version: New version > Deploy.
// (Use "Manage deployments", not "New deployment", so the game's link stays the same.)
//
// Tabs it uses (created automatically):
//   Finished  - one row per child per day when they finish all races: first name and car only. Kept 30 days.
//   Progress  - one row per phone: that phone's coins, medals and tricky words. Kept until the child hasn't played for 180 days.
//   Classes   - class codes allowed to use the "done today" list. To add a class, type its code in column A.
//   Feedback  - messages parents send from Help > Send feedback, with the phone type and app version.

const FINISHED = 'Finished', PROGRESS = 'Progress', CLASSES = 'Classes', FEEDBACK = 'Feedback';
const KEEP_FINISHED_DAYS = 30, KEEP_PROGRESS_DAYS = 180;

function tab_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    if (name === CLASSES) sh.appendRow(['2021']);   // the class that already uses the list
  }
  return sh;
}
const finished_ = () => tab_(FINISHED, ['day', 'class', 'player', 'name', 'car', 'time']);
const progress_ = () => tab_(PROGRESS, ['key', 'player', 'phone', 'updated', 'data']);

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
const clean_ = (v, n) => String(v || '').replace(/[^a-z0-9]/gi, '').toLowerCase().slice(0, n || 20);

function classes_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('classes');
  if (hit) return JSON.parse(hit);
  const sh = tab_(CLASSES, ['class code']);
  const list = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues()
    .map(r => String(r[0]).toUpperCase().trim()).filter(String);
  cache.put('classes', JSON.stringify(list), 600);   // re-read the tab every 10 minutes
  return list;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'pull') return pull_(clean_(p.player));
  const cls = String(p.class || '').toUpperCase().trim();
  if (classes_().indexOf(cls) < 0) return out_({ ok: false, error: 'class' });
  const day = String(p.day || '').replace(/[^0-9-]/g, '').slice(0, 10);
  const player = clean_(p.player);
  if (p.action === 'done') return done_(cls, day, player, p);
  if (p.action === 'leave') return leave_(cls, player);
  return list_(cls, day);
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad' }); }
  if (b.action === 'push') return push_(clean_(b.player), clean_(b.phone), b.data);
  if (b.action === 'feedback') return feedback_(b);
  return out_({ ok: false, error: 'action' });
}

/* ---------- done today ---------- */
function list_(cls, day) {
  const rows = finished_().getDataRange().getValues().slice(1);
  const list = rows
    .filter(r => String(r[1]) === cls && String(r[0]) === 'd' + day)
    .map(r => ({ player: String(r[2]), name: String(r[3]), car: Number(r[4]) || 0 }))
    .sort((a, b) => a.name.localeCompare(b.name));   // A to Z, so nobody is "first"
  return out_({ ok: true, list: list });
}

function done_(cls, day, player, p) {
  if (!player || !day) return out_({ ok: false, error: 'missing' });
  const name = String(p.name || '').replace(/[<>]/g, '').trim().slice(0, 20) || 'Racer';
  const car = Math.max(0, Math.min(20, parseInt(p.car, 10) || 0));
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = finished_();
    const rows = sh.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === 'd' + day && String(rows[i][1]) === cls && String(rows[i][2]) === player) {
        sh.getRange(i + 1, 4, 1, 2).setValues([[name, car]]);
        found = true; break;
      }
    }
    if (!found) sh.appendRow(['d' + day, cls, player, name, car, new Date()]);
    if (Math.random() < 0.05) cleanup_();   // now and then, clear out old rows
  } finally {
    lock.releaseLock();
  }
  return list_(cls, day);
}

function leave_(cls, player) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = finished_();
    const rows = sh.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][1]) === cls && String(rows[i][2]) === player) sh.deleteRow(i + 1);
    }
  } finally {
    lock.releaseLock();
  }
  return out_({ ok: true });
}

/* ---------- progress shared between a child's phones ---------- */
function findRows_(sh, col, value) {
  return sh.getRange(1, col, sh.getLastRow(), 1).createTextFinder(value).matchEntireCell(true).findAll().map(r => r.getRow());
}

function push_(player, phone, data) {
  if (player.length < 4 || phone.length < 4) return out_({ ok: false, error: 'missing' });
  const json = JSON.stringify(data || {});
  if (json.length > 45000) return out_({ ok: false, error: 'big' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = progress_(), key = player + ':' + phone;
    const rows = findRows_(sh, 1, key);
    if (rows.length) sh.getRange(rows[0], 2, 1, 4).setValues([[player, phone, new Date(), json]]);
    else sh.appendRow([key, player, phone, new Date(), json]);
  } finally {
    lock.releaseLock();
  }
  return out_({ ok: true });
}

function pull_(player) {
  if (player.length < 4) return out_({ ok: false, error: 'missing' });
  const sh = progress_();
  const rows = findRows_(sh, 2, player).map(r => sh.getRange(r, 1, 1, 5).getValues()[0]);
  const phones = rows.map(r => {
    let data = {}; try { data = JSON.parse(r[4]); } catch (err) {}
    return { phone: String(r[2]), updated: r[3], data: data };
  });
  return out_({ ok: true, phones: phones });
}

/* ---------- feedback from parents ---------- */
function feedback_(b) {
  const msg = String(b.message || '').replace(/[<>]/g, '').trim().slice(0, 2000);
  if (!msg) return out_({ ok: false, error: 'empty' });
  const i = b.info || {}, s = v => String(v || '').replace(/[<>]/g, '').slice(0, 300);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = tab_(FEEDBACK, ['time', 'name', 'class', 'message', 'app version', 'mic', 'last words heard', 'phone']);
    if (sh.getLastRow() > 2000) sh.deleteRows(2, 500);   // keep the tab small
    sh.appendRow([new Date(), s(i.name), s(i.cls), msg, s(i.version), s(i.mic), s(i.heard), s(i.phone)]);
  } finally {
    lock.releaseLock();
  }
  return out_({ ok: true });
}

/* ---------- housekeeping ---------- */
function cleanup_() {
  const now = Date.now(), day = 864e5;
  const keep = (sh, test) => {
    const rows = sh.getDataRange().getValues();
    const kept = rows.slice(1).filter(test);
    if (kept.length === rows.length - 1) return;
    sh.getRange(2, 1, rows.length - 1, rows[0].length).clearContent();
    if (kept.length) sh.getRange(2, 1, kept.length, rows[0].length).setValues(kept);
  };
  keep(finished_(), r => now - new Date(String(r[0]).slice(1)).getTime() < KEEP_FINISHED_DAYS * day);
  // a phone's progress is kept while ANY phone of the same child is still used (an old phone's coins stay counted)
  const prog = progress_(), rows = prog.getDataRange().getValues().slice(1), active = {};
  rows.forEach(r => { if (now - new Date(r[3]).getTime() < KEEP_PROGRESS_DAYS * day) active[String(r[1])] = true; });
  keep(prog, r => active[String(r[1])]);
}
