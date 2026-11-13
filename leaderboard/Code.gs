// Word Racer "finished today" list. Paste this into the Google Sheet: Extensions > Apps Script.
// It keeps one row per child per day when they finish all their races: first name and car only.
// No scores and no ranking.

const CLASSES = ['2021'];   // class codes that may use the list
const SHEET = 'Finished';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['day', 'class', 'player', 'name', 'car', 'time']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cls = String(p.class || '').toUpperCase().trim();
  if (CLASSES.indexOf(cls) < 0) return out_({ ok: false, error: 'class' });
  const day = String(p.day || '').replace(/[^0-9-]/g, '').slice(0, 10);
  const player = String(p.player || '').replace(/[^a-z0-9]/gi, '').slice(0, 20);
  if (p.action === 'done') return done_(cls, day, player, p);
  if (p.action === 'leave') return leave_(cls, player);
  return list_(cls, day);
}

function list_(cls, day) {
  const rows = sheet_().getDataRange().getValues().slice(1);
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
    const sh = sheet_();
    const rows = sh.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === 'd' + day && String(rows[i][1]) === cls && String(rows[i][2]) === player) {
        sh.getRange(i + 1, 4, 1, 2).setValues([[name, car]]);
        found = true; break;
      }
    }
    if (!found) sh.appendRow(['d' + day, cls, player, name, car, new Date()]);
  } finally {
    lock.releaseLock();
  }
  return list_(cls, day);
}

function leave_(cls, player) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const rows = sh.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][1]) === cls && String(rows[i][2]) === player) sh.deleteRow(i + 1);
    }
  } finally {
    lock.releaseLock();
  }
  return out_({ ok: true });
}
