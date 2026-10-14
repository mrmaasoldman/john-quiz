/**
 * John Quiz progress tracker (Google Apps Script)
 * Receives a short progress summary from the John Quiz website and keeps
 * one row per student in the "Progress" sheet. Only the sheet's owner can see it.
 */
const KEY = 'jq-maas-2026';        // must match the key in the website
const SHEET_NAME = 'Progress';
const HEADERS = ['Student','Last active','Memorized verses','Memory goal','Goal progress',
  'Questions answered','Right answers','Accuracy','Questions mastered','Day streak','Best streak',
  'Answered today','Level','XP','Badges','Bosses beaten','Speed Round best','Daily Challenge today','Verse cards right'];

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out('bad json'); }
  if (!d || d.key !== KEY || !d.name) return out('ignored');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const name = String(d.name).trim().slice(0, 40);
    const row = [
      name, new Date(), num(d.mem),
      d.goal ? 'John ' + String(d.goal).slice(0, 60) : '(no goal set)',
      d.goalN ? Math.round(100 * num(d.goalMem) / num(d.goalN)) / 100 : '',
      num(d.att), num(d.right), d.att ? Math.round(100 * num(d.right) / num(d.att)) / 100 : '',
      num(d.solid), num(d.streak), num(d.best), num(d.today), num(d.level), num(d.xp),
      num(d.badges), num(d.bosses), num(d.speed), d.daily || '', num(d.vright)
    ];
    const names = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).trim().toLowerCase());
    const i = names.indexOf(name.toLowerCase());
    const r = i >= 0 ? i + 2 : sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    sh.getRange(r, 5).setNumberFormat('0%');
    sh.getRange(r, 8).setNumberFormat('0%');
    sh.getRange(r, 2).setNumberFormat('mmm d, h:mm am/pm');
  } finally { lock.releaseLock(); }
  return out('ok');
}

function doGet() { return out('John Quiz tracker is running.'); }

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#E6E0FF');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
  }
  return sh;
}
function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function out(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }
