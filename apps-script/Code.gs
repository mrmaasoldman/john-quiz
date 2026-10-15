/**
 * John Quiz progress tracker (Google Apps Script)
 * Receives a short progress summary from the John Quiz website and keeps
 * one row per student in the "Progress" sheet. Only the sheet's owner can see the sheet;
 * the website reads back just names, leaderboard scores and yesterday's top XP (see doGet).
 */
const KEY = 'jq-maas-2026';        // must match the key in the website
const SHEET_NAME = 'Progress';
const HEADERS = ['Student','Last active','Memorized verses','Memory goal','Goal progress',
  'Questions answered','Right answers','Accuracy','Questions mastered','Day streak','Best streak',
  'Answered today','Level','XP','Badges','Bosses beaten','Speed Round best','Daily Challenge today','Verse cards right',
  'Speed Round best (typed)','XP today','Date (today)','XP yesterday','Date (yesterday)'];

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
      num(d.badges), num(d.bosses), num(d.speed), d.daily || '', num(d.vright),
      num(d.speedTyped),
      num(d.xpToday), "'" + String(d.tdate || ''), num(d.xpYest), "'" + String(d.ydate || '')
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

// GET ?lb=1 returns the class leaderboards and Quizzer of the Day (names and scores only) for the website.
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.lb) return out('John Quiz tracker is running.');
  const sh = sheet_(), n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
  const people = rows.filter(r => String(r[0]).trim()).map(r => ({
    name: String(r[0]).trim(), questions: num(r[5]), streak: num(r[10]), bosses: num(r[15]), speed: num(r[19])
  }));
  const top = (k) => people.filter(p => p[k] > 0).sort((a, b) => b[k] - a[k]).slice(0, 10).map(p => ({ name: p.name, v: p[k] }));
  // Quizzer of the Day: most XP earned yesterday (uses the date each student's device reported)
  const tz = Session.getScriptTimeZone(), stz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const ds = v => v instanceof Date ? Utilities.formatDate(v, stz, 'yyyy-MM-dd') : String(v).replace(/^'/, '').trim();
  // 'yesterday' is checked in both time zones so a setting mismatch can't hide the winner
  const yests = [Utilities.formatDate(new Date(Date.now() - 864e5), tz, 'yyyy-MM-dd'), Utilities.formatDate(new Date(Date.now() - 864e5), stz, 'yyyy-MM-dd')];
  let qotd = null;
  rows.forEach(r => {
    const name = String(r[0]).trim(); if (!name) return;
    const xp = yests.includes(ds(r[21])) ? num(r[20]) : (yests.includes(ds(r[23])) ? num(r[22]) : 0);
    if (xp > 0 && (!qotd || xp > qotd.xp)) qotd = { name: name, xp: xp };
  });
  const body = { at: new Date().toISOString(), yesterday: yests[0], qotd: qotd, questions: top('questions'), streak: top('streak'), speed: top('speed'), bosses: top('bosses') };
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0 || sh.getLastColumn() < HEADERS.length) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#E6E0FF');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
  }
  return sh;
}
function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function out(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }
