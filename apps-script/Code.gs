/**
 * John Quiz progress tracker (Google Apps Script)
 * Receives a short progress summary from the John Quiz website and keeps
 * one row per student in the "Progress" sheet. Only the sheet's owner can see the sheet;
 * the website reads back just names, teams, high scores and yesterday's top XP (see doGet).
 * Full progress backups are saved as text files in the teacher's Drive folder "John Quiz Backups".
 */
const KEY = 'jq-maas-2026';        // must match the key in the website
const SHEET_NAME = 'Progress';
const HEADERS = ['Student','Last active','Memorized verses','Memory goal','Goal progress',
  'Questions answered','Right answers','Accuracy','Questions mastered','Day streak','Best streak',
  'Answered today','Level','XP','Badges','Bosses beaten','Speed Round best','Daily Challenge today','Verse cards right',
  'Speed Round best (typed)','XP today','Date (today)','XP yesterday','Date (yesterday)',
  'Team','XP this week','Week of','Avatar'];

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out('bad json'); }
  if (!d || d.key !== KEY || !d.name) return out('ignored');
  if (d.type === 'backup') return saveBackup_(String(d.name).trim().slice(0, 40), String(d.code || ''));
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
      num(d.xpToday), "'" + String(d.tdate || ''), num(d.xpYest), "'" + String(d.ydate || ''),
      String(d.team || '').trim().replace(/\s+/g, ' ').slice(0, 40), num(d.xpWeek), "'" + String(d.wk || ''),
      /^[a-z.]{0,60}$/.test(String(d.av || '')) ? String(d.av || '') : ''
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

// GET ?lb=1 returns the high scores, team list, Quizzer of the Day and Team of the Day (names and scores only) for the website.
// Full progress backups go to a Drive folder only the teacher can see: one text file per student.
// To restore a student: open their file, copy everything, send it to them; they paste it under Progress > Paste a code.
const BACKUP_FOLDER = 'John Quiz Backups';

// Run this once from the Apps Script editor (choose "setupBackups" in the menu next to Run, then click Run).
// It asks Google for permission to use your Drive and creates the backups folder.
function setupBackups() {
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
  Logger.log('Backups folder is ready: ' + folder.getUrl());
}
function saveBackup_(name, code) {
  if (!name || !/^JQZ1:|^[A-Za-z0-9+/=]+$/.test(code) || code.length > 2000000) return out('ignored');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const it = DriveApp.getFoldersByName(BACKUP_FOLDER);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
    const fname = name.replace(/[\\/:*?"<>|]/g, '') + ' - progress code.txt';
    const files = folder.getFilesByName(fname);
    if (files.hasNext()) files.next().setContent(code); else folder.createFile(fname, code, MimeType.PLAIN_TEXT);
  } finally { lock.releaseLock(); }
  return out('ok');
}

function doGet(e) {
  if (!e || !e.parameter || !e.parameter.lb) return out('John Quiz tracker is running.');
  const sh = sheet_(), n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
  // Everything runs on Central time: the website reports dates in Central time too.
  const tz = 'America/Chicago', stz = 'America/Chicago';
  const ds = v => v instanceof Date ? Utilities.formatDate(v, stz, 'yyyy-MM-dd') : String(v).replace(/^'/, '').trim();
  const fmt = (ms, z) => Utilities.formatDate(new Date(ms), z, 'yyyy-MM-dd');
  const both = ms => [fmt(ms, tz), fmt(ms, stz)];
  const now = Date.now(), todays = both(now), yests = both(now - 864e5);
  const monday = z => fmt(now - (Number(Utilities.formatDate(new Date(now), z, 'u')) - 1) * 864e5, z);  // weeks start Monday
  const weeks = [monday(tz), monday(stz)];

  const teamNames = {};   // lower-case -> first spelling seen
  const people = rows.filter(r => String(r[0]).trim()).map(r => {
    const t = String(r[24] || '').trim().replace(/\s+/g, ' ');
    if (t && !teamNames[t.toLowerCase()]) teamNames[t.toLowerCase()] = t;
    return {
      name: String(r[0]).trim(), team: t ? teamNames[t.toLowerCase()] : '',
      questions: num(r[5]), streak: num(r[10]), bosses: num(r[15]), speed: num(r[19]),
      day: todays.includes(ds(r[21])) ? num(r[20]) : 0,
      yest: yests.includes(ds(r[21])) ? num(r[20]) : (yests.includes(ds(r[23])) ? num(r[22]) : 0),
      week: weeks.includes(ds(r[26])) ? num(r[25]) : 0,
      av: String(r[27] || ''),
      raw: { name: String(r[0]).trim(), team: t ? teamNames[t.toLowerCase()] : '', av: String(r[27] || ''),
             td: ds(r[21]), xt: num(r[20]), yd: ds(r[23]), xy: num(r[22]), wk: ds(r[26]), xw: num(r[25]) }
    };
  });
  const top = k => people.filter(p => p[k] > 0).sort((a, b) => b[k] - a[k]).slice(0, 10).map(p => ({ name: p.name, team: p.team, av: p.av, v: p[k] }));
  const teamTop = k => {
    const m = {};
    people.forEach(p => { if (!p.team || !(p[k] > 0)) return; const t = m[p.team] || (m[p.team] = { name: p.team, v: 0, n: 0 }); t.v += p[k]; t.n++; });
    return Object.keys(m).map(x => m[x]).sort((a, b) => b.v - a.v).slice(0, 10);
  };
  // Quizzer of the Day and Team of the Day: most XP earned yesterday
  const q = top('yest')[0], t = teamTop('yest')[0];
  const body = {
    at: new Date().toISOString(), yesterday: yests[0],
    qotd: q ? { name: q.name, xp: q.v, av: q.av } : null,
    totd: t ? { name: t.name, xp: t.v, n: t.n } : null,
    teams: Object.keys(teamNames).map(x => teamNames[x]).sort(),
    questions: top('questions'), streak: top('streak'), speed: top('speed'), bosses: top('bosses'),
    day: top('day'), week: top('week'), teamDay: teamTop('day'), teamWeek: teamTop('week'),
    // Each student's XP with the dates their own device reported, so the website can decide
    // "today", "yesterday" and "this week" by the student's clock instead of this script's time zone.
    rows: people.filter(p => p.raw.xt > 0 || p.raw.xy > 0 || p.raw.xw > 0).map(p => p.raw)
  };
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
