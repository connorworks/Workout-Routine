// Workout Routine sync (starter kit, from v7).
// Paste this into your Sheet's Apps Script editor (Extensions > Apps Script), set PASSCODE, then deploy
// as a web app: Execute as Me, Who has access Anyone. See README.md in the starter folder.
// Receives sessions from the Workout Routine app and writes them to this Sheet,
// sends the app its program from the Program, Program · Weeks and Program · Lifts tabs,
// keeps training maxes on the Maxes tab in step with the app, and logs runs on the Running tab.
// Rows are matched on Block + Session ID, so re-sending a session updates it instead of duplicating it.

// The passcode you type into the app. Pick your own and keep the quotes.
// The app and the script refuse to talk until you change it from the example.
const PASSCODE = 'change-this-passcode';

function passOk(given) {
  return !!PASSCODE && PASSCODE !== 'change-this-passcode' && String(given || '') === PASSCODE;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    if (!passOk(body.passcode)) {
      return reply({ ok: false, error: 'passcode' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sessions = body.sessions || [];
    const weights = body.bodyweight || [];
    sessions.forEach(function (s) { writeSession(ss, s); });
    weights.forEach(function (b) { writeBodyweight(ss, b); });
    if (Array.isArray(body.maxes)) writeMaxes(ss, body.maxes);
    const runs = body.runs || [];
    runs.forEach(function (r) { writeRun(ss, r); });
    const res = { ok: true, sessions: sessions.length, bodyweight: weights.length, runs: runs.length };
    if (body.weekStarts) res.weekStarts = writeWeekStarts(ss, body.weekStarts);
    return reply(res);
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Opening the web address in a browser shows this, which is a quick way to check the deployment.
// With ?read=1&block=<name>&passcode=<passcode> it returns that block's rows, so a new phone can
// rebuild its history from the Sheet. With ?program=1&passcode=<passcode> it returns the program tabs.
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.program) {
    if (!passOk(p.passcode)) return reply({ ok: false, error: 'passcode' });
    const book = SpreadsheetApp.getActiveSpreadsheet();
    return reply({ ok: true, program: readTab(book, 'Program'), weeks: readTab(book, 'Program · Weeks', true), lifts: readTab(book, 'Program · Lifts'), maxes: readTab(book, 'Maxes') });
  }
  if (!p.read) return reply({ ok: true, app: 'Workout Routine sync' });
  if (!passOk(p.passcode)) return reply({ ok: false, error: 'passcode' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const block = String(p.block || '');
  return reply({
    ok: true,
    sessions: readBlock(ss, 'Sessions', block, 11),
    sets: readBlock(ss, 'Sets', block, 12),
    bodyweight: readBlock(ss, 'Bodyweight', block, 4),
    runs: readRuns(ss)
  });
}

// Rows of one tab for one block. Dates go out as yyyy-mm-dd; other date-times as ISO.
function readBlock(ss, name, block, cols) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const tz = ss.getSpreadsheetTimeZone();
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues()
    .filter(function (r) { return String(r[0]) === block; })
    .map(function (r) {
      return r.map(function (x, i) {
        if (!(x instanceof Date)) return x;
        return i === 2 ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : x.toISOString();
      });
    });
}

// A whole tab as shown in the Sheet (text, TRUE/FALSE). With dates, date cells go out as yyyy-mm-dd.
function readTab(ss, name, dates) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 1) return [];
  const rg = sh.getRange(1, 1, sh.getLastRow(), Math.max(1, sh.getLastColumn()));
  const shown = rg.getDisplayValues();
  if (!dates) return shown;
  const raw = rg.getValues(), tz = ss.getSpreadsheetTimeZone();
  return shown.map(function (r, i) {
    return r.map(function (x, j) { return raw[i][j] instanceof Date ? Utilities.formatDate(raw[i][j], tz, 'yyyy-MM-dd') : x; });
  });
}

// Training maxes from the app. m = [{ id, name, oneRm, tm }]. Rows are matched on ID; the columns are found by header.
function writeMaxes(ss, m) {
  let sh = ss.getSheetByName('Maxes');
  if (!sh) {
    sh = ss.insertSheet('Maxes');
    sh.getRange(1, 1, 1, 5).setValues([['Lift', 'ID', '1RM', 'Training max', 'Updated']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const col = function (name) { return head.indexOf(name) + 1; };
  const cId = col('id'), cName = col('lift'), cOne = col('1rm'), cTm = col('training max'), cAt = col('updated');
  if (!cId || !cTm) return;
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, cId, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  const now = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
  m.forEach(function (x) {
    const id = String(x.id || ''); if (!id) return;
    let r = ids.indexOf(id) + 2;
    if (r < 2) { r = sh.getLastRow() + 1; ids.push(id); sh.getRange(r, cId).setValue(id); if (cName) sh.getRange(r, cName).setValue(x.name || id); }
    const tm = Number(x.tm), one = Number(x.oneRm);
    if (isFinite(tm) && tm > 0 && Number(sh.getRange(r, cTm).getValue()) !== tm) { sh.getRange(r, cTm).setValue(tm); if (cAt) sh.getRange(r, cAt).setValue(now); }
    if (cOne && isFinite(one) && one > 0 && Number(sh.getRange(r, cOne).getValue()) !== one) { sh.getRange(r, cOne).setValue(one); if (cAt) sh.getRange(r, cAt).setValue(now); }
  });
}

// One run. r = { id, block, date, mi, time, pace, deleted }. Matched on Run ID; a deleted run's row is removed.
function writeRun(ss, r) {
  let sh = ss.getSheetByName('Running');
  if (!sh) {
    sh = ss.insertSheet('Running');
    sh.getRange(1, 1, 1, 7).setValues([['Block', 'Run ID', 'Date', 'Distance (mi)', 'Time', 'Pace (/mi)', 'Logged']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('E:F').setNumberFormat('@'); // keep 28:30 as text, not a clock time
  }
  const id = String(r.id || ''); if (!id) return;
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 2, last - 1, 1).getValues().map(function (x) { return String(x[0]); }) : [];
  const at = ids.indexOf(id) + 2;
  if (r.deleted) { if (at >= 2) sh.deleteRow(at); return; }
  const tz = ss.getSpreadsheetTimeZone();
  const row = [String(r.block || ''), id, String(r.date || ''), Number(r.mi) || 0, String(r.time || ''), String(r.pace || ''), Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm')];
  const target = at >= 2 ? at : last + 1;
  sh.getRange(target, 5, 1, 2).setNumberFormat('@');
  sh.getRange(target, 1, 1, row.length).setValues([row]);
}

// Every run, any block: [Block, Run ID, Date, Distance, Time].
function readRuns(ss) {
  const sh = ss.getSheetByName('Running');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues()
    .filter(function (r) { return r[1]; });
}

// A week off from the app: new start dates for the block's weeks. w = { block, starts: [{ week, start: 'yyyy-mm-dd' }] }.
// Only the Starts cells of that block's rows change.
function writeWeekStarts(ss, w) {
  const sh = ss.getSheetByName('Program · Weeks');
  if (!sh || sh.getLastRow() < 2 || !w || !Array.isArray(w.starts)) return 0;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const cB = head.findIndex(function (h) { return /^block/.test(h); }), cW = head.findIndex(function (h) { return /^week/.test(h); }),
    cS = head.findIndex(function (h) { return /^start/.test(h); });
  if (cW < 0 || cS < 0) return 0;
  const want = {};
  w.starts.forEach(function (x) { if (/^\d{4}-\d{2}-\d{2}$/.test(String(x.start))) want[String(x.week)] = String(x.start); });
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
  let n = 0, block = '';
  v.forEach(function (r, i) {
    if (cB >= 0 && r[cB]) block = r[cB];
    if (cB >= 0 && w.block && block && block !== w.block) return;
    const st = want[String(parseInt(r[cW], 10))];
    if (!st) return;
    const p = st.split('-').map(Number);
    sh.getRange(i + 2, cS + 1).setValue(new Date(p[0], p[1] - 1, p[2]));
    n++;
  });
  return n;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// One session: its row on Sessions, and all of its rows on Sets.
// s = { block, id, row: [Sessions columns], sets: [[Sets columns], ...] }
function writeSession(ss, s) {
  const tz = ss.getSpreadsheetTimeZone();
  const block = String(s.block), id = String(s.id);

  const sh = ss.getSheetByName('Sessions');
  const row = clean(s.row);
  row[10] = stamp(row[10], tz); // Last updated
  const r = findRow(sh, block, id);
  if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else { ensureRows(sh, sh.getLastRow() + 1); sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]); }

  const st = ss.getSheetByName('Sets');
  deleteRows(st, block, id);
  const sets = (s.sets || []).map(clean);
  if (sets.length) {
    const start = st.getLastRow() + 1;
    ensureRows(st, start + sets.length - 1);
    st.getRange(start, 11, sets.length, 2).insertCheckboxes(); // Missed, Done
    st.getRange(start, 1, sets.length, sets[0].length).setValues(sets);
  }
}

// One weigh-in. b = { block, week, row: [Block, Week, Date, Weight] }
function writeBodyweight(ss, b) {
  const sh = ss.getSheetByName('Bodyweight');
  const row = clean(b.row);
  const r = findRow(sh, String(b.block), String(b.week));
  if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else { ensureRows(sh, sh.getLastRow() + 1); sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]); }
}

// Deleting and re-adding a session's rows shrinks the sheet, so add spare rows whenever it runs short.
function ensureRows(sh, lastNeeded) {
  const max = sh.getMaxRows();
  if (lastNeeded > max) sh.insertRowsAfter(max, lastNeeded - max + 500);
}

// Row number whose first two columns match, or 0.
function findRow(sh, a, b) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const keys = sh.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === a && String(keys[i][1]) === b) return i + 2;
  }
  return 0;
}

// Removes every row whose first two columns match, working bottom-up in runs.
function deleteRows(sh, a, b) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const keys = sh.getRange(2, 1, last - 1, 2).getValues();
  const hit = function (i) { return String(keys[i][0]) === a && String(keys[i][1]) === b; };
  let i = keys.length - 1;
  while (i >= 0) {
    if (hit(i)) {
      let j = i;
      while (j - 1 >= 0 && hit(j - 1)) j--;
      sh.deleteRows(j + 2, i - j + 1);
      i = j - 1;
    } else {
      i--;
    }
  }
}

// Nulls become blanks, and text that starts with "=" can't turn into a formula.
function clean(row) {
  return (row || []).map(function (v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string' && v.charAt(0) === '=') return "'" + v;
    return v;
  });
}

// ISO time from the phone, written in the Sheet's time zone.
function stamp(v, tz) {
  if (typeof v !== 'string' || !/^\d{4}-\d{2}-\d{2}T/.test(v)) return v || '';
  return Utilities.formatDate(new Date(v), tz, 'yyyy-MM-dd HH:mm:ss');
}
