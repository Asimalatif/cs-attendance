/**
 * MONTHLY TEST + GRAND TEST WORKBOOKS  (needs GrandTest.gs in the same project)
 *
 * 1) The app's Monthly test now reads and writes the real Monthly Test workbooks
 *    ("September - MONTHLY TEST", "October - MONTHLY TEST" ...), same layout as the Grand Test workbook:
 *    one sheet per section, row 3 = subject, row 4 = Total Marks / Obtained Marks, students from row 5.
 *    Marks entered in the old Monthly Test portal now show in the app, and marks saved in the app go into
 *    the workbook (and, as before, into the M.T column of the attendance sheet).
 * 2) "Sheets vs strength sheet" for these workbooks: every section sheet is compared with
 *    Transform-ALL-COLLEGE-STRENGTH (missing students, extra students), with Add / Remove buttons in the app.
 *
 * Install (attendance Apps Script project):
 *   a) + > Script > name it "MonthlyBooks" > paste this file.
 *   b) doPost, under the gt_ line:
 *        if (/^bk_/.test(String(req.action || ''))) return json_(bkRoute_(user, req));
 *   c) getSheetData_: find
 *        var my = monthYear_(meta.mon);
 *        var title = String(sheet.getRange(1, 1).getValue() || '').trim();
 *      and put this line ABOVE them:
 *        try { mt = mtFromBook_(meta, mt); } catch (e) { Logger.log('Monthly test workbook skipped for ' + name + ': ' + e.message); }
 *   d) saveMt_: find
 *        return { at: Date.now(), mt: true, wrote: wrote, total: total };
 *      and put this line ABOVE it:
 *        mtToBook_(I, rec, total);
 *   e) Run mtCheckBooks once from the editor (it asks for Drive permission and logs which workbook each month uses).
 *   f) Deploy > Manage deployments > Edit > New version > Deploy.
 */

// Monthly Test workbook of each month. If an ID is wrong or empty, the workbook is found by its name
// ("September - MONTHLY TEST" ...) in Google Drive, and the ID found is remembered.
var MT_BOOKS = {
  SEPT: '11LaIEImp-zK11dkH0zGOuYm9SvhUTMNEfPosDvTEMjE',
  OCT:  '1b3H60R1zHa7V3zm3vojEtGHJIoPOfjVy08DndPFPbLw',
  NOV:  '1XfLdVOXXtMsZXres7mJFT8WDVe2G3cPmZnyKuTK4W0w',
  DEC:  '',
  JAN:  ''
};
var MT_BOOK_SUFFIX = ' - MONTHLY TEST';
// Monthly attendance workbooks (one sheet per section, per-subject lectures). Only checked against the strength sheet.
var MA_BOOKS = {
  SEPT: '1Gxjgcb82GZqFtNSMbK2QpVmaK2uppFFKtel6T93L-gk',
  OCT:  '1rs_jrVRislaz9LUMIRwGLdwDv24FTvFXS0L5GKi2IzM',
  NOV:  '1KtN4MBqfapaxqT1WafZJRLOzovFfMrEE00Jq7hlSD8A',
  DEC:  '1Y1bk_kmh4--jqOPNcTkdDd8X5a4iLAFi-EcQgbgis-4',
  JAN:  '16nZoP2pGAvtFhADo2sNqqvrS-0OEY4O-9CgBrrqULpg'
};
var BK_SKIP = /report|combine|log|settings|missing|summary|backup|copy of/i;   // tabs that are not section sheets

/* ================= finding the workbooks ================= */
function mtMonthName_(tok) { var i = SEASON_MONTH[String(tok).toUpperCase()]; return i === undefined ? String(tok) : MONTH_NAMES[i].charAt(0) + MONTH_NAMES[i].slice(1).toLowerCase(); }
/** the Monthly Test workbook of a month ('SEPT' ...), or null when there is none */
function mtBook_(tok) {
  tok = String(tok || '').toUpperCase();
  var props = PropertiesService.getScriptProperties(), key = 'MT_BOOK_' + tok;
  var ids = [props.getProperty(key), MT_BOOKS[tok]].filter(Boolean);
  for (var i = 0; i < ids.length; i++) { try { return SpreadsheetApp.openById(ids[i]); } catch (e) {} }
  try {
    var it = DriveApp.getFilesByName(mtMonthName_(tok) + MT_BOOK_SUFFIX);
    if (it.hasNext()) { var f = it.next(); props.setProperty(key, f.getId()); return SpreadsheetApp.openById(f.getId()); }
  } catch (e) {}
  return null;
}
/** Run once from the editor: shows which workbook each month uses. */
function mtCheckBooks() {
  SEASON.forEach(function (t) {
    var b = mtBook_(t);
    Logger.log(t + ': ' + (b ? b.getName() + '  (' + b.getId() + ')' : 'no Monthly Test workbook found'));
  });
  var g = null; try { g = gtBook_(); } catch (e) {}
  Logger.log('GRAND TEST: ' + (g ? g.getName() : 'not found, check GT_CFG.BOOK_ID'));
}

/* ================= shared: reading any section sheet ================= */
/** Like gtSheet_, but also reads sheets without an S.No column (header row = the row that has Enroll#) */
function bkSheet_(sh) {
  var L = gtSheet_(sh);
  if (L) { L.sno = 0; return L; }
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return null;
  var v = sh.getRange(1, 1, lr, lc).getValues(), hi = -1;
  for (var i = 0; i < Math.min(v.length, 8); i++) { if (v[i].some(function (h) { return gtK_(h) === 'ENROLL'; })) { hi = i; break; } }
  if (hi < 0) return null;
  var head = v[hi].map(gtK_), col = function (names) { for (var k = 0; k < names.length; k++) { var x = head.indexOf(names[k]); if (x >= 0) return x; } return -1; };
  return { sh: sh, v: v, hi: hi, lc: lc, sno: col(['SNO', 'SRNO']), enroll: col(['ENROLL']), name: col(['STUDENTNAME', 'NAME']),
           father: col(['FATHERNAME', 'FATHER']), cs: col(['CLASSSEC', 'CLASSSECTION', 'PCS']), subjects: {} };
}
/** true when row i (0-based in L.v) has any typed value after the name columns (formulas do not count) */
function bkHasData_(L, i, F) {
  var from = Math.max(L.enroll, L.name, L.father, L.cs, L.sno) + 1, row = L.v[i] || [], fr = F ? F[i] || [] : [];
  for (var c = from; c < row.length; c++) if (gtS_(row[c]) !== '' && !fr[c]) return true;
  return false;
}

/* ================= shared: adding a student row to a section sheet ================= */
/** m = master row [roll, name, classSec, subjects, father]. Copies the look and formulas of the last student row. */
function bkAppend_(sh, L, m) {
  var last = L.hi + 1;
  for (var i = L.v.length - 1; i > L.hi; i--) { if (gtS_(L.v[i][L.enroll])) { last = i + 1; break; } }
  var sc = L.sno === undefined ? 0 : L.sno, maxSno = 0;
  if (sc >= 0) for (var j = L.hi + 1; j < L.v.length; j++) { var n = Number(L.v[j][sc]); if (n > maxSno) maxSno = n; }
  var row = last + 1;
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 5);
  if (last > L.hi + 1) {
    sh.getRange(last, 1, 1, L.lc).copyTo(sh.getRange(row, 1, 1, L.lc), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sh.getRange(last, 1, 1, L.lc).getFormulasR1C1()[0].forEach(function (fx, c) { if (fx) sh.getRange(row, c + 1).setFormulaR1C1(fx); });
  }
  var put = function (c, v) { if (c >= 0) sh.getRange(row, c + 1).setValue(v); };
  put(sc, maxSno + 1); put(L.enroll, m[0]); put(L.name, m[1]); put(L.father, m[4] || ''); put(L.cs, m[2] || '');
  sh.getRange(row, 1, 1, L.lc).setFontLine('none');
  var line = []; for (var k = 0; k < L.lc; k++) line.push('');
  if (sc >= 0) line[sc] = maxSno + 1; line[L.enroll] = m[0];
  while (L.v.length < row - 1) L.v.push(line.map(function () { return ''; }));
  L.v[row - 1] = line;                                   // so a second student added in the same call goes below her
  return row;
}

/* ================= 1) Monthly test <-> Monthly Test workbook ================= */
/** marks of this class + subject from the Monthly Test workbook; falls back to the attendance M.T column */
function mtFromBook_(meta, mt) {
  var book = mtBook_(meta.mon); if (!book) return mt;
  var sh = gtFindSheet_(book, meta.cls); if (!sh) return mt;
  var L = gtSheet_(sh); if (!L) return mt;
  var S = L.subjects[gtK_(meta.sub)]; if (!S) return mt;
  var out = { total: 0, marks: {}, book: book.getName() }, totals = {};
  for (var i = L.hi + 1; i < L.v.length; i++) {
    var roll = gtS_(L.v[i][L.enroll]), ob = gtS_(L.v[i][S.obtained]), t = gtS_(L.v[i][S.total]);
    if (!roll) continue;
    if (ob !== '') { out.marks[roll] = ob; if (t) totals[t] = (totals[t] || 0) + 1; }
  }
  var best = Object.keys(totals).sort(function (a, b) { return totals[b] - totals[a]; })[0];
  out.total = Number(best) || Number(mt && mt.total) || 0;
  if (!Object.keys(out.marks).length) { (mt || {}).book = out.book; return mt; }      // nothing in the workbook yet
  Object.keys((mt && mt.marks) || {}).forEach(function (r) { if (out.marks[r] === undefined) out.marks[r] = mt.marks[r]; });
  return out;
}
/** writes the marks just saved in the attendance sheet into the Monthly Test workbook too */
function mtToBook_(I, rec, total) {
  var meta = I.meta || {}; if (!meta.cls) return;
  var book = mtBook_(meta.mon); if (!book) return;                 // no Monthly Test workbook for this month
  var sh = gtFindSheet_(book, meta.cls); if (!sh) return;
  var L = gtSheet_(sh); if (!L) return;
  var S = L.subjects[gtK_(meta.sub)];
  if (!S) throw new Error('"' + book.getName() + '" sheet ' + sh.getName() + ' has no Total / Obtained columns for ' + meta.sub + '. Tell the admin.');
  var marks = rec.marks || {}, rowOf = {}, changes = [];
  for (var i = L.hi + 1; i < L.v.length; i++) { var r = gtS_(L.v[i][L.enroll]).toUpperCase(); if (r && !rowOf[r]) rowOf[r] = i + 1; }
  var M = null;
  Object.keys(marks).forEach(function (roll) {
    var s = gtS_(marks[roll]).toUpperCase(), v = s === '' ? '' : (s === 'A' || s === 'AB' || s === 'ABSENT') ? 'A' : Number(s);
    var row = rowOf[roll.toUpperCase()];
    if (!row) {                                                      // new student: add her row
      if (v === '') return;
      M = M || rosterModel_();
      var m = M.byRoll[enr_(roll)]; if (!m) return;
      row = bkAppend_(sh, L, m); rowOf[roll.toUpperCase()] = row;
    }
    var was = gtS_(sh.getRange(row, S.obtained + 1).getValue());
    sh.getRange(row, S.total + 1, 1, 2).setValues([[v === '' ? '' : total, v]]);
    if (String(was).toUpperCase() !== String(v).toUpperCase()) changes.push([roll, was, v, row]);
  });
  SpreadsheetApp.flush();
  if (changes.length) {
    try {
      var lg = book.getSheetByName('WebApp Logs');
      if (!lg) { lg = book.insertSheet('WebApp Logs'); lg.appendRow(['Timestamp', 'Username', 'Enroll#', 'Student', 'Section', 'Subject', 'Teacher', 'Cell Link', 'Before', 'After']); }
      var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy hh:mm a'), base = book.getUrl() + '#gid=' + sh.getSheetId() + '&range=';
      var rows = changes.map(function (c) {
        var a1 = sh.getRange(c[3], S.obtained + 1).getA1Notation();
        return [when, (I.user && I.user.username) || '', c[0], '', sh.getName(), S.name, meta.teacher || '', '=HYPERLINK("' + base + a1 + '","' + sh.getName() + '!' + a1 + '")', c[1], c[2]];
      });
      lg.getRange(lg.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    } catch (e) {}
  }
}

/* ================= 2) workbook sheets vs strength sheet ================= */
function bkRoute_(user, req) {
  var s = scopeOf_(user);
  if (!s || !s.all) return { ok: false, error: 'Only the admin and the principal can do this.' };
  switch (String(req.action)) {
    case 'bk_list':  return bkList_();
    case 'bk_check': return bkCheck_(req.book);
    case 'bk_fix':   return bkFix_(user, req.book, req.sheet, req.roll, req.op);
    case 'bk_batch': return bkBatch_(user, req);
  }
  return { ok: false, error: 'UNKNOWN_ACTION' };
}
/** which workbooks can be checked: the Monthly Test months that exist, and the Grand Test */
function bkList_() {
  var out = [];
  SEASON.forEach(function (t) { if (!MA_BOOKS[t]) return; try { var b = SpreadsheetApp.openById(MA_BOOKS[t]); out.push({ key: 'MA_' + t, group: 'Monthly attendance', label: mtMonthName_(t) + ' attendance', name: b.getName() }); } catch (e) {} });
  SEASON.forEach(function (t) { var b = mtBook_(t); if (b) out.push({ key: t, group: 'Monthly test', label: mtMonthName_(t) + ' monthly test', name: b.getName() }); });
  try { var g = gtBook_(); out.push({ key: 'GT', group: 'Grand test', label: 'Grand test', name: g.getName() }); } catch (e) {}
  try { if (typeof dtBook_ === 'function') { var dt = dtBook_(); out.push({ key: 'DT', group: 'Detention exam', label: 'Detention exam', name: dt.getName() }); } } catch (e) {}
  return { ok: true, books: out };
}
function bkBook_(key) {
  key = String(key || '').toUpperCase();
  if (key === 'GT') return { book: gtBook_(), label: 'Grand test' };
  if (key === 'DT') return { book: dtBook_(), label: 'Detention exam' };
  if (key.indexOf('MA_') === 0) {
    var t = key.slice(3);
    if (!MA_BOOKS[t]) throw new Error('No monthly attendance workbook for ' + mtMonthName_(t) + '.');
    return { book: SpreadsheetApp.openById(MA_BOOKS[t]), label: mtMonthName_(t) + ' monthly attendance' };
  }
  var b = mtBook_(key);
  if (!b) throw new Error('No Monthly Test workbook for ' + mtMonthName_(key) + '.');
  return { book: b, label: mtMonthName_(key) + ' monthly test' };
}
/** students the section sheet should have: everyone of that Class-Sec ("F.A-1 Optional Subjects": all of F.A-1) */
function bkWanted_(cls, RM) {
  var M = getMaster_(), base = gtK_(baseClass_(cls)), opt = base !== gtK_(cls), out = {};
  M.rows.forEach(function (r) { var c = gtK_(r[2]); if (opt ? (c === base || c.indexOf(base) === 0) : c === base) out[enr_(r[0])] = r; });
  return out;
}
function bkClasses_() {
  var seen = {}, out = [];
  getAssignments_({ teacher: '', username: ADMIN_USER }).forEach(function (a) { var k = gtK_(a.cls); if (!seen[k]) { seen[k] = a.cls; out.push(a.cls); } });
  return out;
}
function bkCheck_(key) {
  var B = bkBook_(key), start = Date.now(), RM = rosterModel_(), diffs = [], checked = 0;
  var classes = {}; bkClasses_().forEach(function (c) { classes[gtK_(c)] = c; });
  var noSheet = [];
  var sheets = {}; B.book.getSheets().forEach(function (s) { sheets[gtK_(s.getName())] = s; });
  Object.keys(classes).forEach(function (k) { if (!sheets[k]) noSheet.push(classes[k]); });
  B.book.getSheets().forEach(function (sh) {
    if (Date.now() - start > 300000) return;
    var cls = classes[gtK_(sh.getName())];
    if (!cls || BK_SKIP.test(sh.getName())) return;
    var L = bkSheet_(sh);
    if (!L) { diffs.push({ sheet: sh.getName(), cls: cls, sub: '', teacher: '', missing: [], extra: [], off: [], error: 'No Enroll# header row.' }); return; }
    var F = null;
    checked++;
    var want = bkWanted_(cls, RM), struck = gtStruck_(L), inSheet = {}, d = { sheet: sh.getName(), cls: cls, sub: '', teacher: '', missing: [], extra: [], off: [] };
    for (var i = L.hi + 1; i < L.v.length; i++) {
      var roll = gtS_(L.v[i][L.enroll]), e = enr_(roll); if (!e) continue;
      inSheet[e] = 1;
      if (struck[i - L.hi - 1]) continue;                                    // cancelled: leave as it is
      if (!want[e]) {
        F = F || sh.getRange(1, 1, L.v.length, L.lc).getFormulas();
        var m = RM.byRoll[e], marks = bkHasData_(L, i, F);
        d.extra.push({ roll: roll, name: gtS_(L.v[i][L.name]), father: L.father >= 0 ? gtS_(L.v[i][L.father]) : '', pcs: m ? m[2] : '', marks: marks,
                       why: m ? 'in the strength sheet as ' + m[2] : 'not in the strength sheet' });
      }
    }
    Object.keys(want).forEach(function (e) { if (!inSheet[e]) { var r = want[e]; d.missing.push({ roll: r[0], name: r[1], father: r[4] || '', pcs: r[2] }); } });
    if (d.missing.length || d.extra.length) diffs.push(d);
  });
  return { ok: true, book: String(key).toUpperCase(), label: B.label + ' (' + B.book.getName() + ')', checked: checked, total: checked, done: true, diffs: diffs, noSheet: noSheet };
}
/** op 'add' (missing student) or 'remove' (extra student: row deleted if she has no marks, otherwise struck through) */
function bkFix_(user, key, sheetName, roll, op) {
  var B = bkBook_(key), sh = B.book.getSheetByName(String(sheetName));
  if (!sh) return { ok: false, error: 'Sheet not found: ' + sheetName };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var L = bkSheet_(sh); if (!L) return { ok: false, error: 'No header row in ' + sheetName };
    var row = 0;
    for (var i = L.hi + 1; i < L.v.length; i++) if (enr_(L.v[i][L.enroll]) === enr_(roll)) { row = i + 1; break; }
    var res;
    if (op === 'add') {
      if (row) return { ok: true, result: 'already there' };
      var m = rosterModel_().byRoll[enr_(roll)]; if (!m) return { ok: false, error: roll + ' is not in the strength sheet.' };
      bkAppend_(sh, L, m); res = 'added';
    } else if (op === 'remove') {
      if (!row) return { ok: true, result: 'already gone' };
      var F = [], fr = sh.getRange(row, 1, 1, L.lc).getFormulas()[0]; F[row - 1] = fr;
      var marks = bkHasData_(L, row - 1, F);
      if (marks) { sh.getRange(row, 1, 1, Math.max(5, L.cs + 1, L.name + 1)).setFontLine('line-through').setFontColor('#999999'); res = 'struck through (she has marks or attendance)'; }
      else { sh.deleteRow(row); res = 'removed'; }
    } else return { ok: false, error: 'Unknown action.' };
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('gt_progress_v1'); } catch (e) {}
    return { ok: true, sheet: sheetName, roll: roll, op: op, result: res };
  } finally { lock.releaseLock(); }
}

/** Many fixes of one sheet in one call. req: { kind: 'daily' | 'book', book, sheet, ops: [[op, roll], ...] }
 *  daily = a daily attendance sheet (uses rosterFix_ of Code.gs: add / remove / on); book = a section-sheet workbook.
 *  Stops after about 4 minutes and returns what is left, so the app sends the rest again. */
function bkBatch_(user, req) {
  var start = Date.now(), ops = (req.ops || []).slice(), done = [];
  while (ops.length && Date.now() - start < 240000) {
    var o = ops.shift(), r;
    try { r = req.kind === 'daily' ? rosterFix_(user, req.sheet, o[1], o[0]) : bkFix_(user, req.book, req.sheet, o[1], o[0]); }
    catch (e) { r = { ok: false, error: String(e.message || e) }; }
    done.push({ op: o[0], roll: o[1], ok: !!(r && r.ok), result: r && (r.result || r.error) });
  }
  return { ok: true, done: done, left: ops };
}
