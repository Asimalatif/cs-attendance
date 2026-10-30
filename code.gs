var REPORTS_SHEET = 'Chief Proctor Reports';
var PROCTORS_SHEET = 'Chief Proctors';   // columns: Id | Name | Phone | Dept | Desig   (phone as digits with country code, e.g. 923001234567)
var REPORT_HEAD = ['Report ID','Received','Teacher','Username','Student','Roll No','Class','Subject','Month','Incident date','Lecture','Reason','Severity','Statement','Attendance','Sent to','Via','Guest','Status','Chief Proctor remark'];


// One attendance workbook per month
var BOOK_IDS = {
  SEPT: '1tKdOESL-cqP9IXafk8IbMNv8Y9vRG59a5gPJyCRSCXM',
  OCT:  '10X8dQGSMGvq0S369d6Go2jY2BVKtbU0XfQre8Xi7PBI',
  NOV:  '1B0KJVbMKiXjr8UkkUbsMwmaeZxbg74q59Hldgsskjcc',
  DEC:  '1O0BiNIys7EHzGwLFpfq0tQLW4jFcd2NxkJSkwxBPb3k',
  JAN:  '1H_1mzuNUuJ_sDfQTXmhU0p8MBYul6yb2eNFq0tqYxjE'
};
// Workbook that holds the sheets: Users, Timetable, Transform-ALL-COLLEGE-STRENGTH
var USERS_SS_ID = '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY';

var ADMIN_USER      = 'jcw@uop.edu.pk';  // sees every section
var PROGRAM_FILTER  = '';                // e.g. 'FCS' to only use that Program from the Timetable. '' = all
var HEADER_ROW = 3;
var FIRST_ROW  = 4;

var DAY_COL1  = 7;                       // G = day 1
var LAST_DAY  = DAY_COL1 + 30;           // AK = day 31
var TA_COL    = LAST_DAY + 1;            // AL = T.A
var MT_COL    = LAST_DAY + 2;            // AM = M.T
var TITLE_PREFIX = 'JCW Attendance - ';

var VALID = { P: 1, A: 1, L: 1, B: 1 };  // marks that may be written. ONLY "P" is counted.

// Period-wise attendance
var PD_SHEET   = 'Period Data';          // hidden sheet in every month workbook (real marks of every period)
var MAX_PERIOD = 8;                      // periods 1..8
var NONP       = ['L', 'A', 'B'];        // student has NO P in any period: this letter wins first (most lenient first)

var SEASON       = ['SEPT', 'OCT', 'NOV', 'DEC', 'JAN'];
var SEASON_MONTH = { SEPT: 8, OCT: 9, NOV: 10, DEC: 11, JAN: 0 };
var MONTH_NAMES  = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
var CACHE_TTL    = 600;
var CACHE_KEYS   = ['users_att2', 'tt_att3', 'master_att3', 'ptimes'];

// Big workbook settings
var SPARE_ROWS = 2;          // empty rows kept under the last student (so a teacher can still add a line by hand)
var BATCH_MS   = 240000;     // bulk jobs stop after 4 minutes and continue by themselves (Apps Script stops at 6)

// Disabled students (changed subject / section)
var OFF_TAG = 'LEFT SUBJECT';   // note on the Roll No cell of a disabled student

// Student remarks
var REMARKS_SHEET = 'Student Remarks';
var REMARK_HEAD   = ['Key', 'Teacher', 'Class', 'Roll No', 'Student', 'Tags', 'Latest comment', 'Comments', 'Data (do not edit)', 'Updated (ms)', 'Updated'];
var REMARK_MAX_LOG = 50, REMARK_MAX_TAGS = 10;

function doGet() {
  return json_({ ok: true, service: 'JCW Daily Attendance backend v6.6' });
}

function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) {
    return json_({ ok: false, error: 'BAD_REQUEST' });
  }

  var user = authenticate_(req.username, req.password);
  if (!user) return json_({ ok: false, error: 'BAD_LOGIN' });

  try {
    if (req.action === 'sheets') {
  return json_({ ok: true, teacher: user.teacher, admin: isAdmin_(user), sheets: listSheets_(user), proctors: getProctors_(), view: scopeName_(user) });
}
    if (req.action === 'whoami')       return json_(whoami_(user));
    if (req.action === 'overview')     return json_(overview_(user));
    if (req.action === 'overview_low') return json_(overviewLow_(user));
    if (req.action === 'report_data')  return json_(reportData_(user, req));
    if (req.action === 'reg_opts')     return json_(regOpts_(user));
    if (req.action === 'reg_find')     return json_(regFind_(user, req.q));
    if (req.action === 'reg_plan')     return json_(regPlan_(user, req));
    if (req.action === 'reg_days')     return json_(regDays_(user, req));
    if (req.action === 'tt_all')       return json_(ttAll_(user));
    if (req.action === 'contacts')     return json_(contacts_(user));
    if (req.action === 'report_save')  return json_(saveReports_(user, req.records || []));
if (req.action === 'reports_mine') return json_(getMyReports_(user));
if (req.action === 'report_delete') return json_(deleteReport_(user, req.id));
    if (req.action === 'guest_search') return json_(guestSearch_(user, req.sheet, req.q));
    if (req.action === 'guest_add')    return json_(guestAdd_(user, req.sheet, req.roll));
    if (req.action === 'guest_remove') return json_(guestRemove_(user, req.sheet, req.roll));
    if (req.action === 'sheet')         return json_(getSheetData_(user, req.sheet, req.create === true));
    if (req.action === 'save')          return json_(saveRecords_(user, req.records || []));
    if (req.action === 'delete_day')    return json_(deleteDay_(user, req.sheet, req.day, req.confirmDay, req.period));
    if (req.action === 'pdf')           return json_(exportPdf_(user, req.sheet));
    if (req.action === 'password')      return json_(changePassword_(user, req.oldPass, req.newPass));
    if (req.action === 'remarks_get')   return json_(getRemarks_(user));
    if (req.action === 'remarks_save')  return json_(saveRemarks_(user, req.records || []));
    return json_({ ok: false, error: 'UNKNOWN_ACTION' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
function deleteReport_(user, id) {
  id = String(id || '').trim();
  if (!id) return { ok: false, error: 'Missing report id.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = reportsSheet_(), n = sh.getLastRow() - 1;
    if (n < 1) return { ok: false, error: 'REPORT_NOT_FOUND' };
    var v = sh.getRange(2, 1, n, REPORT_HEAD.length).getValues();
    var me = String(user.username).trim().toLowerCase(), admin = isAdmin_(user);
    for (var i = 0; i < n; i++) {
      if (String(v[i][0]).trim() !== id) continue;
      if (!admin && String(v[i][3]).trim().toLowerCase() !== me) return { ok: false, error: 'This is not your report.' };
      var st = String(v[i][18] || 'Received').trim();
      if (!admin && st !== 'Received') return { ok: false, error: 'The Chief Proctor already acted on this report (' + st + '). It cannot be deleted.' };
      sh.deleteRow(i + 2);
      SpreadsheetApp.flush();
      var chk = sh.getLastRow() - 1, still = false;
      if (chk > 0) sh.getRange(2, 1, chk, 1).getValues().forEach(function (r) { if (String(r[0]).trim() === id) still = true; });
      return still ? { ok: false, error: 'Row could not be removed. Try again.' } : { ok: true, id: id, deleted: true };
    }
    return { ok: false, error: 'REPORT_NOT_FOUND' };
  } finally { lock.releaseLock(); }
}
function getProctors_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users');
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  var n = function (s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); };
  var hdr = v[0].map(n);
  var col = function (name, fallback) { var i = hdr.indexOf(name); return i === -1 ? fallback : i; };
  var iName = col('teacher', 0), iDes = col('designation', 2), iDept = col('department', 3),
      iPh = col('contact', 6), iRole = col('role', 8), iCP = col('chiefproctor', 9);
  var title = function (s) { return String(s).toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); }); };
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var tag = String(v[i][iCP]).trim().toUpperCase();
    if (!/^CP\d*$/.test(tag)) continue;                       // accepts CP, CP1, CP2 ...
    var role = String(v[i][iRole]).trim().toUpperCase();
    if (role === 'PRINCIPAL' || role === 'ADMIN') continue;   // not real proctors
    var phone = String(v[i][iPh]).replace(/\D/g, '');
    if (!phone) continue;
    out.push({
      id: 'CP' + (out.length + 1),                            // CP1, CP2 ... by row order
      name: title(v[i][iName]),
      dept: title(v[i][iDept]),
      desig: title(v[i][iDes]),
      phone: phone                                            // +923158813443 -> 923158813443
    });
  }
  return out;
}

function reportsSheet_() {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(REPORTS_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(REPORTS_SHEET);
  sh.getRange(1, 1, 1, REPORT_HEAD.length).setValues([REPORT_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.setFrozenRows(1);
  sh.getRange(2, 6, 1000, 1).setNumberFormat('@');
  sh.getRange(2, 10, 1000, 2).setNumberFormat('@');   // incident date + lecture stay text
  sh.setColumnWidth(14, 400);
  return sh;
}

function saveReports_(user, records) {
  var okNames = allowed_(user), saved = [], failed = [], results = {};
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = reportsSheet_(), n = sh.getLastRow() - 1, have = {};
    if (n > 0) sh.getRange(2, 1, n, 1).getValues().forEach(function (r) { have[String(r[0])] = true; });
    var cps = getProctors_(), rows = [];
    records.forEach(function (rec) {
      try {
        if (!rec.id) throw new Error('Missing report id');
        if (have[rec.id]) { saved.push(rec.id); results[rec.id] = {}; return; }   // already stored: safe to re-send
        var meta = okNames[rec.sheet];
        if (!meta) throw new Error('Not allowed: ' + rec.sheet);
        var txt = String(rec.txt || '').trim();
        if (txt.length < 10) throw new Error('Statement is too short.');
        var cp = cps.filter(function (c) { return c.id === rec.cp; })[0];
        rows.push([rec.id, new Date(), user.teacher, user.username, String(rec.who || ''), String(rec.roll || ''), meta.cls, meta.sub, meta.mon,
                   String(rec.date || ''), String(rec.period || ''), String(rec.cat || ''), rec.sev === 'Serious' ? 'Serious' : 'Normal',
                   txt.slice(0, 1000), String(rec.att || ''), cp ? cp.name : '', String(rec.via || ''), rec.guest ? (rec.home || 'Guest') : '', 'Received', '']);
        have[rec.id] = true; saved.push(rec.id);
        if (rec.removeGuest) {
          try { removeGuestCore_(meta, String(rec.roll || '').trim()); results[rec.id] = { guestRemoved: true }; }
          catch (ge) { results[rec.id] = { guestRemoved: false, guestError: String(ge.message || ge) }; }
        } else results[rec.id] = {};
      } catch (err) { failed.push({ id: rec.id, error: String(err.message || err) }); }
    });
    if (rows.length) {
      var start = sh.getLastRow() + 1;
      sh.getRange(start, 1, rows.length, 1).setNumberFormat('@');sh.getRange(start, 10, rows.length, 2).setNumberFormat('@');
      sh.getRange(start, 6, rows.length, 1).setNumberFormat('@');sh.getRange(start, 10, rows.length, 2).setNumberFormat('@');
      sh.getRange(start, 1, rows.length, REPORT_HEAD.length).setValues(rows);
    }
    SpreadsheetApp.flush();
    return { ok: true, saved: saved, failed: failed, results: results };
  } finally { lock.releaseLock(); }
}

function getMyReports_(user) {
  var sh = reportsSheet_(), n = sh.getLastRow() - 1, out = [];
  if (n < 1) return { ok: true, reports: out };
  var me = String(user.username).trim().toLowerCase();
  sh.getRange(2, 1, n, REPORT_HEAD.length).getValues().forEach(function (r) {
    if (String(r[3]).trim().toLowerCase() !== me) return;
    out.push({ id: r[0], at: Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'd MMM yyyy HH:mm'), teacher: r[2], who: r[4], roll: r[5],
      cls: r[6], sub: r[7], mon: r[8], date: (r[9] instanceof Date) ? Utilities.formatDate(r[9], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[9]),
lecture: r[10], cat: r[11], sev: r[12], txt: r[13], att: r[14],
      sent: r[15], guest: !!r[17], home: r[17], status: r[18] || 'Received' });
  });
  out.reverse();
  return { ok: true, reports: out };
}
/* ---------- small helpers ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function norm_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, '').toUpperCase();
}

function normalizeName(name) {
  if (!name) return '';
  return name.toString().toLowerCase()
    .replace(/\b(dr|ms|mr|mrs)\b\.?/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function isAdmin_(user) {
  return String(user.username).trim().toLowerCase() === ADMIN_USER;
}

/** True when a Roll No cell note marks the student as disabled (left the subject). */
function isOffNote_(n) { return String(n || '').indexOf(OFF_TAG) === 0; }

function academicStartYear_() {
  var d = new Date(), y = d.getFullYear(), m = d.getMonth();
  return (m > 7 || (m === 7 && d.getDate() >= 15)) ? y : y - 1;
}

function monthYear_(tok) {
  var m = SEASON_MONTH[tok];
  var start = academicStartYear_();
  return { m: m, y: m >= 7 ? start : start + 1 };
}

/** Months shown to teachers: from SEPT up to the current month (all of them outside the season). */
function activeMonths_() {
  var m = new Date().getMonth();
  var idx = m >= 8 ? m - 8 : (m === 0 ? 4 : -1);
  return idx < 0 ? SEASON.slice() : SEASON.slice(0, idx + 1);
}

/** Season token of the current calendar month, or null outside Sept-Jan. */
function currentSeasonToken_() {
  var map = { 8: 'SEPT', 9: 'OCT', 10: 'NOV', 11: 'DEC', 0: 'JAN' };
  return map[new Date().getMonth()] || null;
}

/* ---------- one workbook per month ---------- */

var BOOK_CACHE_ = {};

/** Spreadsheet id of the workbook that holds the sheets of this month token (SEPT, OCT, NOV, DEC, JAN). */
function bookIdFor_(tok) {
  var id = BOOK_IDS[String(tok || '').toUpperCase()];
  if (!id) throw new Error('No workbook is set for month "' + tok + '". Check BOOK_IDS.');
  return id;
}

/** Opens the workbook of that month (opened once per request). */
function bookFor_(tok) {
  var id = bookIdFor_(tok);
  if (!BOOK_CACHE_[id]) BOOK_CACHE_[id] = SpreadsheetApp.openById(id);
  return BOOK_CACHE_[id];
}

/** All five month workbooks, in season order: [{ tok, book }] */
function allBooks_() {
  return SEASON.map(function (tok) { return { tok: tok, book: bookFor_(tok) }; });
}

/* ---------- cache helpers (values over 100KB are chunked) ---------- */

function cachePutBig_(key, obj) {
  var c = CacheService.getScriptCache();
  var s = JSON.stringify(obj), size = 30000, n = Math.max(1, Math.ceil(s.length / size)), o = {};
  for (var i = 0; i < n; i++) o[key + '_' + i] = s.substr(i * size, size);
  o[key + '_n'] = String(n);
  try { c.putAll(o, CACHE_TTL); } catch (e) {}
}

function cacheGetBig_(key) {
  var c = CacheService.getScriptCache();
  var n = Number(c.get(key + '_n'));
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var m = c.getAll(keys), s = '';
  for (var j = 0; j < keys.length; j++) {
    if (m[keys[j]] == null) return null;
    s += m[keys[j]];
  }
  try { return JSON.parse(s); } catch (e) { return null; }
}

function cacheClear_(k) {
  var c = CacheService.getScriptCache();
  var n = Number(c.get(k + '_n')) || 0, keys = [k + '_n'];
  for (var i = 0; i < n; i++) keys.push(k + '_' + i);
  c.removeAll(keys);
}

/** Run this by hand after you change the Timetable, Users or the strength sheet. */
function clearAppCaches() {
  CACHE_KEYS.concat(['master_att', 'users_att', 'tt_att', 'tt_att2']).forEach(cacheClear_);
    SEASON.forEach(function (t) { cacheClear_('ll2_' + t); });
}

/* ---------- login ---------- */

/** Column number of a Users header (lower-case, letters/digits only), or the fallback when the header is missing. */
function usersCol_(hdr, name, fallback) {
  var h = hdr.map(function (s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); });
  var i = h.indexOf(name);
  return i === -1 ? fallback : i;
}

function getUsers_() {
  var rows = cacheGetBig_('users_att2');
  if (rows) return rows;
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users');
  if (!sh) throw new Error('Users sheet not found');
  var d = sh.getDataRange().getValues();
  var iRole = usersCol_(d[0], 'role', 8), iDept = usersCol_(d[0], 'department', 3);
  rows = [];
  for (var i = 1; i < d.length; i++) {
    if (!d[i][4]) continue;
    // teacher, username, password, role, department
    rows.push([String(d[i][0]).trim(), String(d[i][4]).trim(), String(d[i][5]),
               String(d[i][iRole] || '').trim(), String(d[i][iDept] || '').trim()]);
  }
  cachePutBig_('users_att2', rows);
  return rows;
}

function authenticate_(username, password) {
  var u = String(username || '').trim().toLowerCase();
  if (!u) return null;
  var rows = getUsers_();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][1].toLowerCase() === u && rows[i][2] === String(password)) {
      return { teacher: rows[i][0], username: rows[i][1], password: rows[i][2], role: rows[i][3] || '', dept: rows[i][4] || '' };
    }
  }
  return null;
}

function changePassword_(user, oldPass, newPass) {
  if (String(oldPass) !== user.password) return { ok: false, error: 'WRONG_OLD' };
  if (!newPass || String(newPass).length < 4) return { ok: false, error: 'TOO_SHORT' };
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users');
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][4]).trim().toLowerCase() === user.username.toLowerCase()) {
      sh.getRange(i + 1, 6).setValue(String(newPass));
      cacheClear_('users_att2');
      return { ok: true };
    }
  }
  return { ok: false, error: 'User not found' };
}

/* ---------- timetable -> this teacher's sections ---------- */

function getTimetable_() {
  var t = cacheGetBig_('tt_att3');
  if (t) return t;
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Timetable');
  if (!sh) throw new Error('Timetable sheet not found');
  var data = sh.getDataRange().getValues();
  t = { has: {}, rows: [] };
  if (data.length >= 2) {
    var h = data[0];
    // row = [teacher, program, class, pcs, subject, period, theory days, practical days, days, room, course name, section]
    var names = ['Teacher', 'Program', 'Class', 'PCS', 'Subject', 'Period', 'Theory', 'Practical', 'Days', 'Room', 'Course Name', 'Section'];
    var hl = h.map(function (x) { return String(x).trim().toLowerCase(); });
    var ix = names.map(function (n) { var i = h.indexOf(n); if (i === -1) i = hl.indexOf(n.toLowerCase()); t.has[n] = i !== -1; return i; });
    for (var i = 1; i < data.length; i++) {
      t.rows.push(ix.map(function (c) { return c === -1 ? '' : String(data[i][c] || '').trim(); }));
    }
  }
  cachePutBig_('tt_att3', t);
  return t;
}

/** [{cls:'FCS-1 D2', subject:'COMPUTER SCIENCE', teacher:'Asima Latif'}] - only this teacher's rows (admin: all). */
function getAssignments_(user) {
  var t = getTimetable_();
  ['PCS', 'Subject', 'Teacher'].forEach(function (n) {
    if (t.rows.length && !t.has[n]) throw new Error('Timetable must contain a "' + n + '" column.');
  });
  var admin = isAdmin_(user), me = normalizeName(user.teacher), seen = {}, out = [];
  t.rows.forEach(function (r) {
    var teacher = r[0], program = r[1], pcs = r[3], subject = r[4];
    if (!pcs || !subject) return;
    if (PROGRAM_FILTER && program.toUpperCase() !== PROGRAM_FILTER.toUpperCase()) return;
    if (!admin && normalizeName(teacher) !== me) return;
    var key = pcs.toUpperCase() + '|' + subject.toUpperCase();
    if (seen[key]) return;
    seen[key] = true;
    out.push({ cls: pcs, subject: subject.toUpperCase(), teacher: teacher });
  });
  return out;
}

/** "F.A-1 Optional Subjects" -> "F.A-1" (other classes unchanged) */
function baseClass_(cls) {
  return String(cls || '').replace(/\s*Optional\s*Subjects?/i, '').replace(/\s+/g, ' ').trim();
}

/** Section letter from a student's PCS: "FCS-1 D2" -> "D2", "F.A-1" -> "" */
function sectionOf_(classSec) {
  var p = String(classSec || '').trim().split(/\s+/);
  return p.length > 1 ? p[p.length - 1] : '';
}

function makeName_(cls, mon, subject) {
  return (baseClass_(cls) + ' ' + mon + ' ' + subject)
    .replace(/[\[\]\*\?:\/\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 100);
}

function listSheets_(user) {
  var months = activeMonths_(), out = [];
  getAssignments_(user).forEach(function (a) {
    months.forEach(function (mon) {
      out.push({ name: makeName_(a.cls, mon, a.subject), cls: a.cls, sub: a.subject, mon: mon, teacher: a.teacher });
    });
  });
  out.sort(function (x, y) {
    return x.cls.localeCompare(y.cls, undefined, { numeric: true }) ||
           x.sub.localeCompare(y.sub) ||
           SEASON.indexOf(x.mon) - SEASON.indexOf(y.mon);
  });
  return out;
}

/** name -> meta, only for sections this user may touch. */
function allowed_(user) {
  var map = {};
  listSheets_(user).forEach(function (s) { map[s.name] = s; });
  return map;
}

/* ---------- student roster for a new sheet ---------- */

function getMaster_() {
  var M = cacheGetBig_('master_att3');
  if (M) return M;
  M = { hasClass: false, rows: [] };      // rows: [enroll, name, classSec, [SUBJECTS...], fatherName]
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Transform-ALL-COLLEGE-STRENGTH');
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getDataRange().getValues();
    var h = v[0].map(norm_);
    var find = function (list) { for (var i = 0; i < list.length; i++) { var x = h.indexOf(list[i]); if (x !== -1) return x; } return -1; };
    var ec = h.indexOf('ENROLL#');
    var nc = find(['NAME', 'STUDENTNAME', 'STUDENT', 'FULLNAME']);
    var cc = find(['CLASS-SEC', 'CLASSSEC', 'CLASS-SECTION', 'PCS', 'SECTION']);
    var fc = find(['FATHERNAME', 'FATHER', "FATHER'SNAME", 'FATHER/GUARDIAN', 'GUARDIANNAME']);
    var sc = ['SUB1', 'SUB2', 'SUB3', 'SUB4', 'SUB5', 'SUB6', 'SUB7'].map(function (s) { return h.indexOf(s); })
      .filter(function (i) { return i !== -1; });
    M.hasClass = cc !== -1;
    if (ec !== -1) {
      for (var i = 1; i < v.length; i++) {
        var e = v[i][ec];
        if (!e) continue;
        var subs = [];
        sc.forEach(function (c) { var s = String(v[i][c] || '').trim().toUpperCase(); if (s) subs.push(s); });
        M.rows.push([String(e).trim(), nc === -1 ? String(e).trim() : String(v[i][nc] || '').trim(),
                     cc === -1 ? '' : String(v[i][cc] || '').trim(), subs,
                     fc === -1 ? '' : String(v[i][fc] || '').trim()]);
      }
    }
  }
  cachePutBig_('master_att3', M);
  return M;
}

function enrollNum_(e) {
  var p = String(e || '').split('-');
  return parseInt(p[p.length - 1], 10) || 0;
}

/** Returns [[roll, name, father, classSec], ...] sorted by roll number. */
function getRoster_(book, meta) {
  var M = getMaster_();
  var subj = meta.sub.toUpperCase();
  var squash = function (s) { return String(s || '').replace(/\s+/g, '').toUpperCase(); };
  var base = baseClass_(meta.cls);
  var baseSq = squash(base);
  var isOpt = baseSq !== squash(meta.cls);
  var classOf = {};
  M.rows.forEach(function (r) { classOf[r[0]] = r[2] || ''; });
  var list = [];

  // 1) Master sheet: students of this class who take this subject
  if (M.hasClass) {
    M.rows.forEach(function (r) {
      var c = squash(r[2]);
      var classOk = isOpt ? (c === baseSq || c.indexOf(baseSq) === 0) : (c === baseSq);
      if (classOk && r[3].indexOf(subj) !== -1) list.push([r[0], r[1], r[4] || '', r[2] || base]);
    });
  }

  // 2) Fallback: roster of another attendance sheet of the same class, filtered by subject
  var fromSibs = function (sibs) {
    for (var k = 0; k < sibs.length && !list.length; k++) {
      try {
        var L = readLayout_(sibs[k]);
        var n = sibs[k].getLastRow() - FIRST_ROW + 1;
        if (n < 1) continue;
        var rolls = sibs[k].getRange(FIRST_ROW, L.rollCol, n, 1).getValues();
        var names = L.nameCol > 0 ? sibs[k].getRange(FIRST_ROW, L.nameCol, n, 1).getValues() : null;
        var taking = {}, any = false;
        M.rows.forEach(function (r) { if (r[3].indexOf(subj) !== -1) { taking[r[0]] = true; any = true; } });
        var fathers = {};
        M.rows.forEach(function (r) { fathers[r[0]] = r[4] || ''; });
        for (var i = 0; i < n; i++) {
          var roll = String(rolls[i][0]).trim();
          if (!roll) continue;
          if (any && !taking[roll]) continue;
          list.push([roll, names ? String(names[i][0]).trim() : roll, fathers[roll] || '', classOf[roll] || base]);
        }
      } catch (e) { /* try next sheet */ }
    }
  };

  // 2a) fast path: the same class + subject in the other months. Each month lives in its own workbook,
  //     so every month's workbook is checked with a direct lookup (no scan of all tabs).
  if (!list.length) {
    var fast = [];
    SEASON.forEach(function (tok) {
      var nm = makeName_(meta.cls, tok, meta.sub);
      if (nm === meta.name) return;
      try {
        var s = bookFor_(tok).getSheetByName(nm);
        if (s) fast.push(s);
      } catch (e) { /* month workbook not reachable, skip it */ }
    });
    fromSibs(fast);
  }

  // 2b) slow path: any sheet of the same class, in this month's workbook first, then the other months
  if (!list.length) {
    var sibs = [], seenIds = {};
    var scan = function (bk) {
      var id = bk.getId();
      if (seenIds[id]) return;
      seenIds[id] = true;
      bk.getSheets().forEach(function (s) {
        if (s.getName().toUpperCase().indexOf(base.toUpperCase() + ' ') === 0 && s.getName() !== meta.name) sibs.push(s);
      });
    };
    scan(book);
    SEASON.forEach(function (tok) { try { scan(bookFor_(tok)); } catch (e) {} });
    fromSibs(sibs);
  }

  list.sort(function (a, b) { return enrollNum_(a[0]) - enrollNum_(b[0]); });
  return list;
}

function titleCase_(s) {
  return String(s).toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}

/** Writes "Teacher: <name>" in A2 (only if A2 is empty or already a teacher stamp). */
function stampTeacher_(sheet, meta) {
  if (!meta || !meta.teacher) return;
  var cell = sheet.getRange(2, 1);
  var v = String(cell.getValue() || '');
  var want = 'Teacher: ' + meta.teacher;
  if (v === want) return;
  if (!v || /^Teacher:/i.test(v)) cell.setValue(want).setFontWeight('bold').setHorizontalAlignment('left');
}

/** Row 2, above each date: Mon, Tue, Wed... (blank for days that don't exist in the month). */
function setWeekdayRow_(sh, month, year, firstDayCol) {
  var col = firstDayCol || DAY_COL1;
  var names = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var dim = new Date(year, month + 1, 0).getDate(), vals = [];
  for (var d = 1; d <= 31; d++) vals.push(d > dim ? '' : names[new Date(year, month, d).getDay()]);
  sh.getRange(2, col, 1, 31).setValues([vals])
    .setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
}

/**
 * Removes the unused empty rows and columns of a sheet. A new Google sheet is 1,000 x 26 cells and every
 * cell (even empty) counts toward the 10,000,000 cell limit of the workbook.
 * usedRows = last row that holds data, usedCols = last column that holds data (at least the M.T column).
 */
function trimSheet_(sh, usedRows, usedCols) {
  var keepRows = Math.max(usedRows, FIRST_ROW) + SPARE_ROWS, maxR = sh.getMaxRows();
  if (maxR > keepRows) sh.deleteRows(keepRows + 1, maxR - keepRows);
  var keepCols = Math.max(usedCols || MT_COL, MT_COL), maxC = sh.getMaxColumns();
  if (maxC > keepCols) sh.deleteColumns(keepCols + 1, maxC - keepCols);
}

/* ---------- period-wise attendance: formulas, colours, hidden "Period Data" sheet ---------- */

/**
 * T.A of one row = number of lectures the student attended.
 * Cells look like P (1), 2P, 3P ... -> the number in front of the P (1 when there is no number).
 * A, L, B, 2A, 3L ... count 0.
 */
function taFormula_(row) {
  var r = colLetter_(DAY_COL1) + row + ':' + colLetter_(LAST_DAY) + row;
  return '=SUMPRODUCT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(' + r + '&"","^\\d+")),1)*(RIGHT(' + r + '&"",1)="P")))';
}

/** Colour rules of an attendance sheet: A red, L amber, B blue (cell CONTAINS the letter, so 3A, 2P+... work);
 *  low attendance (<73%) and low monthly test (<50%) highlighted. n = number of student rows. */
function applyRules_(sh, n) {
  var days = sh.getRange(FIRST_ROW, DAY_COL1, n, 31);
  var taL = colLetter_(TA_COL), mtL = colLetter_(MT_COL);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('A')
      .setBackground('#FF0000').setFontColor('#FFFFFF').setBold(true).setRanges([days]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('L')
      .setBackground('#FFE599').setFontColor('#7F6000').setBold(true).setRanges([days]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('B')
      .setBackground('#9FC5E8').setFontColor('#0B3D6B').setBold(true).setRanges([days]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + taL + '$2>0,$' + taL + FIRST_ROW + '/$' + taL + '$2<0.73)')
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([sh.getRange(FIRST_ROW, 4, n, 1), sh.getRange(FIRST_ROW, TA_COL, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($' + mtL + FIRST_ROW + '),$' + mtL + '$2>0,$' + mtL + FIRST_ROW + '/$' + mtL + '$2<0.5)')
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([sh.getRange(FIRST_ROW, MT_COL, n, 1)]).build()
  ]);
}

/** The hidden sheet that keeps the real marks of every period (one row per attendance sheet). */
function pdSheet_(book) {
  var sh = book.getSheetByName(PD_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(PD_SHEET);
  sh.getRange(1, 1, 1, 3).setValues([['Sheet name', 'Period data (do not edit)', 'Updated']]).setFontWeight('bold');
  sh.setColumnWidth(1, 320);
  sh.hideSheet();
  return sh;
}

/** Loads the period marks of one attendance sheet: { book, sh, row, d: { day: { period: { roll: mark } } } } */
function pdLoad_(book, name) {
  var pd = { book: book, sh: book.getSheetByName(PD_SHEET), row: 0, d: {} };
  if (!pd.sh || pd.sh.getLastRow() < 2) return pd;
  var f = pd.sh.getRange(2, 1, pd.sh.getLastRow() - 1, 1).createTextFinder(name)
    .matchCase(true).matchEntireCell(true).findNext();
  if (!f) return pd;
  pd.row = f.getRow();
  var raw;
  try { raw = JSON.parse(pd.sh.getRange(pd.row, 2).getValue() || '{}'); } catch (e) { raw = {}; }
  var rolls = raw.r || [];
  Object.keys(raw.d || {}).forEach(function (day) {
    pd.d[day] = {};
    Object.keys(raw.d[day]).forEach(function (p) {
      var s = String(raw.d[day][p]), m = {};
      for (var i = 0; i < rolls.length; i++) { var c = s.charAt(i); if (VALID[c]) m[rolls[i]] = c; }
      pd.d[day][p] = m;
    });
  });
  return pd;
}

/** Saves the period marks (compact: one string per period, one character per student of the stored roll list). */
function pdSave_(pd, name) {
  var rolls = [], seen = {};
  Object.keys(pd.d).forEach(function (day) {
    Object.keys(pd.d[day]).forEach(function (p) {
      Object.keys(pd.d[day][p]).forEach(function (r) { if (!seen[r]) { seen[r] = 1; rolls.push(r); } });
    });
  });
  var d = {};
  Object.keys(pd.d).forEach(function (day) {
    d[day] = {};
    Object.keys(pd.d[day]).forEach(function (p) {
      d[day][p] = rolls.map(function (r) { return pd.d[day][p][r] || '-'; }).join('');
    });
  });
  var json = JSON.stringify({ r: rolls, d: d });
  if (json.length > 45000) throw new Error('Period data of ' + name + ' is too large.');
  if (!pd.sh) pd.sh = pdSheet_(pd.book);
  if (!pd.row) pd.row = pd.sh.getLastRow() + 1;
  pd.sh.getRange(pd.row, 1, 1, 3).setValues([[name, json, new Date()]]);
}

/** Old single-letter cells of a day -> { 1: { roll: mark } } (old days become "period 1"). */
function legacyDay_(I, col) {
  var vals = I.sheet.getRange(FIRST_ROW, col, I.n, 1).getValues(), m = {}, any = false;
  for (var i = 0; i < I.n; i++) {
    var v = String(vals[i][0]).trim().toUpperCase();
    if (I.rolls[i] && VALID[v]) { m[I.rolls[i]] = v; any = true; }
  }
  return any ? { 1: m } : {};
}

/** Set of rolls that have a mark in any period of this day. */
function rollsOf_(dayPm) {
  var s = {};
  Object.keys(dayPm || {}).forEach(function (p) {
    Object.keys(dayPm[p]).forEach(function (r) { s[r] = true; });
  });
  return s;
}

/** The text of the day cell: P, 2P, 3P ... or (no P) A / 3A, L / 3L, B / 3B. */
function cellFor_(dayPm, roll) {
  var ps = Object.keys(dayPm || {}), held = ps.length, p = 0, seen = {};
  ps.forEach(function (k) {
    var m = dayPm[k][roll];
    if (!m) return;
    if (m === 'P') p++; else seen[m] = true;
  });
  if (p) return p === 1 ? 'P' : p + 'P';
  for (var i = 0; i < NONP.length; i++) {
    if (seen[NONP[i]]) return held === 1 ? NONP[i] : held + NONP[i];
  }
  return '';
}

/**
 * Rewrites the day column (only students that have, or had, period data; disabled students are never written)
 * and the "periods held" number in row 1, then reads the cells back to verify.
 * rollSet = rolls that had data BEFORE this change (so removed marks are cleared too).
 */
function writeDay_(I, pd, day, col, rollSet) {
  var pm = pd.d[String(day)] || {}, held = Object.keys(pm).length;
  var now = rollsOf_(pm);
  var rng = I.sheet.getRange(FIRST_ROW, col, I.n, 1), vals = rng.getValues(), want = {};
  for (var i = 0; i < I.n; i++) {
    var roll = I.rolls[i];
    if (!roll || I.off[i] || !(rollSet[roll] || now[roll])) continue;
    var v = cellFor_(pm, roll);
    vals[i][0] = v; want[i] = v;
  }
  rng.setValues(vals);
  I.sheet.getRange(1, col).setValue(held || '').setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
  SpreadsheetApp.flush();
  var back = rng.getValues(), bad = 0;
  Object.keys(want).forEach(function (k) {
    if (String(back[k][0]).trim().toUpperCase() !== want[k]) bad++;
  });
  if (bad) throw new Error('Verification failed: ' + bad + ' cell(s) in ' + I.sheet.getName() + ' day ' + day + ' do not match. Tap Sync now to retry.');
}

/** Creates the monthly sheet (same layout every time) if it does not exist yet.
 *  `book` must be the workbook of meta.mon (use bookFor_(meta.mon)). */
function ensureSheet_(book, meta) {
  var sh = book.getSheetByName(meta.name);
  if (sh) return sh;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sh = book.getSheetByName(meta.name);
    if (sh) return sh;

    var roster = getRoster_(book, meta);
    if (!roster.length) {
      throw new Error('No students found for ' + meta.cls + ' / ' + meta.sub +
        '. Check the section and subject columns in Transform-ALL-COLLEGE-STRENGTH.');
    }

    var my = monthYear_(meta.mon);
        var gIdx = {};
    activeGuests_(meta.cls, meta.sub, meta.mon).forEach(function (g) {
      if (!roster.some(function (r) { return r[0] === g.roll; })) { gIdx[roster.length] = g; roster.push([g.roll, g.name, g.father, g.home]); }
    });
    var n = roster.length;
    var totalCols = MT_COL;

    sh = book.insertSheet(meta.name);

    // rows 1-2: grey banner
    sh.getRange(1, 1, 2, totalCols).setBackground('#CFCCCC');
    sh.getRange(1, 1).setValue(TITLE_PREFIX + titleCase_(MONTH_NAMES[my.m]) + ' ' + my.y)
      .setFontWeight('bold').setFontSize(11);
    stampTeacher_(sh, meta);
    sh.getRange(1, TA_COL).setValue('T.L').setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(2, TA_COL).setValue(0).setFontWeight('bold').setHorizontalAlignment('center');
    setWeekdayRow_(sh, my.m, my.y);

    // row 3: headers
    var headers = ['Sr. No', 'Roll No', 'Section', 'StudentName', 'FatherName', 'Class-Sec'];
    for (var d = 1; d <= 31; d++) headers.push(d);
    headers.push('T.A', 'M.T');
    sh.getRange(HEADER_ROW, 1, 1, totalCols).setValues([headers])
      .setFontWeight('bold').setBackground('#D9D9D9').setHorizontalAlignment('center').setVerticalAlignment('middle');

    // students
    sh.getRange(FIRST_ROW, 2, n, 1).setNumberFormat('@');
    var rows = roster.map(function (r, i) { return [i + 1, r[0], sectionOf_(r[3]), r[1], r[2] || '', r[3] || baseClass_(meta.cls)]; });
        sh.getRange(FIRST_ROW, 1, n, 6).setValues(rows);
        Object.keys(gIdx).forEach(function (i) { var g = gIdx[i]; sh.getRange(FIRST_ROW + Number(i), 2).setNote(guestNote_(g.home, g.by, g.on)); });

    // T.A = number of lectures attended: P = 1, 2P = 2, 3P = 3 ... (A, L, B are NOT counted)
    var ta = [];
    for (var r = 0; r < n; r++) ta.push([taFormula_(FIRST_ROW + r)]);
    sh.getRange(FIRST_ROW, TA_COL, n, 1).setFormulas(ta).setHorizontalAlignment('center');

    // alignment
    sh.getRange(FIRST_ROW, 1, n, 3).setHorizontalAlignment('center');
    sh.getRange(FIRST_ROW, 4, n, 1).setHorizontalAlignment('left');
    sh.getRange(FIRST_ROW, 6, n, 1).setHorizontalAlignment('center');
    var days = sh.getRange(FIRST_ROW, DAY_COL1, n, 31);
    days.setHorizontalAlignment('center');
    sh.getRange(FIRST_ROW, MT_COL, n, 1).setHorizontalAlignment('center');

    // day column shading: Sunday black, Saturday pink, days not in this month grey
    var dim = new Date(my.y, my.m + 1, 0).getDate();
    for (var dd = 1; dd <= 31; dd++) {
      var bg = null;
      if (dd > dim) bg = '#B7B7B7';
      else {
        var wd = new Date(my.y, my.m, dd).getDay();
        if (wd === 0) bg = '#000000'; else if (wd === 6) bg = '#EAD1DC';
      }
      if (bg) sh.getRange(FIRST_ROW, DAY_COL1 + dd - 1, n, 1).setBackground(bg);
    }

    // A = red, L = amber, B = blue; low attendance (<73%) and low monthly test (<50%) highlighted
    applyRules_(sh, n);

    // borders, filter, freeze, widths
    sh.getRange(HEADER_ROW, 1, n + 1, totalCols)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(HEADER_ROW, 1, n + 1, totalCols).createFilter();
    sh.setFrozenRows(HEADER_ROW);
    sh.setFrozenColumns(4);
    sh.setColumnWidth(1, 50);
    sh.setColumnWidth(2, 115);
    sh.setColumnWidth(3, 65);
    sh.setColumnWidth(4, 190);
    sh.setColumnWidth(5, 150);
    sh.setColumnWidth(6, 100);
    sh.setColumnWidths(DAY_COL1, 31, 34);
    sh.setColumnWidth(TA_COL, 50);
    sh.setColumnWidth(MT_COL, 50);

    // keep the workbook small: drop the empty rows/columns a new sheet comes with
    try { trimSheet_(sh, FIRST_ROW + n - 1, totalCols); } catch (e) { Logger.log('Trim skipped for ' + meta.name + ': ' + e.message); }

    SpreadsheetApp.flush();
    return sh;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- subject / section changes: disable old row, add new admission ---------- */

/** Disables students who no longer take this subject/section, re-enables those who came back.
 *  Returns { off: n, on: n }. Only disables when allowOff is true (current month). */
function syncSubjectChanges_(book, sh, meta, allowOff) {
  var res = { off: 0, on: 0 };
  var M = getMaster_();
  if (!M.hasClass || !M.rows.length) return res;
  var roster = getRoster_(book, meta);
  if (!roster.length) return res;

  var inRoster = {}, inMaster = {};
  roster.forEach(function (r) { inRoster[String(r[0]).trim()] = true; });
  M.rows.forEach(function (r) { inMaster[String(r[0]).trim()] = true; });

  var L = readLayout_(sh);
  var total = sh.getLastRow() - FIRST_ROW + 1;
  if (total < 1) return res;
  var rg = sh.getRange(FIRST_ROW, L.rollCol, total, 1);
  var rolls = rg.getValues(), notes = rg.getNotes();

  var toOff = [], toOn = [], active = 0;
  for (var i = 0; i < total; i++) {
    var roll = String(rolls[i][0]).trim();
    if (!roll) continue;
    var off = isOffNote_(notes[i][0]);
    if (inRoster[roll]) { if (off) toOn.push(i); else active++; }
    else if (!off) { active++; if (inMaster[roll]) toOff.push(i); }   // still a student, but not in this subject
  }

  if (allowOff && toOff.length > 3 && toOff.length > active * 0.4) {
    Logger.log('SKIPPED disabling in ' + meta.name + ': ' + toOff.length + ' of ' + active +
      ' students would be disabled. Check the subject name / Transform sheet.');
    toOff = [];
  }
  if (!allowOff) toOff = [];

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy');
  toOff.forEach(function (i) {
    var row = FIRST_ROW + i;
    sh.getRange(row, L.rollCol).setNote(OFF_TAG + ' - changed subject' );
    sh.getRange(row, 1, 1, 6).setBackground('#EFEFEF').setFontColor('#999999');
    sh.getRange(row, 4).setFontLine('line-through');
    res.off++;
  });
  toOn.forEach(function (i) {
    var row = FIRST_ROW + i;
    sh.getRange(row, L.rollCol).clearNote();
    sh.getRange(row, 1, 1, 6).setBackground(null).setFontColor('#000000');
    sh.getRange(row, 4).setFontLine('none');
    res.on++;
  });
  if (res.off || res.on) SpreadsheetApp.flush();
  return res;
}



/* ---------- read / save ---------- */

/** Finds Roll No / Name / M.T columns and the column of each day (1..31) from the header row. */
function readLayout_(sheet) {
  var lastCol = sheet.getLastColumn();
  var hdr = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  var L = { rollCol: -1, nameCol: -1, mtCol: -1, dayCols: {} };
  for (var c = 0; c < hdr.length; c++) {
    var t = norm_(hdr[c]);
    if (t === 'ROLLNO' && L.rollCol < 0) L.rollCol = c + 1;
    else if (t === 'STUDENTNAME' && L.nameCol < 0) L.nameCol = c + 1;
    else if ((t === 'M.T' || t === 'MT') && L.mtCol < 0) L.mtCol = c + 1;
    else if (t !== '' && /^\d+$/.test(t)) {
      var d = Number(t);
      if (d >= 1 && d <= 31 && !L.dayCols[d]) L.dayCols[d] = c + 1;
    }
  }
  if (L.rollCol < 0) throw new Error('Roll No header not found in ' + sheet.getName());
  return L;
}

function dayColRange_(L) {
  var cols = Object.keys(L.dayCols).map(function (d) { return L.dayCols[d]; });
  if (!cols.length) return null;
  return { c1: Math.min.apply(null, cols), c2: Math.max.apply(null, cols) };
}

/**
 * Returns the sheet for the app. marks has the shape marks[day][period][roll] = 'P'|'A'|'L'|'B'.
 * Days saved before period-wise attendance existed (single letters in the cells) are returned as period 1.
 */
function getSheetData_(user, name, create) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };

  var book = bookFor_(meta.mon);          // the workbook of the SELECTED month
  var sheet = book.getSheetByName(name);
  if (!sheet) {
    if (!create) return { ok: true, missing: true, sheet: name };
    sheet = ensureSheet_(book, meta);
  }

  var L = readLayout_(sheet);
  var n = sheet.getLastRow() - FIRST_ROW + 1;
  var students = [], marks = {}, notes = {}, rolls = [];
  var days = Object.keys(L.dayCols).map(Number).sort(function (a, b) { return a - b; });
  var rg = dayColRange_(L);
      var pd = pdLoad_(book, name);
  var guestAt = guestAtMap_(meta);

  if (n > 0) {
    var rollRange = sheet.getRange(FIRST_ROW, L.rollCol, n, 1);
    rolls = rollRange.getValues();
    var rollNotes = rollRange.getNotes();     // a LEFT SUBJECT note = disabled student
    var names = L.nameCol > 0 ? sheet.getRange(FIRST_ROW, L.nameCol, n, 1).getValues() : null;
    var rollList = [];
    for (var i = 0; i < n; i++) {
      var roll = String(rolls[i][0]).trim();
      if (!roll) continue;
      rollList.push(i);
          var st = { roll: roll, name: names ? String(names[i][0]).trim() : roll };
      var gnote = String(rollNotes[i][0] || '');
      if (isOffNote_(gnote)) { st.off = true; if (gnote.indexOf('guest removed') !== -1) st.offGuest = true; }
      if (gnote.indexOf('GUEST|') === 0) { var gp = gnote.split('|'); st.guest = { home: gp[1] || '', by: gp[2] || '', on: gp[3] || '' }; }
      if (guestAt[roll]) st.guestAt = guestAt[roll];
      students.push(st);
    }
    if (rg) {
      // ONE read for the whole month instead of one per day
      var block = sheet.getRange(FIRST_ROW, rg.c1, n, rg.c2 - rg.c1 + 1).getValues();
      days.forEach(function (d) {
        if (pd.d[d] && Object.keys(pd.d[d]).length) { marks[d] = pd.d[d]; return; }   // period-wise day
        var j = L.dayCols[d] - rg.c1, m = {}, any = false;                              // old day: single letters = period 1
        rollList.forEach(function (i) {
          var v = String(block[i][j]).trim().toUpperCase();
          if (VALID[v]) { m[String(rolls[i][0]).trim()] = v; any = true; }
        });
        if (any) marks[d] = { 1: m };
      });
    }
  }
  if (rg) {
    var nr = sheet.getRange(HEADER_ROW, rg.c1, 1, rg.c2 - rg.c1 + 1).getNotes()[0];
    days.forEach(function (d) { var t = nr[L.dayCols[d] - rg.c1]; if (t) notes[d] = t; });
  }

  var mt = { total: 0, marks: {} };
  if (L.mtCol > 0 && n > 0) {
    mt.total = Number(sheet.getRange(2, L.mtCol).getValue()) || 0;
    var mv = sheet.getRange(FIRST_ROW, L.mtCol, n, 1).getValues();
    for (var q = 0; q < n; q++) {
      var rr = String(rolls[q][0]).trim(), mval = String(mv[q][0]).trim();
      if (rr && mval !== '') mt.marks[rr] = mval;
    }
  }

  var my = monthYear_(meta.mon);
  var title = String(sheet.getRange(1, 1).getValue() || '').trim();
  var lv = {};
  try { lv = leavesForSheet_(meta, students); } catch (e) { Logger.log('Approved leaves skipped for ' + name + ': ' + e.message); }
  return { ok: true, sheet: name, title: title, m: my.m, y: my.y, days: days, students: students, marks: marks, notes: notes, mt: mt, lv: lv };
}
 

/**
 * Total lectures held in the month, written in row 2 after the last day column (T.L).
 * A day counts as many lectures as the number in row 1 above its date (periods held).
 * Old days without that number count as 1 lecture when somebody has P or A. L and B alone do not count.
 */
function updateLectures_(sheet, L, n) {
  var rg = dayColRange_(L);
  if (!rg || n < 1) return;
  var width = rg.c2 - rg.c1 + 1;
  var held = sheet.getRange(1, rg.c1, 1, width).getValues()[0];
  var block = sheet.getRange(FIRST_ROW, rg.c1, n, width).getValues();
  var rolls = sheet.getRange(FIRST_ROW, L.rollCol, n, 1).getValues();
  var count = 0;
  Object.keys(L.dayCols).forEach(function (d) {
    var j = L.dayCols[d] - rg.c1, h = Number(held[j]) || 0;
    if (!h) {
      for (var i = 0; i < n; i++) {
        if (!String(rolls[i][0]).trim()) continue;          // ignore rows with no student
        var v = String(block[i][j]).trim().toUpperCase();
        if (v === 'P' || v === 'A') { h = 1; break; }
      }
    }
    count += h;
  });
  sheet.getRange(2, rg.c2 + 1).setValue(count);
}

function recalcLecturesAllSheets() {
  batchSheets_('recalcLecturesAllSheets', function (sh) {
    var L = readLayout_(sh);
    var n = sh.getLastRow() - FIRST_ROW + 1;
    if (n < 1) return;
    updateLectures_(sh, L, n);
  });
}

/** Puts the note on the date header of that day (yellow header). Empty text clears it. */
function setDayNote_(sheet, L, day, text) {
  var col = L.dayCols[Number(day)];
  if (!col) throw new Error('Day ' + day + ' column not found');
  var cell = sheet.getRange(HEADER_ROW, col);
  text = String(text == null ? '' : text).trim();
  if (text) cell.setNote(text).setBackground('#FFD966');
  else cell.clearNote().setBackground('#D9D9D9');
}

/** Monthly test: total in row 2 of the M.T column, each student's mark (number or A) in that column.
 *  Disabled students (I.off) are never written. */
function saveMt_(I, rec) {
  if (!I.L.mtCol || I.L.mtCol < 0) throw new Error('M.T column not found in ' + rec.sheet);
  var total = Number(rec.total);
  if (!(total > 0)) throw new Error('Total marks must be more than 0.');
  var marks = rec.marks || {};
  var rng = I.sheet.getRange(FIRST_ROW, I.L.mtCol, I.n, 1), vals = rng.getValues(), want = {}, wrote = 0;
  for (var i = 0; i < I.n; i++) {
    var roll = I.rolls[i];
    if (!roll || I.off[i] || !marks.hasOwnProperty(roll)) continue;
    var s = String(marks[roll]).trim().toUpperCase(), v;
    if (s === '') v = '';
    else if (s === 'A' || s === 'AB' || s === 'ABSENT') v = 'A';
    else {
      v = Number(s);
      if (isNaN(v) || v < 0 || v > total) throw new Error('Marks for ' + roll + ' must be between 0 and ' + total + ' (or A).');
    }
    vals[i][0] = v; want[i] = v; wrote++;
  }
  rng.setValues(vals);
  I.sheet.getRange(1, I.L.mtCol).setValue('MT TOTAL').setFontSize(7).setFontWeight('bold').setHorizontalAlignment('center');
  I.sheet.getRange(2, I.L.mtCol).setValue(total).setFontWeight('bold').setHorizontalAlignment('center');
  SpreadsheetApp.flush();
  var back = I.sheet.getRange(FIRST_ROW, I.L.mtCol, I.n, 1).getValues(), bad = 0;
  Object.keys(want).forEach(function (k) {
    if (String(back[k][0]).trim().toUpperCase() !== String(want[k]).toUpperCase()) bad++;
  });
  if (bad) throw new Error('Verification failed: ' + bad + ' mark(s) in ' + rec.sheet + ' do not match. Tap Sync now to retry.');
  return { at: Date.now(), mt: true, wrote: wrote, total: total };
}

/**
 * records: [{ id, sheet, day, period, marks: { 'JCW-F26-2301': 'P'|'A'|'L'|'B'|'' }, note?: 'text' }]
 *   or monthly test: { id, kind:'mt', sheet, total, marks:{roll:'12'|'A'|''} }
 * period = 1..MAX_PERIOD (missing = 1). Only the marks of THAT period are replaced; the other periods of the
 * day stay. The day cell (P / 2P / 3A ...) is then rebuilt from all periods of the day.
 * Saving the same sheet/day/period again simply overwrites, so a re-sent record is harmless.
 * Every write is read back; a record is only reported as saved if the sheet really holds the values.
 * Each record is saved in the workbook of ITS month. Disabled students are never written.
 */
function saveRecords_(user, records) {
  var okNames = allowed_(user);
  var saved = [], failed = [], confirmed = {}, info = {}, touched = {}, createFail = {};

  // create missing sheets first, each in the workbook of its own month (ensureSheet_ takes the lock itself)
  records.forEach(function (rec) {
    if (!okNames[rec.sheet] || createFail[rec.sheet]) return;
    try {
      var bk = bookFor_(okNames[rec.sheet].mon);
      if (bk.getSheetByName(rec.sheet)) return;
      ensureSheet_(bk, okNames[rec.sheet]);
    }
    catch (err) { createFail[rec.sheet] = String(err.message || err); }
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    records.forEach(function (rec) {
      try {
        if (!okNames[rec.sheet]) throw new Error('Not allowed: ' + rec.sheet);
        if (createFail[rec.sheet]) throw new Error(createFail[rec.sheet]);
        var I = info[rec.sheet];
        if (!I) {
          var sheet = bookFor_(okNames[rec.sheet].mon).getSheetByName(rec.sheet);
          if (!sheet) throw new Error('Sheet not found: ' + rec.sheet);
          var L = readLayout_(sheet);
          var n = sheet.getLastRow() - FIRST_ROW + 1;
          if (n < 1) throw new Error('No students in ' + rec.sheet);
          var rollRange = sheet.getRange(FIRST_ROW, L.rollCol, n, 1);
          var rl = rollRange.getValues().map(function (r) { return String(r[0]).trim(); });
          var offs = rollRange.getNotes().map(function (r) { return isOffNote_(r[0]); });
          var idx = {};
          rl.forEach(function (r, i) { if (r) idx[r] = i; });
          I = info[rec.sheet] = { sheet: sheet, L: L, n: n, rolls: rl, off: offs, idx: idx, pd: null };
        }

        if (rec.kind === 'mt') {
          var cm = saveMt_(I, rec);
          touched[rec.sheet] = true; confirmed[rec.id] = cm; saved.push(rec.id);
          return;
        }

        var col = I.L.dayCols[Number(rec.day)];
        if (!col) throw new Error('Day ' + rec.day + ' column not found in ' + rec.sheet);

        var c = { at: Date.now() };
        var marks = rec.marks || {};
        var hasMarks = Object.keys(marks).length > 0;

        if (hasMarks) {
          var per = Number(rec.period) || 1;
          if (per < 1 || per > MAX_PERIOD) throw new Error('Period must be 1 to ' + MAX_PERIOD + '.');
          if (!I.pd) I.pd = pdLoad_(bookFor_(okNames[rec.sheet].mon), rec.sheet);
          var pd = I.pd, dk = String(Number(rec.day));
          if (!pd.d[dk]) pd.d[dk] = legacyDay_(I, col);          // old day with single letters = period 1
          var before = rollsOf_(pd.d[dk]);
          var pm = pd.d[dk][per] = pd.d[dk][per] || {};
          var wrote = 0;
          Object.keys(marks).forEach(function (roll) {
            if (I.idx[roll] == null || I.off[I.idx[roll]]) return;     // unknown or disabled student: never written
            var m = String(marks[roll]).trim().toUpperCase();
            if (VALID[m]) { pm[roll] = m; wrote++; }
            else if (m === '') delete pm[roll];
          });
          if (!Object.keys(pm).length) delete pd.d[dk][per];
          if (!Object.keys(pd.d[dk]).length) delete pd.d[dk];
          pdSave_(pd, rec.sheet);
          writeDay_(I, pd, rec.day, col, before);

          // counts of THIS period (active students only)
          c.p = 0; c.a = 0; c.l = 0; c.b = 0;
          var nowPm = (pd.d[dk] && pd.d[dk][per]) || {};
          Object.keys(nowPm).forEach(function (r) {
            if (I.idx[r] == null || I.off[I.idx[r]]) return;
            var v = nowPm[r];
            if (v === 'P') c.p++; else if (v === 'A') c.a++; else if (v === 'L') c.l++; else if (v === 'B') c.b++;
          });
          c.wrote = wrote;
          c.period = per;
        }

        if (rec.note !== undefined && rec.note !== null) {
          setDayNote_(I.sheet, I.L, rec.day, rec.note);
          var got = String(I.sheet.getRange(HEADER_ROW, col).getNote() || '').trim();
          if (got !== String(rec.note).trim()) throw new Error('Note could not be verified in ' + rec.sheet);
          c.note = true;
        }

        touched[rec.sheet] = true;
        confirmed[rec.id] = c;
        saved.push(rec.id);
      } catch (err) {
        failed.push({ id: rec.id, error: String(err.message || err) });
      }
    });

    Object.keys(touched).forEach(function (name) {
      try {
        var I = info[name];
        stampTeacher_(I.sheet, okNames[name]);
        updateLectures_(I.sheet, I.L, I.n);
      } catch (e) {}
      try { indexUpdate_(info[name].sheet, okNames[name]); } catch (e) { Logger.log('Index skipped for ' + name + ': ' + e.message); }
    });

    SpreadsheetApp.flush();
    return { ok: true, saved: saved, failed: failed, confirmed: confirmed };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- student remarks (tags + comments) ---------- */
/* Remarks are kept in ONE place: the SEPT workbook (BOOK_ID). */

function remarksSheet_() {
  var book = bookFor_('SEPT');
  var sh = book.getSheetByName(REMARKS_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(REMARKS_SHEET);
  sh.getRange(1, 1, 1, REMARK_HEAD.length).setValues([REMARK_HEAD])
    .setFontWeight('bold').setBackground('#D9D9D9');
  sh.getRange(2, 1, 1000, 4).setNumberFormat('@');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 140); sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 120); sh.setColumnWidth(5, 180); sh.setColumnWidth(6, 200);
  sh.setColumnWidth(7, 320); sh.setColumnWidth(8, 80);
  sh.hideColumns(1); sh.hideColumns(9); sh.hideColumns(10);
  return sh;
}

function remarkKey_(user, cls, roll) {
  return String(user.username).trim().toLowerCase() + '|' + cls + '|' + roll;
}

/** Everything this teacher has written: { 'FCS-1 D2|JCW-F26-2301': {tags:[], log:[{t,txt}], u:ms} } */
function getRemarks_(user) {
  var sh = remarksSheet_(), out = {};
  var n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true, remarks: out };
  var prefix = String(user.username).trim().toLowerCase() + '|';
  var v = sh.getRange(2, 1, n, REMARK_HEAD.length).getValues();
  v.forEach(function (r) {
    var key = String(r[0]);
    if (key.indexOf(prefix) !== 0) return;
    var log = [];
    try { log = JSON.parse(r[8] || '[]'); } catch (e) {}
    out[String(r[2]) + '|' + String(r[3])] = {
      tags: String(r[5] || '').split(',').map(function (t) { return t.trim(); }).filter(Boolean),
      log: log,
      u: Number(r[9]) || 0
    };
  });
  return { ok: true, remarks: out };
}

/**
 * records: [{ cls, roll, name, tags:[...], log:[{t,txt}], u }]
 * Newest edit wins (compared by u). An empty tags+log record clears that student's remarks.
 * Returns saved:[ 'cls|roll' ] and skipped:[ 'cls|roll' ] (server already had a newer edit).
 */
function saveRemarks_(user, records) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = remarksSheet_(), saved = [], skipped = [], appendRows = [];
    var n = sh.getLastRow() - 1, rowOf = {}, uOf = {};
    if (n > 0) {
      var keys = sh.getRange(2, 1, n, 1).getValues(), us = sh.getRange(2, 10, n, 1).getValues();
      for (var i = 0; i < n; i++) { rowOf[String(keys[i][0])] = i + 2; uOf[String(keys[i][0])] = Number(us[i][0]) || 0; }
    }
    records.forEach(function (rec) {
      var cls = String(rec.cls || '').trim(), roll = String(rec.roll || '').trim();
      if (!cls || !roll) return;
      var id = cls + '|' + roll, key = remarkKey_(user, cls, roll), u = Number(rec.u) || Date.now();
      if (rowOf[key] && uOf[key] > u) { skipped.push(id); return; }

      var tags = (Array.isArray(rec.tags) ? rec.tags : []).map(function (t) { return String(t).trim().slice(0, 40); })
        .filter(Boolean).slice(0, REMARK_MAX_TAGS);
      var log = (Array.isArray(rec.log) ? rec.log : []).map(function (e) {
        return { t: Number(e.t) || Date.now(), txt: String(e.txt || '').trim().slice(0, 400) };
      }).filter(function (e) { return e.txt; }).slice(-REMARK_MAX_LOG);
      var last = log.length ? log[log.length - 1] : null;
      var row = [key, user.teacher, cls, roll, String(rec.name || ''), tags.join(', '),
                 last ? Utilities.formatDate(new Date(last.t), Session.getScriptTimeZone(), 'd MMM yyyy') + ': ' + last.txt : '',
                 log.length, JSON.stringify(log), u,
                 Utilities.formatDate(new Date(u), Session.getScriptTimeZone(), 'd MMM yyyy HH:mm')];
      if (rowOf[key]) sh.getRange(rowOf[key], 1, 1, REMARK_HEAD.length).setValues([row]);
      else { appendRows.push(row); rowOf[key] = -1; }
      uOf[key] = u;
      saved.push(id);
    });
    if (appendRows.length) {
      var start = sh.getLastRow() + 1;
      sh.getRange(start, 1, appendRows.length, 4).setNumberFormat('@');
      sh.getRange(start, 1, appendRows.length, REMARK_HEAD.length).setValues(appendRows);
    }
    SpreadsheetApp.flush();
    return { ok: true, saved: saved, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- delete attendance of one period (or of the whole day) ---------- */

/**
 * period given  -> only that period of the day is removed (the other periods stay and the day cell is rebuilt).
 * period empty  -> the whole day is removed (all periods).
 * When no period is left on the day, the day column, the "periods held" number and the day note are cleared.
 */
function deleteDay_(user, name, day, confirmDay, period) {
  day = Number(day);
  if (!(day >= 1 && day <= 31) || Number(confirmDay) !== day) return { ok: false, error: 'Day confirmation does not match.' };
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  var book = bookFor_(meta.mon), sheet = book.getSheetByName(name);
  if (!sheet) return { ok: false, error: 'This sheet does not exist yet.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var L = readLayout_(sheet), col = L.dayCols[day];
    if (!col) throw new Error('Day ' + day + ' column not found in ' + name);
    var n = sheet.getLastRow() - FIRST_ROW + 1;
    if (n < 1) throw new Error('No students in ' + name);

    var rr = sheet.getRange(FIRST_ROW, L.rollCol, n, 1);
    var I = {
      sheet: sheet, L: L, n: n,
      rolls: rr.getValues().map(function (r) { return String(r[0]).trim(); }),
      off: rr.getNotes().map(function (r) { return isOffNote_(r[0]); })
    };

    var pd = pdLoad_(book, name), dk = String(day), cnt = 0;
    if (!pd.d[dk]) pd.d[dk] = legacyDay_(I, col);
    var before = rollsOf_(pd.d[dk]);
    period = Number(period) || 0;
    if (period) {
      cnt = Object.keys(pd.d[dk][period] || {}).length;
      delete pd.d[dk][period];
    } else {
      Object.keys(pd.d[dk]).forEach(function (p) { cnt += Object.keys(pd.d[dk][p]).length; });
      pd.d[dk] = {};
    }
    var empty = !Object.keys(pd.d[dk]).length;
    if (empty) delete pd.d[dk];
    pdSave_(pd, name);
    writeDay_(I, pd, day, col, before);
    if (empty) {
      sheet.getRange(FIRST_ROW, col, n, 1).clearContent();
      setDayNote_(sheet, L, day, '');
    }
    SpreadsheetApp.flush();
    updateLectures_(sheet, L, n);
    SpreadsheetApp.flush();
    try { indexUpdate_(sheet, meta); } catch (e) { Logger.log('Index skipped for ' + name + ': ' + e.message); }
    return { ok: true, count: cnt };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- delete one day's attendance (with backup) ---------- */
var DELLOG_SHEET = 'Deleted Attendance Log';

function logDeletedDay_(user, meta, day, deleted, note) {
  var book = bookFor_('SEPT');
  var sh = book.getSheetByName(DELLOG_SHEET);
  if (!sh) {
    sh = book.insertSheet(DELLOG_SHEET);
    sh.getRange(1, 1, 1, 8).setValues([['Deleted at', 'Teacher', 'Username', 'Sheet', 'Date', 'Marks deleted', 'Note', 'Data (roll: mark)']])
      .setFontWeight('bold').setBackground('#D9D9D9');
    sh.setFrozenRows(1);
  }
  var my = monthYear_(meta.mon);
  sh.appendRow([new Date(), user.teacher, user.username, meta.name,
                day + ' ' + MONTH_NAMES[my.m] + ' ' + my.y, Object.keys(deleted).length, note, JSON.stringify(deleted)]);
}

/* ---------- print: sheet as PDF ---------- */

function exportPdf_(user, name) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  var bookId = bookIdFor_(meta.mon);       // workbook of the selected month
  var book = bookFor_(meta.mon);
  var sh = book.getSheetByName(name);
  if (!sh) return { ok: false, error: 'This sheet does not exist yet.' };

  // hide notes from the printout: take them off, export, then put every one back
  var rng = sh.getDataRange();
  var notes = rng.getNotes();
  var hasNotes = notes.some(function (r) { return r.some(function (n) { return n; }); });
  if (hasNotes) { rng.clearNote(); }
  SpreadsheetApp.flush();

  var url = 'https://docs.google.com/spreadsheets/d/' + bookId + '/export?format=pdf' +
    '&gid=' + sh.getSheetId() +
    '&size=A4&portrait=false&fitw=true&gridlines=true&printtitle=false&sheetnames=false' +
    '&pagenum=UNDEFINED&shownotes=false&fzr=false&top_margin=0.4&bottom_margin=0.4&left_margin=0.3&right_margin=0.3';

  var resp;
  try {
    resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
  } finally {
    if (hasNotes) { rng.setNotes(notes); SpreadsheetApp.flush(); }   // notes are always restored
  }

  if (resp.getResponseCode() !== 200) {
    return { ok: false, error: 'PDF export failed (' + resp.getResponseCode() + ').' };
  }
  return { ok: true, name: name, pdf: Utilities.base64Encode(resp.getBlob().getBytes()) };
}

/* ---------- bulk jobs that must not hit the 6-minute limit ---------- */

/** Deletes pending one-time continuation triggers of a function (never touches the monthly trigger). */
function clearContinuation_(handler) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });
}

/**
 * Runs fn(sheet) on every attendance sheet of ALL five month workbooks. Stops after BATCH_MS, remembers
 * where it stopped and schedules itself again one minute later, until every sheet is done.
 * Returns true when finished.
 * `handler` is the name of the function that calls this (so the continuation runs the same job).
 * Tabs whose A1 does not start with TITLE_PREFIX (Period Data, Student Remarks ...) are skipped.
 */
function batchSheets_(handler, fn, toks) {
  var start = Date.now(), props = PropertiesService.getScriptProperties(), key = 'cursor_' + handler;
  clearContinuation_(handler);
  var sheets = [];
  var books = toks ? toks.map(function (t) { return { tok: t, book: bookFor_(t) }; }) : allBooks_();
  books.forEach(function (b) { b.book.getSheets().forEach(function (s) { sheets.push(s); }); });
  var i = Number(props.getProperty(key)) || 0, done = 0;
  for (; i < sheets.length; i++) {
    if (Date.now() - start > BATCH_MS) {
      props.setProperty(key, String(i));
      ScriptApp.newTrigger(handler).timeBased().after(60 * 1000).create();
      Logger.log(handler + ': ' + done + ' sheet(s) done this round, continuing automatically in 1 minute (at ' + i + ' of ' + sheets.length + ').');
      return false;
    }
    var sh = sheets[i];
    try {
      if (String(sh.getRange(1, 1).getValue()).indexOf(TITLE_PREFIX) !== 0) continue;
      fn(sh);
      done++;
    } catch (e) {
      Logger.log('Skipped ' + sh.getName() + ': ' + e.message);
    }
  }
  props.deleteProperty(key);
  Logger.log(handler + ': finished (' + sheets.length + ' tabs checked).');
  return true;
}

/**
 * Run ONCE by hand for sheets that already exist: removes their empty rows/columns so the workbooks stay far
 * below the 10,000,000 cell limit. Continues by itself until every sheet is done. Safe to run again.
 */
function trimAllSheets() {
  batchSheets_('trimAllSheets', function (sh) {
    var last = sh.getLastRow();
    if (last < HEADER_ROW) return;
    trimSheet_(sh, last, sh.getLastColumn());
  });
}

/** Shows how many of the 10,000,000 cells each month workbook uses (see View > Logs). */
function checkWorkbookCells() {
  var grand = 0;
  allBooks_().forEach(function (b) {
    var sheets = b.book.getSheets(), total = 0;
    sheets.forEach(function (s) { total += s.getMaxRows() * s.getMaxColumns(); });
    Logger.log(b.tok + ': ' + sheets.length + ' tabs, ' + total + ' of 10,000,000 cells used (' + Math.round(total / 100000) + '%).');
    grand += total;
  });
  return grand;
}

/* ---------- create this month's sheets for every teacher ---------- */

/**
 * Creates the current month's sheet for every section/subject in the Timetable (skips existing ones),
 * inside the workbook of the current month.
 * With hundreds of sections it works in rounds of about 4 minutes and continues by itself until done.
 */
function createCurrentMonthSheets() {
  var tok = currentSeasonToken_();
  if (!tok) return;
  clearContinuation_('createMonthSheetsContinue');
  var start = Date.now(), more = false;
  var book = bookFor_(tok);
  var have = {};
  book.getSheets().forEach(function (s) { have[s.getName()] = true; });
  var admin = { teacher: '', username: ADMIN_USER };
  var list = getAssignments_(admin);
  for (var i = 0; i < list.length; i++) {
    var a = list[i];
    var meta = { name: makeName_(a.cls, tok, a.subject), cls: a.cls, sub: a.subject, mon: tok, teacher: a.teacher };
    if (have[meta.name]) continue;
    if (Date.now() - start > BATCH_MS) { more = true; break; }
    try { ensureSheet_(book, meta); have[meta.name] = true; }
    catch (err) { Logger.log('Skipped ' + meta.name + ': ' + err.message); }
  }
  if (more) {
    ScriptApp.newTrigger('createMonthSheetsContinue').timeBased().after(60 * 1000).create();
    Logger.log('createCurrentMonthSheets: continuing automatically in 1 minute.');
  }
}

/** Continuation of createCurrentMonthSheets (scheduled by itself, do not run by hand). */
function createMonthSheetsContinue() {
  createCurrentMonthSheets();
}

/** Run ONCE by hand: creates the sheets automatically on the 1st of every month. */
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'createCurrentMonthSheets') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('createCurrentMonthSheets').timeBased().onMonthDay(1).atHour(1).create();
  createCurrentMonthSheets();   // also create this month's sheets right now
}

/** Run once by hand: asks Google for permission to open ALL five month workbooks and export PDFs. */
function authorizePdf() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  allBooks_().forEach(function (b) { Logger.log(b.tok + ': ' + b.book.getName()); });
  Logger.log(ScriptApp.getOAuthToken() ? 'OK' : 'no token');
}

/* ---------- one-time maintenance helpers (run by hand if needed; they continue by themselves) ---------- */

/**
 * Run ONCE by hand after installing v5.4: gives every existing attendance sheet the new T.A formula
 * (counts 2P, 3P ...) and the new colour rules. Continues by itself until every sheet is done. Safe to run again.
 * Then run recalcLecturesAllSheets().
 */
function upgradePeriodFormatAllSheets() {
  batchSheets_('upgradePeriodFormatAllSheets', function (sh) {
    var n = sh.getLastRow() - FIRST_ROW + 1; if (n < 1) return;
    var f = [];
    for (var r = 0; r < n; r++) f.push([taFormula_(FIRST_ROW + r)]);
    sh.getRange(FIRST_ROW, TA_COL, n, 1).setFormulas(f);
    applyRules_(sh, n);
  });
}

function addLowAttendanceRule() {
  var L = colLetter_(TA_COL);
  batchSheets_('addLowAttendanceRule', function (sh) {
    var n = sh.getLastRow() - FIRST_ROW + 1; if (n < 1) return;
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + L + '$2>0,$' + L + FIRST_ROW + '/$' + L + '$2<0.73)')
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([sh.getRange(FIRST_ROW, 4, n, 1), sh.getRange(FIRST_ROW, TA_COL, n, 1)]).build();
    var rules = sh.getConditionalFormatRules(); rules.push(rule); sh.setConditionalFormatRules(rules);
  });
}

function addMtLowRule() {
  var C = colLetter_(MT_COL);
  batchSheets_('addMtLowRule', function (sh) {
    var n = sh.getLastRow() - FIRST_ROW + 1; if (n < 1) return;
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($' + C + FIRST_ROW + '),$' + C + '$2>0,$' + C + FIRST_ROW + '/$' + C + '$2<0.5)')
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([sh.getRange(FIRST_ROW, MT_COL, n, 1)]).build();
    var rules = sh.getConditionalFormatRules(); rules.push(rule); sh.setConditionalFormatRules(rules);
  });
}

/** Updates every existing attendance sheet (removes TOTAL LECTURES label, adds T.L and weekday row). */
function fixHeaderRowsAllSheets() {
  batchSheets_('fixHeaderRowsAllSheets', function (sh) {
    var title = String(sh.getRange(1, 1).getValue() || '');
    if (title.indexOf(TITLE_PREFIX) !== 0) return;
    var mt = title.match(/-\s*([A-Za-z]+)\s+(\d{4})/);
    if (!mt) return;
    var month = MONTH_NAMES.indexOf(mt[1].toUpperCase());
    if (month < 0) return;
    var L = readLayout_(sh), rg = dayColRange_(L);
    if (!rg) return;

    var row2 = sh.getRange(2, 1, 1, Math.max(sh.getLastColumn(), rg.c2 + 1)).getValues()[0];
    row2.forEach(function (v, i) {
      if (/TOTAL\s*LECTURES/i.test(String(v))) sh.getRange(2, i + 1).clearContent();
    });
    sh.getRange(1, rg.c2 + 1).setValue('T.L').setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
    setWeekdayRow_(sh, month, Number(mt[2]), L.dayCols[1] || rg.c1);
  });
}
function debugPeriods() {
  var name = 'FCS-1 D2 OCT COMPUTER SCIENCE';
  var m = allowed_({username: ADMIN_USER, teacher: ''})[name];
  var pd = pdLoad_(bookFor_(m.mon), name);
  Logger.log('row=' + pd.row + ' data=' + JSON.stringify(pd.d));
}

/* ---------- guest students ---------- */
var GUESTS_SHEET = 'Guests';
var GUEST_HEAD = ['Roll','Name','Father','Home class','Host class','Subject','Added by','Added on','From month','Status'];

function guestsSheet_() {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(GUESTS_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(GUESTS_SHEET);
  sh.getRange(1, 1, 1, GUEST_HEAD.length).setValues([GUEST_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, 2000, GUEST_HEAD.length).setNumberFormat('@');
  return sh;
}
function guestRows_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(GUESTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, GUEST_HEAD.length).getValues();
}
function guestNote_(home, by, on) { return 'GUEST|' + home + '|' + by + '|' + on; }

/** Active guests of a class+subject that apply to this month (used when a new month sheet is created). */
function activeGuests_(cls, sub, mon) {
  var out = [], mi = SEASON.indexOf(mon);
  guestRows_().forEach(function (r) {
    if (String(r[9]) !== 'Active' || String(r[4]) !== cls || String(r[5]).toUpperCase() !== String(sub).toUpperCase()) return;
    if (SEASON.indexOf(String(r[8])) > mi) return;
    out.push({ roll: String(r[0]).trim(), name: r[1], father: r[2], home: r[3], by: r[6], on: r[7] });
  });
  return out;
}
/** roll -> [{cls, sub, by}] : where else this student is a guest (shown to her own teacher). */
/** "2", "P2", "II", "Period 2", "2nd" -> "2" */
function normPeriod_(p) {
  p = String(p || '').trim().toUpperCase().replace(/^P(ERIOD)?\s*[-.:]?\s*/, '');
  var R = { I: 1, II: 2, III: 3, IV: 4, V: 5, VI: 6, VII: 7, VIII: 8 };
  if (R[p]) return String(R[p]);
  var m = p.match(/\d+/);
  return m ? String(Number(m[0])) : '';
}
var DAYCODE_ = { M: 1, MO: 1, MON: 1, T: 2, TU: 2, TUE: 2, W: 3, WE: 3, WED: 3, TH: 4, THU: 4, THUR: 4, F: 5, FR: 5, FRI: 5, S: 6, SA: 6, SAT: 6 };
var DAYNAME_ = ['', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
/** "M-TH" -> {1,2,3,4}; "F,S" -> {5,6}; "W-S" -> {3..6}; empty -> every day (Mon-Sat). */
function daySet_(txt) {
  var out = {}, any = false;
  String(txt || '').toUpperCase().replace(/\s*-\s*/g, '-').split(/[,;\/&+\s]+/).forEach(function (part) {
    if (!part) return;
    var rg = part.split('-');
    if (rg.length === 2 && DAYCODE_[rg[0]] && DAYCODE_[rg[1]]) {
      for (var d = DAYCODE_[rg[0]]; d <= DAYCODE_[rg[1]]; d++) { out[d] = 1; any = true; }
    } else if (DAYCODE_[part]) { out[DAYCODE_[part]] = 1; any = true; }
  });
  if (!any) for (var k = 1; k <= 6; k++) out[k] = 1;
  return out;
}
/** {1,2,3,4} -> "Mon-Thu"; {5,6} -> "Fri-Sat"; {1,3} -> "Mon, Wed"; all six -> "" */
function dayLabel_(set) {
  var ds = Object.keys(set).map(Number).sort(function (a, b) { return a - b; });
  if (ds.length >= 6) return '';
  var parts = [], i = 0;
  while (i < ds.length) {
    var j = i; while (j + 1 < ds.length && ds[j + 1] === ds[j] + 1) j++;
    parts.push(j - i >= 1 ? DAYNAME_[ds[i]] + '-' + DAYNAME_[ds[j]] : DAYNAME_[ds[i]]);
    i = j + 1;
  }
  return parts.join(', ');
}
/** Timetable slots of one class + subject: [{ p:'2', days:{1:1,..} }] (one per period of each timetable row). */
var SLOTS_MEMO_ = {};
function slotsOf_(cls, sub) {
  var mk = String(cls) + '|' + String(sub).toUpperCase();
  if (SLOTS_MEMO_[mk]) return SLOTS_MEMO_[mk];
  var t = getTimetable_(), out = [];
  var c = String(cls || '').replace(/\s+/g, '').toUpperCase(), s = String(sub || '').trim().toUpperCase();
  t.rows.forEach(function (r) {
    if (String(r[3] || '').replace(/\s+/g, '').toUpperCase() !== c || String(r[4] || '').toUpperCase() !== s) return;
    var days = daySet_([r[6], r[7], r[8]].filter(function (x) { return String(x || '').trim(); }).join(','));
    String(r[5] || '').split(/[,\/&+;]|\band\b/i).forEach(function (x) {
      var rg = String(x).split('-').map(normPeriod_).filter(Boolean);
      if (rg.length === 2 && Number(rg[1]) > Number(rg[0]) && Number(rg[1]) - Number(rg[0]) < 6) {
        for (var n = Number(rg[0]); n <= Number(rg[1]); n++) out.push({ p: String(n), days: days });
      } else rg.forEach(function (n) { out.push({ p: n, days: days }); });
    });
  });
  SLOTS_MEMO_[mk] = out;
  return out;
}
/** Where the two classes meet at the same period on the same day: [{ p, days:{..} }] (empty = no clash, or timetable incomplete). */
function clashSlots_(hostCls, hostSub, myCls, mySub) {
  var a = slotsOf_(hostCls, hostSub), b = slotsOf_(myCls, mySub), out = [];
  a.forEach(function (x) {
    b.forEach(function (y) {
      if (x.p !== y.p) return;
      var both = {}, any = false;
      Object.keys(x.days).forEach(function (d) { if (y.days[d]) { both[d] = 1; any = true; } });
      if (any) out.push({ p: x.p, days: both });
    });
  });
  return out;
}
/** roll -> [{ cls, sub, by, p:'P2 (Mon-Thu)' }].
 *  Shown ONLY to the teacher of her own section whose lecture is at the same period on the same days as her guest class
 *  (that teacher will not see her in that lecture). Every other teacher of her section sees nothing.
 *  The teacher who added the guest sees "Guest from ..." on her own sheet, as before. */
function guestAtMap_(meta) {
  var map = {};
  guestRows_().forEach(function (r) {
    if (String(r[9]) !== 'Active') return;
    if (String(r[4]) === meta.cls && String(r[5]).toUpperCase() === meta.sub.toUpperCase()) return;
    var cl = [];
    try { cl = clashSlots_(r[4], r[5], meta.cls, meta.sub); } catch (e) {}
    if (!cl.length) return;
    var byP = {};
    cl.forEach(function (x) { byP[x.p] = byP[x.p] || {}; Object.keys(x.days).forEach(function (d) { byP[x.p][d] = 1; }); });
    var label = Object.keys(byP).sort(function (a, b) { return a - b; }).map(function (p) { var dl = dayLabel_(byP[p]); return 'P' + p + (dl ? ' (' + dl + ')' : ''); }).join(', ');
    var roll = String(r[0]).trim();
    (map[roll] = map[roll] || []).push({ cls: r[4], sub: r[5], by: r[6], p: label, clash: true });
  });
  return map;
}

function guestSearch_(user, name, q) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  q = String(q || '').trim().toLowerCase();
  if (q.length < 2) return { ok: true, results: [] };
  var sheet = bookFor_(meta.mon).getSheetByName(name), have = {};
  if (sheet && sheet.getLastRow() >= FIRST_ROW) {
    var L = readLayout_(sheet), cnt = sheet.getLastRow() - FIRST_ROW + 1;
    var rg = sheet.getRange(FIRST_ROW, L.rollCol, cnt, 1), v = rg.getValues(), nt = rg.getNotes();
    for (var i = 0; i < cnt; i++) if (!isOffNote_(nt[i][0])) have[String(v[i][0]).trim()] = true;
  }
  var out = [];
  getMaster_().rows.forEach(function (r) {
    if (out.length >= 15 || have[r[0]]) return;
    if (String(r[1]).toLowerCase().indexOf(q) === -1 && String(r[0]).toLowerCase().indexOf(q) === -1) return;
    out.push({ roll: r[0], name: r[1], cls: r[2] });
  });
  return { ok: true, results: out };
}

/** Adds (or re-enables) the guest row in one month sheet. m = master row [roll, name, classSec, subs, father]. */
function putGuestRow_(sh, roll, m, by, on) {
  var L = readLayout_(sh), last = sh.getLastRow(), n = last - FIRST_ROW + 1, row = 0;
  if (n < 1) return;
  var rolls = sh.getRange(FIRST_ROW, L.rollCol, n, 1).getValues();
  for (var i = 0; i < n; i++) if (String(rolls[i][0]).trim() === roll) { row = FIRST_ROW + i; break; }
  if (row) {                                   // she was a guest before and was removed: switch her on again
    sh.getRange(row, L.rollCol).setNote(guestNote_(m[2], by, on));
    sh.getRange(row, 1, 1, 6).setBackground(null).setFontColor('#000000');
    sh.getRange(row, 4).setFontLine('none');
    return;
  }
  row = last + 1;
  if (sh.getMaxRows() < row) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
  sh.getRange(last, 1, 1, MT_COL).copyTo(sh.getRange(row, 1, 1, MT_COL));   // same look and T.A formula
  sh.getRange(row, DAY_COL1, 1, 31).clearContent();
  sh.getRange(row, MT_COL).clearContent();
  sh.getRange(row, 1, 1, 6).setBackground(null).setFontColor('#000000');
  sh.getRange(row, 4).setFontLine('none');
  sh.getRange(row, L.rollCol).clearNote().setNumberFormat('@');
  sh.getRange(row, 1, 1, 6).setValues([[n + 1, roll, sectionOf_(m[2]), m[1], m[4] || '', m[2]]]);
  sh.getRange(row, L.rollCol).setNote(guestNote_(m[2], by, on));
  applyRules_(sh, row - FIRST_ROW + 1);
}

function guestAdd_(user, name, roll) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  roll = String(roll || '').trim();
  var m = getMaster_().rows.filter(function (r) { return r[0] === roll; })[0];
  if (!m) return { ok: false, error: 'Student not found.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var gs = guestsSheet_(), already = false;
    guestRows_().forEach(function (r) {
      if (String(r[0]).trim() === roll && String(r[4]) === meta.cls && String(r[5]).toUpperCase() === meta.sub.toUpperCase() && String(r[9]) === 'Active') already = true;
    });
    var on = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy');
    if (!already) gs.appendRow([roll, m[1], m[4], m[2], meta.cls, meta.sub, user.teacher, on, meta.mon, 'Active']);
    var mi = SEASON.indexOf(meta.mon);
    SEASON.forEach(function (tok, k) {
      if (k < mi) return;
      try {
        var sh = bookFor_(tok).getSheetByName(makeName_(meta.cls, tok, meta.sub));
        if (sh) putGuestRow_(sh, roll, m, user.teacher, on);
      } catch (e) { Logger.log('guest_add ' + tok + ': ' + e.message); }
    });
    SpreadsheetApp.flush();
    return { ok: true, already: already };
  } finally { lock.releaseLock(); }
}

/** Switches the guest off from this month onwards (old attendance stays). Caller must hold the lock. */
function removeGuestCore_(meta, roll) {
  var gs = guestsSheet_(), n = gs.getLastRow() - 1;
  if (n > 0) {
    var v = gs.getRange(2, 1, n, GUEST_HEAD.length).getValues();
    for (var i = 0; i < n; i++) {
      if (String(v[i][0]).trim() === roll && String(v[i][4]) === meta.cls && String(v[i][5]).toUpperCase() === meta.sub.toUpperCase() && String(v[i][9]) === 'Active') gs.getRange(i + 2, 10).setValue('Removed');
    }
  }
  var mi = SEASON.indexOf(meta.mon);
  SEASON.forEach(function (tok, k) {
    if (k < mi) return;
    try {
      var sh = bookFor_(tok).getSheetByName(makeName_(meta.cls, tok, meta.sub));
      if (!sh) return;
      var L = readLayout_(sh), cnt = sh.getLastRow() - FIRST_ROW + 1;
      if (cnt < 1) return;
      var rolls = sh.getRange(FIRST_ROW, L.rollCol, cnt, 1).getValues();
      for (var j = 0; j < cnt; j++) {
        if (String(rolls[j][0]).trim() !== roll) continue;
        var row = FIRST_ROW + j;
        sh.getRange(row, L.rollCol).setNote(OFF_TAG + ' - guest removed');
        sh.getRange(row, 1, 1, 6).setBackground('#EFEFEF').setFontColor('#999999');
        sh.getRange(row, 4).setFontLine('line-through');
      }
    } catch (e) { Logger.log('guest_remove ' + tok + ': ' + e.message); }
  });
  SpreadsheetApp.flush();
  return { ok: true, guestRemoved: true };
}
function guestRemove_(user, name, roll) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return removeGuestCore_(meta, String(roll || '').trim()); } finally { lock.releaseLock(); }
}
/* =====================================================================================================
   v6.0  DASHBOARDS: attendance index + college / department overview
   One row per attendance sheet in the sheet "Attendance Index" (Users workbook). It is refreshed every
   time a teacher saves or deletes attendance, so the overview reads ONE sheet instead of 1,600.
   Run rebuildAttendanceIndex() once by hand after installing v6.0 (it continues by itself).
   ===================================================================================================== */
var INDEX_SHEET = 'Attendance Index';
var INDEX_HEAD  = ['Sheet', 'Month', 'Class', 'Subject', 'Teacher', 'Students', 'Lectures held', 'Attended (sum)',
                   'Possible (sum)', 'Avg %', 'Below 73%', 'Marked days', 'Low students (do not edit)', 'Updated',
                   'Test total', 'Students data (do not edit)'];
var LOW_PCT = 73;

var INDEX_OK_ = false;
function indexSheet_() {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(INDEX_SHEET);
  if (sh) {
    if (!INDEX_OK_ && String(sh.getRange(1, INDEX_HEAD.length).getValue()) !== INDEX_HEAD[INDEX_HEAD.length - 1]) {
      if (sh.getMaxColumns() < INDEX_HEAD.length) sh.insertColumnsAfter(sh.getMaxColumns(), INDEX_HEAD.length - sh.getMaxColumns());
      sh.getRange(1, 1, 1, INDEX_HEAD.length).setValues([INDEX_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
    }
    INDEX_OK_ = true;
    return sh;
  }
  sh = book.insertSheet(INDEX_SHEET);
  sh.getRange(1, 1, 1, INDEX_HEAD.length).setValues([INDEX_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.setFrozenRows(1);
  sh.getRange(2, 12, 1000, 1).setNumberFormat('@');      // "1,2,3" must stay text
  sh.setColumnWidth(1, 300); sh.setColumnWidth(13, 300);
  return sh;
}

/** Lectures a guest was marked in (P or A) from her row of day cells: P=1, 2P=2, 3A=3 ... (L and B count 0). */
function guestHeld_(cells) {
  var n = 0;
  cells.forEach(function (c) {
    var m = String(c || '').trim().toUpperCase().match(/^(\d*)([PA])$/);
    if (m) n += m[1] ? Number(m[1]) : 1;
  });
  return n;
}

/** Re-reads one attendance sheet and writes its summary row to the index. meta = { name, mon, cls, sub, teacher }. */
function indexUpdate_(sh, meta) {
  if (!sh || !meta) return;
  var L = readLayout_(sh), n = sh.getLastRow() - FIRST_ROW + 1, rg = dayColRange_(L);
  var row = [meta.name || sh.getName(), meta.mon, meta.cls, meta.sub, meta.teacher || '', 0, 0, 0, 0, '', 0, '', '[]', new Date(), 0, '[]'];
  if (n > 0 && rg) {
    var dayW = rg.c2 - rg.c1 + 1;
    var block = sh.getRange(FIRST_ROW, rg.c1, n, dayW + 1).getValues();      // all days + T.A
    var rr = sh.getRange(FIRST_ROW, L.rollCol, n, 1), rolls = rr.getValues(), rnotes = rr.getNotes();
    var names = L.nameCol > 0 ? sh.getRange(FIRST_ROW, L.nameCol, n, 1).getValues() : null;
    var held = Number(sh.getRange(2, rg.c2 + 1).getValue()) || 0;
    var hdrNotes = sh.getRange(HEADER_ROW, rg.c1, 1, dayW).getNotes()[0];
    var mts = L.mtCol > 0 ? sh.getRange(FIRST_ROW, L.mtCol, n, 1).getValues() : null;
    var mtTotal = L.mtCol > 0 ? Number(sh.getRange(2, L.mtCol).getValue()) || 0 : 0;
    var all = [];                                    // every student: [roll, name, attended, held, test, off(1/0), guest home]
    var marked = [];
    Object.keys(L.dayCols).forEach(function (d) {
      var j = L.dayCols[d] - rg.c1, any = !!hdrNotes[j];                     // a note (college closed ...) counts as done
      for (var i = 0; i < n && !any; i++) if (String(rolls[i][0]).trim() && String(block[i][j]).trim()) any = true;
      if (any) marked.push(Number(d));
    });
    marked.sort(function (a, b) { return a - b; });
    var students = 0, sumP = 0, sumPos = 0, low = [];
    for (var i = 0; i < n; i++) {
      var roll = String(rolls[i][0]).trim();
      if (!roll) continue;
      var note = String(rnotes[i][0] || ''), off = isOffNote_(note), guest = note.indexOf('GUEST|') === 0;
      var ta = Number(block[i][dayW]) || 0;
      var den = guest ? guestHeld_(block[i].slice(0, dayW)) : held;
      var mv = mts ? String(mts[i][0]).trim() : '';
      all.push([roll, names ? String(names[i][0]).trim() : roll, ta, den, mv === '' ? '' : (/^A/i.test(mv) ? 'A' : Number(mv)), off ? 1 : 0, guest ? (note.split('|')[1] || '') : '']);
      if (off) continue;
      students++;
      if (den > 0) {
        sumP += ta; sumPos += den;
        if (ta / den * 100 < LOW_PCT) low.push([roll, names ? String(names[i][0]).trim() : roll, ta, den]);
      }
    }
    var lowJson = JSON.stringify(low);
    while (lowJson.length > 45000 && low.length) { low.pop(); lowJson = JSON.stringify(low); }
    row[5] = students; row[6] = held; row[7] = sumP; row[8] = sumPos;
    row[9] = sumPos ? Math.round(sumP / sumPos * 1000) / 10 : '';
    row[10] = low.length; row[11] = marked.join(','); row[12] = lowJson;
    row[14] = mtTotal; row[15] = JSON.stringify(all).slice(0, 49000);
  }
  var ix = indexSheet_(), last = ix.getLastRow(), r = 0;
  if (last > 1) {
    var f = ix.getRange(2, 1, last - 1, 1).createTextFinder(row[0]).matchCase(true).matchEntireCell(true).findNext();
    if (f) r = f.getRow();
  }
  if (!r) r = last + 1;
  ix.getRange(r, 12).setNumberFormat('@');
  ix.getRange(r, 1, 1, INDEX_HEAD.length).setValues([row]);
}

/** Run ONCE by hand after installing v6.1: fills the index for every month from September up to now
 *  (continues by itself in rounds of 4 minutes). Safe to run again. */
function rebuildAttendanceIndex() {
  var metas = allowed_({ username: ADMIN_USER, teacher: '' });
  batchSheets_('rebuildAttendanceIndex', function (sh) {
    var meta = metas[sh.getName()];
    if (meta) indexUpdate_(sh, meta);
  }, activeMonths_());
}

/** 'all' (admin, principal), 'dept' (HOD) or '' (teacher). */
function scopeOf_(user) {
  if (isAdmin_(user)) return { all: true, dept: '' };
  var role = String(user.role || '').toUpperCase();
  if (/ADMIN|PRINCIPAL/.test(role)) return { all: true, dept: '' };
  if (/\bHOD\b|HEAD OF DEP/.test(role)) return { all: false, dept: String(user.dept || '').trim().toUpperCase() };
  return null;
}
function scopeName_(user) { var s = scopeOf_(user); return s ? (s.all ? 'all' : 'dept') : ''; }
function whoami_(user) {
  var s = scopeOf_(user), out = { ok: true, view: scopeName_(user), dept: s && !s.all ? user.dept : '' };
  try { out.cp = isCP_(user); } catch (e) { out.cp = false; }
  try { out.tt = myTimetable_(user); } catch (e) { out.tt = []; }
  try { out.times = periodTimes_(); } catch (e) { out.times = []; }
  return out;
}

/** normalized teacher name -> department (from the Users sheet). */
function teacherDepts_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users'), map = {};
  if (!sh) return map;
  var v = sh.getDataRange().getValues(), iD = usersCol_(v[0], 'department', 3);
  for (var i = 1; i < v.length; i++) if (v[i][0]) map[normalizeName(v[i][0])] = String(v[i][iD] || '').trim();
  return map;
}

/** name -> index row, for one month. */
function indexRows_(tok) {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(INDEX_SHEET), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, INDEX_HEAD.length).getValues().forEach(function (r) {
    if (String(r[1]) === tok) out[String(r[0])] = r;
  });
  return out;
}

/** The sections this admin / HOD may see in the current month: [{ a, dept, name, x }] */
function scopedSections_(user) {
  var scope = scopeOf_(user);
  if (!scope) return null;
  var tok = currentSeasonToken_();
  if (!tok) return { tok: null, list: [], scope: scope };
  var depts = teacherDepts_(), idx = indexRows_(tok), list = [];
  getAssignments_({ teacher: '', username: ADMIN_USER }).forEach(function (a) {
    var dept = depts[normalizeName(a.teacher)] || '';
    if (!scope.all && dept.toUpperCase() !== scope.dept) return;
    var name = makeName_(a.cls, tok, a.subject);
    list.push({ a: a, dept: dept, name: name, x: idx[name] || null });
  });
  return { tok: tok, list: list, scope: scope };
}

/** Admin / HOD overview of the current month: who marked today, section and subject percentages. */
function overview_(user) {
  var S = scopedSections_(user);
  if (!S) return { ok: false, error: 'NOT_ALLOWED' };
  var now = new Date(), day = now.getDate();
  var res = { ok: true, tok: S.tok, day: day, sunday: now.getDay() === 0, scope: S.scope.all ? 'College' : (S.scope.dept || 'Department'), rows: [] };
  if (!S.tok) return res;
  var my = monthYear_(S.tok);
  res.m = my.m; res.y = my.y;
  var dow = now.getDay();
  S.list.forEach(function (s) {
    var x = s.x, days = x ? String(x[11]).split(',').map(Number).filter(Boolean) : [];
    var sl = [], sched = null;
    try { sl = slotsOf_(s.a.cls, s.a.subject); } catch (e) {}
    if (sl.length) sched = sl.some(function (q) { return !!q.days[dow]; });   // null = not in the timetable
    res.rows.push({ sched: sched,
      name: s.name, cls: s.a.cls, sub: s.a.subject, teacher: s.a.teacher, dept: s.dept,
      students: x ? Number(x[5]) || 0 : 0, held: x ? Number(x[6]) || 0 : 0,
      att: x ? Number(x[7]) || 0 : 0, pos: x ? Number(x[8]) || 0 : 0,
      avg: x && x[9] !== '' ? Number(x[9]) : null, low: x ? Number(x[10]) || 0 : 0,
      today: days.indexOf(day) !== -1, last: days.length ? days[days.length - 1] : 0, marked: days.length
    });
  });
  return res;
}

/** Students below 73% this month in every section this admin / HOD may see. */
function overviewLow_(user) {
  var S = scopedSections_(user);
  if (!S) return { ok: false, error: 'NOT_ALLOWED' };
  var out = [];
  S.list.forEach(function (s) {
    if (!s.x) return;
    var low = [];
    try { low = JSON.parse(s.x[12] || '[]'); } catch (e) {}
    low.forEach(function (l) {
      out.push({ roll: l[0], name: l[1], p: l[2], held: l[3], pct: Math.round(l[2] / l[3] * 1000) / 10,
                 cls: s.a.cls, sub: s.a.subject, teacher: s.a.teacher, dept: s.dept, sheet: s.name });
    });
  });
  out.sort(function (a, b) { return a.pct - b.pct; });
  return { ok: true, tok: S.tok, rows: out.slice(0, 4000), total: out.length };
}

/* =====================================================================================================
   v6.1  REPORTS: per-student attendance and monthly test data for any month range, from the index.
   Teacher: only her own sections/subjects.  HOD: her department.  Admin / Principal: whole college.
   req: { from:'SEPT', to:'OCT', cls:'FCS-1 D2' (optional), sub:'COMPUTER SCIENCE' (optional), roll (optional) }
   ===================================================================================================== */
function reportData_(user, req) {
  var scope = scopeOf_(user), own = null, depts = null;
  if (!scope) { own = {}; listSheets_(user).forEach(function (s) { own[s.name] = true; }); }
  else if (!scope.all) depts = teacherDepts_();
  var act = activeMonths_(), fi = Math.max(0, SEASON.indexOf(String(req.from || 'SEPT'))), ti = SEASON.indexOf(String(req.to || ''));
  if (ti < 0 || ti > act.length - 1) ti = act.length - 1;
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(INDEX_SHEET), out = [];
  var wantCls = String(req.cls || ''), wantSub = String(req.sub || '').toUpperCase(), wantRoll = String(req.roll || '');
  if (sh && sh.getLastRow() > 1) {
    var w = Math.min(sh.getLastColumn(), INDEX_HEAD.length);
    sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues().forEach(function (r) {
      var name = String(r[0]), mon = String(r[1]), mi = SEASON.indexOf(mon);
      if (mi < fi || mi > ti) return;
      if (wantCls && String(r[2]) !== wantCls) return;
      if (wantSub && String(r[3]).toUpperCase() !== wantSub) return;
      if (own && !own[name]) return;
      if (depts && String(depts[normalizeName(r[4])] || '').toUpperCase() !== scope.dept) return;
      var st = [];
      try { st = JSON.parse(r[15] || '[]'); } catch (e) {}
      if (wantRoll) { st = st.filter(function (x) { return x[0] === wantRoll; }); if (!st.length) return; }
      out.push({ cls: String(r[2]), sub: String(r[3]), teacher: String(r[4]), mon: mon, held: Number(r[6]) || 0, mt: Number(r[14]) || 0, st: st });
    });
  }
  return { ok: true, from: SEASON[fi], to: SEASON[ti], view: scope ? (scope.all ? 'all' : 'dept') : '', rows: out };
}

/* =====================================================================================================
   v6.2  TIMETABLE IN THE APP: period timings, the teacher's own timetable, workload
   ===================================================================================================== */
/* Period timings and evening sections come from the timetable editor's own settings sheet ("NTT Settings",
   edited in the timetable web interface: Settings -> Timings / Evening sections). Nothing to set up here. */
var NTT_SETTINGS_SHEET = 'NTT Settings';
var DEFAULT_TIMES_ = ['08:00-09:00', '09:00-10:00', '10:00-11:00', '11:00-12:00', '12:00-13:00'];

/** "08:00", "8:00 AM", "01:00" (afternoon), a time cell -> "HH:MM" (24 h). College hours: 1:00-6:59 are taken as afternoon. */
function hhmm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  var m = String(v || '').trim().match(/(\d{1,2})\s*[:.]\s*(\d{2})\s*(am|pm)?/i);
  if (!m) return '';
  var h = Number(m[1]), ap = (m[3] || '').toLowerCase();
  if (ap === 'pm' && h < 12) h += 12; else if (ap === 'am' && h === 12) h = 0; else if (!ap && h >= 1 && h <= 6) h += 12;
  return (h < 10 ? '0' : '') + h + ':' + m[2];
}
/** "08:00-09:00" / "08:00 – 09:00" -> { start:'08:00', end:'09:00' } or null */
function timeRange_(txt) {
  var m = String(txt || '').match(/(\d{1,2}\s*[:.]\s*\d{2}\s*(?:am|pm)?)\s*[–—-]+\s*(\d{1,2}\s*[:.]\s*\d{2}\s*(?:am|pm)?)/i);
  return m ? { start: hhmm_(m[1]), end: hhmm_(m[2]) } : null;
}
var NTT_MEMO_ = null;
/** { times:{ M:[{p,start,end}], E:[...] }, evening:{ 'PROGRAM|CLASS|SECTION':1, 'PROGRAM|CLASS|*':1 } } */
function nttSettings_() {
  if (NTT_MEMO_) return NTT_MEMO_;
  var c = cacheGetBig_('ptimes');
  if (c) return (NTT_MEMO_ = c);
  var out = { times: { M: [], E: [] }, evening: {} }, raw = { M: [], E: [] };
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(NTT_SETTINGS_SHEET);
  if (sh) sh.getDataRange().getValues().forEach(function (r, i) {
    if (!i) return;
    var type = String(r[0] || '').trim().toUpperCase(), key = String(r[1] || '').trim();
    if (type === 'EVENING' && key) out.evening[key.split('|').map(function (x) { return String(x).trim().toUpperCase().replace(/\s+/g, ' '); }).join('|')] = 1;
    if (type === 'TIME') { var m = key.match(/^([ME])\|(\d+)\|(.*)$/); if (m) raw[m[1]][Number(m[2]) - 1] = m[3]; }
  });
  if (!raw.M.some(Boolean)) raw.M = DEFAULT_TIMES_.slice();
  ['M', 'E'].forEach(function (k) {
    var n = Math.max(raw.M.length, raw[k].length);
    for (var i = 0; i < n; i++) {
      var t = timeRange_(raw[k][i]) || (k === 'E' ? null : null);      // an evening period left blank uses the morning timing (handled in the app)
      if (t && t.start && t.end) out.times[k].push({ p: String(i + 1), start: t.start, end: t.end });
    }
  });
  cachePutBig_('ptimes', out);
  return (NTT_MEMO_ = out);
}
/** kept for the app: { M:[...], E:[...] } */
function periodTimes_() { return nttSettings_().times; }
/** true when this Program / Class / Section is an evening section in the timetable settings */
function isEvening_(program, cls, section) {
  var ev = nttSettings_().evening, n = function (x) { return String(x || '').trim().toUpperCase().replace(/\s+/g, ' '); };
  return !!(ev[n(program) + '|' + n(cls) + '|' + n(section)] || ev[n(program) + '|' + n(cls) + '|*']);
}
/** "M-TH" -> [1,2,3,4]; empty -> [] */
function dayList_(txt) {
  if (!String(txt || '').trim()) return [];
  return Object.keys(daySet_(txt)).map(Number).sort(function (a, b) { return a - b; });
}
/** "2" -> ['2'], "5-6" -> ['5','6'], "II, IV" -> ['2','4'] */
function periodList_(txt) {
  var out = [];
  String(txt || '').split(/[,\/&+;]|\band\b/i).forEach(function (x) {
    var rg = String(x).split('-').map(normPeriod_).filter(Boolean);
    if (rg.length === 2 && Number(rg[1]) > Number(rg[0]) && Number(rg[1]) - Number(rg[0]) < 6) {
      for (var n = Number(rg[0]); n <= Number(rg[1]); n++) if (out.indexOf(String(n)) === -1) out.push(String(n));
    } else rg.forEach(function (n) { if (out.indexOf(n) === -1) out.push(n); });
  });
  return out;
}
/** One timetable row for the app: { teacher, cls, sub, ps:['2'], th:[days], pr:[days], room, course, program } */
function ttRow_(r) {
  var th = dayList_(r[6]), pr = dayList_(r[7]);
  if (!th.length && String(r[8] || '').trim()) th = dayList_(r[8]);
  if (!th.length && !pr.length) th = [1, 2, 3, 4, 5, 6];
  var eve = false;
  try { eve = isEvening_(r[1], r[2], r[11]) || /\(\s*EVENING\s*\)/i.test(String(r[9] || '')); } catch (e) {}
  return { teacher: r[0], cls: r[3], sub: String(r[4] || '').toUpperCase(), ps: periodList_(r[5]), th: th, pr: pr,
           room: r[9] || '', course: r[10] || '', program: r[1] || '', sh: eve ? 'E' : 'M' };
}
/** This teacher's own timetable rows. */
function myTimetable_(user) {
  var me = normalizeName(user.teacher), out = [];
  if (!me) return out;
  getTimetable_().rows.forEach(function (r) {
    if (r[3] && r[4] && normalizeName(r[0]) === me) out.push(ttRow_(r));
  });
  return out;
}
/** Admin / Principal / Chief Proctors: the whole timetable. HOD: her department. For Workload and Free teachers. */
function ttAll_(user) {
  var scope = scopeOf_(user), own = scope && !scope.all ? scope.dept : '';
  if (isCP_(user)) scope = { all: true, dept: '' };                    // Chief Proctors (also an HOD who is CP) see the whole college: duties go to any department
  if (!scope) return { ok: false, error: 'NOT_ALLOWED' };
  var depts = teacherDepts_(), rows = [];
  getTimetable_().rows.forEach(function (r) {
    if (!r[0] || !r[3] || !r[4]) return;
    var dept = depts[normalizeName(r[0])] || '';
    if (!scope.all && dept.toUpperCase() !== scope.dept) return;
    var x = ttRow_(r); x.dept = dept; rows.push(x);
  });
  var info = {}, all = usersInfo_();
  rows.forEach(function (x) {                                         // no phone numbers here: they go only through 'contacts'
    var k = normalizeName(x.teacher), a = all[k];
    if (a && !info[k]) info[k] = { status: a.status, desig: a.desig, dept: a.dept };
  });
  return { ok: true, v: 2, rows: rows, times: periodTimes_(), info: info, wdept: own };   // wdept: an HOD-CP's own department (her Workload tab)
}

/* =====================================================================================================
   v6.4  TEACHER CONTACTS: kept apart from every report.
   Admin, principal and Chief Proctors (also when HOD): the whole college. Other HODs: their own department. Everybody else: refused.
   ===================================================================================================== */
function contacts_(user) {
  var scope = scopeOf_(user);
  if (isCP_(user)) scope = { all: true, dept: '' };                    // a Chief Proctor (even if also HOD) reaches every teacher
  if (!scope) return { ok: false, error: 'NOT_ALLOWED' };
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users'), out = [];
  if (!sh) return { ok: true, list: out };
  var v = sh.getDataRange().getValues(), h = v[0];
  var iSt = usersCol_(h, 'status', 1), iDs = usersCol_(h, 'designation', 2), iDp = usersCol_(h, 'department', 3),
      iPh = usersCol_(h, 'contact', 6), iWa = usersCol_(h, 'whatsapplink', 7);
  for (var i = 1; i < v.length; i++) {
    var name = String(v[i][0] || '').trim(); if (!name) continue;
    var dept = String(v[i][iDp] || '').trim();
    if (!scope.all && dept.toUpperCase() !== scope.dept) continue;
    var phone = String(v[i][iPh] || '').trim(), wa = String(v[i][iWa] || '').trim();
    if (!phone && !wa) continue;
    out.push({ name: name, dept: dept, desig: String(v[i][iDs] || '').trim(), status: String(v[i][iSt] || '').trim(), phone: phone, wa: wa });
  }
  out.sort(function (a, b) { return (a.dept || '~').localeCompare(b.dept || '~') || a.name.localeCompare(b.name); });
  return { ok: true, list: out, scope: scope.all ? 'all' : scope.dept };
}

/* =====================================================================================================
   v6.3  FREE TEACHERS (admin, principal, HOD, Chief Proctor)
   ===================================================================================================== */
/** normalized teacher name -> { phone, status, desig, dept } from the Users sheet */
function usersInfo_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users'), map = {};
  if (!sh) return map;
  var v = sh.getDataRange().getValues(), h = v[0];
  var iSt = usersCol_(h, 'status', 1), iDs = usersCol_(h, 'designation', 2), iDp = usersCol_(h, 'department', 3), iPh = usersCol_(h, 'contact', 6);
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    map[normalizeName(v[i][0])] = { phone: String(v[i][iPh] || '').trim(), status: String(v[i][iSt] || '').trim(),
                                    desig: String(v[i][iDs] || '').trim(), dept: String(v[i][iDp] || '').trim() };
  }
  return map;
}
/** true when this user is marked CP / CP1 / CP2 ... in the Chief Proctor column of the Users sheet */
function isCP_(user) {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users');
  if (!sh) return false;
  var v = sh.getDataRange().getValues(), iCP = usersCol_(v[0], 'chiefproctor', 9), me = String(user.username || '').trim().toLowerCase();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][4] || '').trim().toLowerCase() === me) return /^CP\d*$/.test(String(v[i][iCP] || '').trim().toUpperCase());
  }
  return false;
}
/* =====================================================================================================
   v6.6  REGISTERS & RESULTS (Reports screen -> "Registers & results" card)
   - reg_opts : the sections / subjects this user may report on (from the Attendance Index)
   - reg_find : student search (name, father name, roll no) inside those sections
   - reg_plan : the attendance sheets for a choice (months, section or all, subject, student)
   - reg_days : day-by-day cells of up to 30 sheets per call (the app asks in parts and builds the PDF itself)
   Same rights as Reports: teacher = own subjects, HOD = department, admin / principal = whole college.
   ===================================================================================================== */
var REG_MAX_SHEETS = 30;

/** All index rows: { name, mon, cls, sub, teacher, json } */
function regIndex_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(INDEX_SHEET), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var w = Math.min(sh.getLastColumn(), INDEX_HEAD.length);
  sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues().forEach(function (r) {
    if (!r[0]) return;
    out.push({ name: String(r[0]), mon: String(r[1]), cls: String(r[2]), sub: String(r[3]).toUpperCase(),
               teacher: String(r[4] || ''), json: String(r[15] || '') });
  });
  return out;
}

/** function(indexRow) -> true when this user may see that sheet */
function regFilter_(user) {
  var scope = scopeOf_(user), own = null, depts = null;
  if (!scope) { own = {}; listSheets_(user).forEach(function (s) { own[s.name] = true; }); }
  else if (!scope.all) depts = teacherDepts_();
  return function (r) {
    if (own) return !!own[r.name];
    if (depts) return String(depts[normalizeName(r.teacher)] || '').toUpperCase() === scope.dept;
    return true;
  };
}

function regOpts_(user) {
  var ok = regFilter_(user), map = {}, act = activeMonths_();
  regIndex_().forEach(function (r) {
    if (act.indexOf(r.mon) === -1 || !ok(r)) return;
    (map[r.cls] = map[r.cls] || {})[r.sub] = true;
  });
  var list = Object.keys(map).sort(function (a, b) { return a.localeCompare(b, undefined, { numeric: true }); })
    .map(function (c) { return { cls: c, subs: Object.keys(map[c]).sort() }; });
  return { ok: true, view: scopeName_(user), sections: list, months: act };
}

/** Students of the sections this user may see. q = part of a name, father name or roll no. */
function regFind_(user, q) {
  q = String(q || '').trim().toLowerCase();
  if (q.length < 2) return { ok: true, list: [] };
  var ok = regFilter_(user), act = activeMonths_(), clsSet = {};
  regIndex_().forEach(function (r) { if (act.indexOf(r.mon) !== -1 && ok(r)) clsSet[norm_(baseClass_(r.cls))] = true; });
  var bases = Object.keys(clsSet), out = [];
  getMaster_().rows.forEach(function (r) {
    if (out.length >= 40) return;
    var c = norm_(r[2]);
    if (!bases.some(function (b) { return c === b || c.indexOf(b) === 0; })) return;
    if (String(r[0]).toLowerCase().indexOf(q) === -1 && String(r[1]).toLowerCase().indexOf(q) === -1 &&
        String(r[4] || '').toLowerCase().indexOf(q) === -1) return;
    out.push({ roll: r[0], name: r[1], father: r[4] || '', cls: r[2] });
  });
  return { ok: true, list: out };
}

/** req: { from, to, cls ('' or '*' = all), sub ('' = all), roll ('' = everybody) } -> sheets, in order */
function regPlan_(user, req) {
  var ok = regFilter_(user), act = activeMonths_();
  var fi = Math.max(0, SEASON.indexOf(String(req.from || 'SEPT'))), ti = SEASON.indexOf(String(req.to || ''));
  if (ti < 0 || ti > act.length - 1) ti = act.length - 1;
  var cls = String(req.cls || ''), sub = String(req.sub || '').toUpperCase(), roll = String(req.roll || '').trim();
  if (cls === '*') cls = '';
  var list = [];
  regIndex_().forEach(function (r) {
    var mi = SEASON.indexOf(r.mon);
    if (mi < fi || mi > ti) return;
    if (cls && r.cls !== cls) return;
    if (sub && r.sub !== sub) return;
    if (roll && r.json.indexOf('["' + roll + '"') === -1) return;
    if (!ok(r)) return;
    list.push({ name: r.name, mon: r.mon, cls: r.cls, sub: r.sub, teacher: r.teacher });
  });
  list.sort(function (a, b) {
    return a.cls.localeCompare(b.cls, undefined, { numeric: true }) || a.sub.localeCompare(b.sub) ||
           SEASON.indexOf(a.mon) - SEASON.indexOf(b.mon);
  });
  return { ok: true, from: SEASON[fi], to: SEASON[ti], sheets: list };
}

/** req: { names:[sheet names], roll:'' } -> day-by-day data of each sheet */
function regDays_(user, req) {
  var names = (req.names || []).slice(0, REG_MAX_SHEETS), roll = String(req.roll || '').trim();
  var ok = regFilter_(user), meta = {};
  regIndex_().forEach(function (r) { meta[r.name] = r; });
  var out = [];
  names.forEach(function (nm) {
    var m = meta[nm];
    if (!m || !ok(m)) { out.push({ name: nm, missing: true }); return; }
    try { out.push(regSheetDays_(m, roll)); }
    catch (e) { out.push({ name: nm, missing: true, error: String(e && e.message || e) }); }
  });
  return { ok: true, sheets: out };
}

/** One sheet: { name, mon, cls, sub, teacher, m, y, days:[d], held:[n], dn:{d:note}, st:[[roll, name, [cells], off, guestHome]] } */
function regSheetDays_(meta, roll) {
  var sh = bookFor_(meta.mon).getSheetByName(meta.name);
  if (!sh) return { name: meta.name, missing: true };
  var lastR = sh.getLastRow(), lastC = sh.getLastColumn();
  if (lastR < HEADER_ROW) return { name: meta.name, missing: true };
  var v = sh.getRange(1, 1, lastR, lastC).getValues(), hdr = v[HEADER_ROW - 1];
  var rollCol = -1, nameCol = -1, dayCols = {};
  for (var c = 0; c < hdr.length; c++) {
    var t = norm_(hdr[c]);
    if (t === 'ROLLNO' && rollCol < 0) rollCol = c;
    else if (t === 'STUDENTNAME' && nameCol < 0) nameCol = c;
    else if (t !== '' && /^\d+$/.test(t)) { var d = Number(t); if (d >= 1 && d <= 31 && dayCols[d] == null) dayCols[d] = c; }
  }
  if (rollCol < 0) return { name: meta.name, missing: true };
  var days = Object.keys(dayCols).map(Number).sort(function (a, b) { return a - b; });
  var n = lastR - FIRST_ROW + 1;
  var rnotes = n > 0 ? sh.getRange(FIRST_ROW, rollCol + 1, n, 1).getNotes() : [];
  var hnotes = sh.getRange(HEADER_ROW, 1, 1, lastC).getNotes()[0];
  var held = days.map(function (d) {
    var h = Number(v[0][dayCols[d]]) || 0;
    if (!h) for (var i = 0; i < n; i++) {
      var r = v[FIRST_ROW - 1 + i];
      if (!String(r[rollCol]).trim()) continue;
      var x = String(r[dayCols[d]]).trim().toUpperCase();
      if (x === 'P' || x === 'A') { h = 1; break; }
    }
    return h;
  });
  var dn = {};
  days.forEach(function (d) { var t = hnotes[dayCols[d]]; if (t) dn[d] = String(t).slice(0, 60); });
  var st = [];
  for (var i = 0; i < n; i++) {
    var r = v[FIRST_ROW - 1 + i], rl = String(r[rollCol]).trim();
    if (!rl || (roll && rl !== roll)) continue;
    var note = String((rnotes[i] && rnotes[i][0]) || '');
    var guest = note.indexOf('GUEST|') === 0 ? (note.split('|')[1] || '?') : '';
    st.push([rl, nameCol >= 0 ? String(r[nameCol]).trim() : rl,
             days.map(function (d) { return String(r[dayCols[d]]).trim().toUpperCase(); }),
             isOffNote_(note) ? 1 : 0, guest]);
  }
  var my = monthYear_(meta.mon);
  return { name: meta.name, mon: meta.mon, cls: meta.cls, sub: meta.sub, teacher: meta.teacher,
           m: my.m, y: my.y, days: days, held: held, dn: dn, st: st };
}
