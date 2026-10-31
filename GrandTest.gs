/**
 * GRAND TEST module for the CS Attendance backend  (adds the "Grand test" page to the attendance app)
 *
 * Teachers enter grand test marks in the app (also offline). The marks are written straight into the
 * Grand Test workbook, in the same layout the old Grand Test portal used:
 *   one sheet per section (A1, FCS-1 D2, F.A-1 Optional Subjects ...),
 *   row 2 = teacher, row 3 = subject, row 4 = headers (S.No, Enroll#, Student Name, Father Name, Class-Sec,
 *   then "Total Marks" / "Obtained Marks" for every subject), students from row 5.
 * The deadline is kept in Settings!B2 of the Grand Test workbook (the same cell the old portal used).
 * The opening date and the total marks (one total for the whole college, e.g. 40) are kept in a sheet
 * "App Settings" of the Grand Test workbook. Only the admin changes them, from the Grand test page.
 * If no opening date is set, the Grand test opens on 1 January of the session.
 * Every change is written to "WebApp Logs" (Before / After), like the old portal.
 *
 * Install:
 *   1) Attendance Apps Script project: + (Add a file) > Script > name it "GrandTest" > paste this whole file.
 *   2) In doPost, right under the Admissions line you added before, add:
 *        if (/^gt_/.test(String(req.action || ''))) return json_(gtRoute_(user, req));
 *   3) Deploy > Manage deployments > Edit > Version: New version > Deploy.
 *
 * Who sees what: a teacher sees the classes and subjects she has in the Timetable (same as attendance).
 * Admin and users whose Role is ADMIN / PRINCIPAL see every class, the progress list, and may still
 * enter marks after the deadline.
 */

var GT_CFG = {
  BOOK_ID: '1FIT93zzKQnwMMjbaGy0aqrdCMv2QcM1ZTrRBh4SLKM8',   // Grand Test workbook
  SETTINGS_SHEET: 'Settings',
  DEADLINE_CELL: 'B2',
  LOG_SHEET: 'WebApp Logs',
  TITLE: 'JCW GRAND TEST',
  PASS: 0.40,                 // 40% = pass, below is highlighted
  DEFAULT_TOTAL: 40,          // used until the admin sets the total marks
  APP_SETTINGS: 'App Settings',
  LEAVE: ['L', 'LEAVE'],
  ABSENT: ['A', 'AB', 'ABS', 'ABSENT'],
  SKIP_SHEETS: ['settings', 'app settings', 'webapp logs', 'login-logout logs', 'awardlist_log', 'missingtransformstudents']
};

/* ================= Router (called from doPost) ================= */
function gtRoute_(user, req) {
  switch (String(req.action)) {
    case 'gt_list':     return gtList_(user);
    case 'gt_get':      return gtGet_(user, req.cls, req.sub);
    case 'gt_save':     return gtSave_(user, req);
    case 'gt_pdf':      return gtPdf_(user, req.cls, req.sub);
    case 'gt_progress': return gtProgress_(user);
    case 'gt_window':   return gtWindowFor_(user);
    case 'gt_settings': return gtSetSettings_(user, req);
  }
  return { ok: false, error: 'UNKNOWN_ACTION' };
}

/* ================= helpers ================= */
function gtS_(v) { return String(v == null ? '' : v).trim(); }
function gtK_(v) { return gtS_(v).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function gtBook_() { return SpreadsheetApp.openById(GT_CFG.BOOK_ID); }
function gtAll_(user) {
  try { if (isAdmin_(user)) return true; } catch (e) {}
  try { var s = scopeOf_(user); if (s && s.all) return true; } catch (e) {}
  return false;
}
function gtSession_() {
  var d = new Date(), y = d.getFullYear(), m = d.getMonth(), day = d.getDate();
  var start = (m > 7 || (m === 7 && day >= 15)) ? y : y - 1;
  return start + '-' + String(start + 1).slice(-2);
}
function gtIsAbsent_(v) { return GT_CFG.ABSENT.indexOf(gtS_(v).toUpperCase()) >= 0; }
function gtIsLeave_(v) { return GT_CFG.LEAVE.indexOf(gtS_(v).toUpperCase()) >= 0; }
/** 'A' / 'L' / number, or null when the entry is not valid */
function gtMark_(v, total) {
  var s = gtS_(v).toUpperCase();
  if (gtIsAbsent_(s)) return 'A';
  if (gtIsLeave_(s)) return 'L';
  if (/^\d+(\.\d+)?$/.test(s) && Number(s) <= total) return Number(s);
  return null;
}
/** the class + subject this user may enter, or null */
function gtAssignment_(user, cls, sub) {
  var c = gtK_(cls), s = gtK_(sub);
  return getAssignments_(user).filter(function (a) { return gtK_(a.cls) === c && gtK_(a.subject) === s; })[0] || null;
}
function gtDeadline_() {
  try {
    var sh = gtBook_().getSheetByName(GT_CFG.SETTINGS_SHEET);
    if (!sh) return null;
    var v = sh.getRange(GT_CFG.DEADLINE_CELL).getValue();
    var d = Object.prototype.toString.call(v) === '[object Date]' ? v : (v ? new Date(String(v)) : null);
    return d && !isNaN(d.getTime()) ? d.getTime() : null;
  } catch (e) { return null; }
}

/** "App Settings" sheet: Setting | Value rows (Opens, Total marks) */
function gtAppSheet_() {
  var book = gtBook_(), sh = book.getSheetByName(GT_CFG.APP_SETTINGS);
  if (!sh) {
    sh = book.insertSheet(GT_CFG.APP_SETTINGS);
    sh.getRange(1, 1, 3, 3).setValues([['Setting', 'Value', 'Note'],
      ['Opens', '', 'Grand test button opens for teachers (empty = 1 January)'],
      ['Total marks', GT_CFG.DEFAULT_TOTAL, 'Total marks of every subject, whole college']]);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  return sh;
}
function gtApp_() {
  var out = { opens: null, total: GT_CFG.DEFAULT_TOTAL };
  try {
    var sh = gtBook_().getSheetByName(GT_CFG.APP_SETTINGS);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var k = gtK_(r[0]), v = r[1];
      if (k === 'OPENS' && v) { var d = Object.prototype.toString.call(v) === '[object Date]' ? v : new Date(String(v)); if (!isNaN(d.getTime())) out.opens = d.getTime(); }
      if (k === 'TOTALMARKS' && Number(v) > 0) out.total = Number(v);
    });
  } catch (e) {}
  if (!out.opens) {                                   // default: 1 January of this session
    var now = new Date(), y = now.getMonth() >= 7 ? now.getFullYear() + 1 : now.getFullYear();
    out.opens = new Date(y, 0, 1, 0, 0, 0).getTime(); out.defaultOpens = true;
  }
  return out;
}
/** opening date, deadline and total marks, for the Home button and the Grand test page */
function gtWindow_() {
  var app = gtApp_(), dl = gtDeadline_(), now = Date.now();
  return { opens: app.opens, defaultOpens: !!app.defaultOpens, deadline: dl, total: app.total, now: now,
           open: now >= app.opens && (!dl || now <= dl) };
}
function gtWindowFor_(user) { var w = gtWindow_(); w.ok = true; w.all = gtAll_(user); return w; }
/** admin: { total, opens (ms or ''), deadline (ms or '') } */
function gtSetSettings_(user, req) {
  if (!gtAll_(user)) return { ok: false, error: 'ADMIN_ONLY' };
  var total = Number(req.total);
  if (!(total > 0) || total > 1000) return { ok: false, error: 'Total marks must be a number, for example 40.' };
  var opens = req.opens ? new Date(Number(req.opens)) : null, dl = req.deadline ? new Date(Number(req.deadline)) : null;
  if (opens && dl && dl.getTime() <= opens.getTime()) return { ok: false, error: 'The deadline must be after the opening date.' };
  var sh = gtAppSheet_(), rows = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(function (r) { return gtK_(r[0]); });
  var put = function (key, label, val) {
    var i = rows.indexOf(key);
    var row = i >= 0 ? i + 1 : sh.getLastRow() + 1;
    if (i < 0) sh.getRange(row, 1).setValue(label);
    sh.getRange(row, 2).setValue(val);
  };
  put('OPENS', 'Opens', opens || '');
  put('TOTALMARKS', 'Total marks', total);
  var st = gtBook_().getSheetByName(GT_CFG.SETTINGS_SHEET) || gtBook_().insertSheet(GT_CFG.SETTINGS_SHEET);
  st.getRange(GT_CFG.DEADLINE_CELL).setValue(dl || '');      // same cell as the old portal
  SpreadsheetApp.flush();
  var w = gtWindowFor_(user); w.saved = true;
  return w;
}

/** reads a section sheet: header row, enroll / name columns, subject column pairs, student rows */
function gtSheet_(sh) {
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return null;
  var v = sh.getRange(1, 1, lr, lc).getValues();
  var hi = -1;
  for (var i = 0; i < Math.min(v.length, 8); i++) { if (gtK_(v[i][0]) === 'SNO') { hi = i; break; } }
  if (hi < 1) return null;
  var head = v[hi].map(gtK_), subjRow = v[hi - 1], teachRow = hi >= 2 ? v[hi - 2] : [];
  var col = function (names) { for (var k = 0; k < names.length; k++) { var x = head.indexOf(names[k]); if (x >= 0) return x; } return -1; };
  var L = { sh: sh, v: v, hi: hi, lc: lc, enroll: col(['ENROLL']), name: col(['STUDENTNAME', 'NAME']),
            father: col(['FATHERNAME', 'FATHER']), cs: col(['CLASSSEC', 'CLASSSECTION']), subjects: {} };
  var cur = '';
  for (var c = 0; c < lc - 1; c++) {
    if (gtS_(subjRow[c])) cur = gtS_(subjRow[c]);
    if (cur && head[c] === 'TOTALMARKS' && head[c + 1] === 'OBTAINEDMARKS') {
      L.subjects[gtK_(cur)] = { name: cur, total: c, obtained: c + 1, teacher: gtS_(teachRow[c]) };
      cur = '';
    }
  }
  return L.enroll < 0 ? null : L;
}
function gtFindSheet_(book, cls) {
  var want = gtK_(cls);
  return book.getSheets().filter(function (s) { return gtK_(s.getName()) === want; })[0] || null;
}
/** Transform students of this class who take this subject: { roll: [roll, name, father, classSec] } */
function gtTaking_(cls, sub) {
  var M = getMaster_(), subj = gtS_(sub).toUpperCase(), base = gtK_(baseClass_(cls)), isOpt = base !== gtK_(cls), out = {};
  M.rows.forEach(function (r) {
    var c = gtK_(r[2]);
    var ok = isOpt ? (c === base || c.indexOf(base) === 0) : c === base;
    if (ok && r[3].indexOf(subj) !== -1) out[gtS_(r[0]).toUpperCase()] = r;
  });
  return out;
}
/** struck-through names = cancelled admissions in the old portal */
function gtStruck_(L) {
  var n = L.v.length - L.hi - 1, out = {};
  if (n < 1 || L.name < 0) return out;
  try {
    var fl = L.sh.getRange(L.hi + 2, L.name + 1, n, 1).getFontLines();
    for (var i = 0; i < n; i++) if (fl[i][0] === 'line-through') out[i] = true;
  } catch (e) {}
  return out;
}
/** everything the phone needs for one class + subject */
function gtLoad_(cls, sub) {
  var book = gtBook_(), sh = gtFindSheet_(book, cls);
  if (!sh) return { error: 'The Grand Test workbook has no sheet called "' + cls + '". Ask the admin to add it.' };
  var L = gtSheet_(sh);
  if (!L) return { error: 'Sheet "' + sh.getName() + '" has no S.No / Enroll# header row.' };
  var S = L.subjects[gtK_(sub)];
  if (!S) return { error: 'Sheet "' + sh.getName() + '" has no Total Marks / Obtained Marks columns for ' + sub + '. Ask the admin to add them.' };
  var taking = gtTaking_(cls, sub), struck = gtStruck_(L), seen = {}, students = [], totals = {};
  for (var i = L.hi + 1; i < L.v.length; i++) {
    var r = L.v[i], roll = gtS_(r[L.enroll]), key = roll.toUpperCase();
    if (!roll || !taking[key] || seen[key]) continue;
    seen[key] = true;
    var t = gtS_(r[S.total]); if (t) totals[t] = (totals[t] || 0) + 1;
    students.push({ roll: roll, name: gtS_(r[L.name]) || taking[key][1], father: L.father >= 0 ? gtS_(r[L.father]) : taking[key][4],
                    cs: L.cs >= 0 ? gtS_(r[L.cs]) : taking[key][2], mark: gtS_(r[S.obtained]), row: i + 1, off: !!struck[i - L.hi - 1] });
  }
  Object.keys(taking).forEach(function (key) {      // new admissions not in the grand test sheet yet
    if (seen[key]) return;
    var m = taking[key];
    students.push({ roll: m[0], name: m[1], father: m[4] || '', cs: m[2] || '', mark: '', row: 0, missing: true });
  });
  students.sort(function (a, b) { return enrollNum_(a.roll) - enrollNum_(b.roll); });
  var total = Object.keys(totals).sort(function (a, b) { return totals[b] - totals[a]; })[0] || '';
  return { sh: sh, L: L, S: S, students: students, total: total };
}

/* ================= list of my classes ================= */
function gtList_(user) {
  var all = gtAll_(user), items = getAssignments_(user).map(function (a) { return { cls: a.cls, sub: a.subject, teacher: a.teacher }; });
  items.sort(function (x, y) { return x.cls.localeCompare(y.cls, undefined, { numeric: true }) || x.sub.localeCompare(y.sub); });
  var w = gtWindow_();
  return { ok: true, user: String(user.username).toLowerCase(), items: items, all: all, opens: w.opens, deadline: w.deadline, total: w.total,
           now: w.now, session: gtSession_(), title: GT_CFG.TITLE, pass: GT_CFG.PASS };
}

function gtGet_(user, cls, sub) {
  var a = gtAssignment_(user, cls, sub);
  if (!a) return { ok: false, error: 'NOT_ALLOWED: ' + cls + ' ' + sub };
  var x = gtLoad_(cls, sub);
  if (x.error) return { ok: false, error: x.error };
  var w = gtWindow_();
  return { ok: true, cls: cls, sub: x.S.name, teacher: a.teacher, sheet: x.sh.getName(), total: w.total,
           students: x.students.map(function (s) { return { roll: s.roll, name: s.name, father: s.father, cs: s.cs, mark: s.mark, off: !!s.off, missing: !!s.missing }; }),
           opens: w.opens, deadline: w.deadline, now: w.now };
}

/* ================= save ================= */
/** req: { cls, sub, total, marks: { roll: '23' | 'A' }, ts (when it was saved on the phone) } */
function gtSave_(user, req) {
  var a = gtAssignment_(user, req.cls, req.sub);
  if (!a) return { ok: false, error: 'NOT_ALLOWED: ' + req.cls + ' ' + req.sub };
  var w = gtWindow_(), dl = w.deadline, all = gtAll_(user), ts = Math.min(Number(req.ts) || Date.now(), Date.now());
  var fmt = function (t) { return Utilities.formatDate(new Date(t), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a'); };
  if (!all && dl && ts > dl) return { ok: false, closed: true, error: 'The grand test closed on ' + fmt(dl) + '. Ask the admin to extend the deadline.' };
  if (!all && ts < w.opens) return { ok: false, closed: true, error: 'The grand test opens on ' + fmt(w.opens) + '.' };
  var total = w.total;                                   // one total for the whole college, set by the admin
  var marks = req.marks || {};
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var x = gtLoad_(req.cls, req.sub);
    if (x.error) return { ok: false, error: x.error };
    var sh = x.sh, L = x.L, S = x.S, bad = [];
    var byRoll = {}; x.students.forEach(function (s) { byRoll[s.roll.toUpperCase()] = s; });
    Object.keys(marks).forEach(function (roll) {
      var v = gtS_(marks[roll]).toUpperCase();
      if (!byRoll[roll.toUpperCase()]) return;
      if (gtMark_(v, total) === null) bad.push(byRoll[roll.toUpperCase()].name + ' (' + marks[roll] + ')');
    });
    if (bad.length) return { ok: false, error: 'Wrong marks: ' + bad.join(', ') + '. Use a number up to ' + total + ', A for absent or L for leave.' };
    var empty = x.students.filter(function (s) { return !s.off && gtS_(marks[s.roll]) === ''; });
    if (empty.length) return { ok: false, error: empty.length + ' student(s) have no marks: ' + empty.slice(0, 5).map(function (s) { return s.name; }).join(', ') + (empty.length > 5 ? '…' : '') + '. Every student needs marks, A or L.' };

    // 1) new admissions: add a row at the bottom (same formatting and formulas as the last student row)
    var added = 0, lastRow = L.v.length;
    for (var i = L.v.length - 1; i > L.hi; i--) { if (gtS_(L.v[i][L.enroll])) { lastRow = i + 1; break; } }
    var maxSno = 0; for (var j = L.hi + 1; j < L.v.length; j++) { var sn = Number(L.v[j][0]); if (sn > maxSno) maxSno = sn; }
    x.students.forEach(function (s) {
      if (!s.missing || marks[s.roll] === undefined) return;
      var row = lastRow + 1;
      if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 5);
      if (lastRow > L.hi + 1) {
        sh.getRange(lastRow, 1, 1, L.lc).copyTo(sh.getRange(row, 1, 1, L.lc), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        var f = sh.getRange(lastRow, 1, 1, L.lc).getFormulasR1C1()[0];
        f.forEach(function (fx, c) { if (fx) sh.getRange(row, c + 1).setFormulaR1C1(fx); });
      }
      var line = {}; line[0] = ++maxSno; line[L.enroll] = s.roll;
      if (L.name >= 0) line[L.name] = s.name; if (L.father >= 0) line[L.father] = s.father; if (L.cs >= 0) line[L.cs] = s.cs;
      Object.keys(line).forEach(function (c) { sh.getRange(row, Number(c) + 1).setValue(line[c]); });
      s.row = row; lastRow = row; added++;
    });

    // 2) marks: Total + Obtained of this subject, only for this subject's students
    var rows = x.students.filter(function (s) { return s.row && marks[s.roll] !== undefined; });
    var changes = [], when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy hh:mm a');
    rows.forEach(function (s) {
      var val = gtMark_(marks[s.roll], total);
      var was = s.missing ? '' : s.mark;
      sh.getRange(s.row, S.total + 1, 1, 2).setValues([[total, val]]);
      if (String(was).toUpperCase() !== String(val).toUpperCase()) changes.push([s, was, val]);
    });
    SpreadsheetApp.flush();

    // 3) log, like the old portal ("WebApp Logs")
    if (changes.length) {
      try {
        var book = gtBook_(), lg = book.getSheetByName(GT_CFG.LOG_SHEET);
        if (!lg) { lg = book.insertSheet(GT_CFG.LOG_SHEET); lg.appendRow(['Timestamp', 'Username', 'Enroll#', 'Student', 'Section', 'Subject', 'Teacher', 'Cell Link', 'Before', 'After']); }
        var base = book.getUrl() + '#gid=' + sh.getSheetId() + '&range=';
        var out = changes.map(function (c) {
          var a1 = sh.getRange(c[0].row, S.obtained + 1).getA1Notation();
          return [when, user.username, c[0].roll, c[0].name, sh.getName(), S.name, a.teacher,
                  '=HYPERLINK("' + base + a1 + '","' + sh.getName().replace(/"/g, '') + '!' + a1 + '")', c[1], c[2]];
        });
        lg.getRange(lg.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
      } catch (e) { /* logging never stops a save */ }
    }
    CacheService.getScriptCache().remove('gt_progress_v1');
    return { ok: true, saved: rows.length, changed: changes.length, added: added, late: !!(dl && ts > dl) };
  } finally { lock.releaseLock(); }
}

/* ================= award list PDF (same look as the old portal) ================= */
function gtPdf_(user, cls, sub) {
  var a = gtAssignment_(user, cls, sub);
  if (!a) return { ok: false, error: 'NOT_ALLOWED: ' + cls + ' ' + sub };
  var x = gtLoad_(cls, sub);
  if (x.error) return { ok: false, error: x.error };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var title = function (s) { return String(s).toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); }); };
  var list = x.students.filter(function (s) { return !s.missing; });
  var total = gtWindow_().total, hi = 0, pres = 0, abs = 0, nr = 0, pass = 0, fail = 0, lv = 0;
  list.forEach(function (s) {
    if (s.mark === '') nr++;
    else if (gtIsAbsent_(s.mark)) abs++;
    else if (gtIsLeave_(s.mark)) lv++;
    else if (!isNaN(Number(s.mark))) { pres++; var m = Number(s.mark); if (m > hi) hi = m; if (total && m >= total * GT_CFG.PASS) pass++; else if (total) fail++; }
  });
  var rows = list.map(function (s, i) {
    var m = s.mark, cls2 = '', show = m;
    if (m === '') { show = 'N/R'; cls2 = 'nr'; }
    else if (gtIsAbsent_(m)) { show = 'ABSENT'; cls2 = 'ab'; }
    else if (gtIsLeave_(m)) { show = 'LEAVE'; cls2 = 'nr'; }
    else if (total && Number(m) < total * GT_CFG.PASS) cls2 = 'low';
    else if (Number(m) === hi && hi) cls2 = 'top';
    return '<tr><td>' + (i + 1) + '</td><td>' + esc(s.roll) + '</td><td class="l">' + esc(title(s.name)) + (s.off ? ' (cancelled)' : '') + '</td><td class="l">' + esc(title(s.father)) +
      '</td><td>' + esc(s.cs) + '</td><td>' + esc(total || '') + '</td><td class="' + cls2 + '">' + esc(show) + '</td></tr>';
  }).join('');
  var html = '<html><head><style>@page{size:A4 portrait;margin:10mm}body{font-family:Arial,sans-serif;font-size:10px}' +
    'h2,h4{text-align:center;margin:4px 0}h2{color:#6B3423}h4{color:#8A4A35}table{border-collapse:collapse;width:100%;margin-top:6px}th,td{border:1px solid #8A4A35;padding:3px;text-align:center;font-size:12px}' +
    'td.l{text-align:left}th{background:#6B3423;color:#fff}.ab{background:#ffd6d6;color:#c00;font-weight:bold}.nr{background:#e0e0e0;color:#060;font-weight:bold}' +
    '.top{background:#d4f8d4;font-weight:bold}.low{background:#fff0b3;color:#d9534f;font-weight:bold}.sum{margin-top:10px;font-weight:bold;font-size:12px}.sig{margin-top:28px;text-align:right;font-weight:bold;font-size:11px}</style></head><body>' +
    '<h2>' + esc(GT_CFG.TITLE) + ' AWARD LIST</h2><h4>Section: ' + esc(x.sh.getName()) + ' | Subject: ' + esc(x.S.name) + ' | Teacher: ' + esc(a.teacher) + ' | Session: ' + gtSession_() + '</h4>' +
    '<table><tr><th>S.No</th><th>Enroll#</th><th>Name</th><th>Father Name</th><th>Class-Sec</th><th>Total</th><th>Obtained</th></tr>' + rows + '</table>' +
    '<div class="sum">Total students: ' + list.length + ' | Present: ' + pres + ' | Absent: ' + abs + ' | Leave: ' + lv + ' | N/R: ' + nr + ' | Pass: ' + pass + ' | Fail: ' + fail + (hi ? ' | Highest: ' + hi : '') + '</div>' +
    '<div class="sig">___________________________<br>' + esc(a.teacher) + '</div></body></html>';
  var name = 'G.T_' + x.sh.getName() + '_' + x.S.name + '_' + a.teacher + '_' + gtSession_() + '.pdf';
  var blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(name);
  return { ok: true, name: name, pdf: Utilities.base64Encode(blob.getBytes()) };
}

/* ================= progress for admin / principal: which subjects still have missing marks ================= */
function gtProgress_(user) {
  if (!gtAll_(user)) return { ok: false, error: 'ADMIN_ONLY' };
  var cache = CacheService.getScriptCache(), hit = cache.get('gt_progress_v1');
  if (hit) return JSON.parse(hit);
  var book = gtBook_(), sheets = {}, M = getMaster_(), out = [];
  book.getSheets().forEach(function (s) { if (GT_CFG.SKIP_SHEETS.indexOf(s.getName().toLowerCase()) < 0) sheets[gtK_(s.getName())] = s; });
  var read = {};
  getAssignments_(user).forEach(function (a) {
    var it = { cls: a.cls, sub: a.subject, teacher: a.teacher, students: 0, entered: 0, status: '' };
    var sh = sheets[gtK_(a.cls)];
    if (!sh) { it.status = 'nosheet'; out.push(it); return; }
    var L = read[sh.getName()] || (read[sh.getName()] = gtSheet_(sh) || { bad: true });
    var S = L.bad ? null : L.subjects[gtK_(a.subject)];
    if (!S) { it.status = 'nocol'; out.push(it); return; }
    var taking = {}, subj = a.subject.toUpperCase(), base = gtK_(baseClass_(a.cls)), isOpt = base !== gtK_(a.cls);
    M.rows.forEach(function (r) { var c = gtK_(r[2]); if ((isOpt ? (c === base || c.indexOf(base) === 0) : c === base) && r[3].indexOf(subj) !== -1) taking[gtS_(r[0]).toUpperCase()] = true; });
    if (!L.struck) L.struck = gtStruck_(L);           // cancelled (struck-through) students need no marks
    var seen = {};
    for (var i = L.hi + 1; i < L.v.length; i++) {
      var k = gtS_(L.v[i][L.enroll]).toUpperCase();
      if (!k || !taking[k] || seen[k]) continue;
      seen[k] = true;
      if (L.struck[i - L.hi - 1]) { delete taking[k]; continue; }
      if (gtS_(L.v[i][S.obtained]) !== '') it.entered++;
    }
    it.students = Object.keys(taking).length;
    it.status = it.students && it.entered >= it.students ? 'done' : it.entered ? 'partial' : 'none';
    out.push(it);
  });
  var res = { ok: true, items: out, at: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'), deadline: gtDeadline_() };
  try { cache.put('gt_progress_v1', JSON.stringify(res), 180); } catch (e) {}
  return res;
}
