/**
 * DETENTION EXAM module for the CS Attendance backend  (needs GrandTest.gs in the same project)
 *
 * Teachers enter detention marks in the app (also offline). Marks go straight into the Detention workbook,
 * same layout as the old Detention portal: one sheet per section; row 2 = teacher, row 3 = subject,
 * row 4 = S.No, Enroll#, Student Name, Father Name, Class-Sec, then THEORY / PRACTICAL / TOTAL for every subject.
 * Theory and practical are typed (a number, A for absent, L for leave); TOTAL is worked out.
 * Practical is only for practical subjects (DT_CFG.PRACTICAL).
 *
 * Window: teachers can enter marks from "Opens" (default 1 February) until the deadline in Settings!B2.
 * Opens, subject total and practical marks are kept in the sheet "App Settings" of the Detention workbook;
 * the admin changes them on the Detention page (Settings tab).
 *
 * Install:
 *   1) + > Script > name it "Detention" > paste this file.
 *   2) doPost, under the gt_ line:
 *        if (/^dt_/.test(String(req.action || ''))) return json_(dtRoute_(user, req));
 *   3) Deploy > Manage deployments > Edit > New version > Deploy.
 */
var DT_CFG = {
  BOOK_ID: '1pCfG_bSixEyiwtH2qVRA01ilfF6F-2P5b1C1HSUXQu8',   // Detention workbook
  SETTINGS_SHEET: 'Settings',
  DEADLINE_CELL: 'B2',
  APP_SETTINGS: 'App Settings',
  LOG_SHEET: 'WebApp Logs',
  TITLE: 'JCW DETENTION AWARDLIST',
  PASS: 0.33,
  // Marks of each subject: [theory, practical]. Every subject not listed: theory 100, no practical.
  SCHEME: {
    'BOTANY': [85, 15], 'CHEMISTRY': [85, 15], 'PHYSICS': [85, 15], 'ZOOLOGY': [85, 15],
    'COMPUTER SCIENCE': [75, 25], 'GEOGRAPHY': [75, 25],
    'PAK-STUDIES': [50, 0], 'ISLAMIYAT': [50, 0], 'MUTALIA QURAN': [50, 0]
  },
  DEFAULT_SCHEME: [100, 0],
  OPENS_MONTH: 1,             // February (0 = January)
  ALIASES: { PASHTO: 'PUSHTO', ISLAMIYATELECTIVE: 'ISLAMIYATELE' },
  SKIP: /combine|report|log|settings|missing|summary|topper|top3|contacts|dashboard|rooms|courses|timetable|link sheet|attendance|ethics|names|unassigned|backup|copy of/i
};

function dtRoute_(user, req) {
  switch (String(req.action)) {
    case 'dt_list':     return dtList_(user);
    case 'dt_get':      return dtGet_(user, req.cls, req.sub);
    case 'dt_save':     return dtSave_(user, req);
    case 'dt_pdf':      return dtPdf_(user, req.cls, req.sub);
    case 'dt_progress': return dtProgress_(user);
    case 'dt_window':   var w = dtWindow_(); w.ok = true; w.all = gtAll_(user); return w;
    case 'dt_settings': return dtSetSettings_(user, req);
  }
  return { ok: false, error: 'UNKNOWN_ACTION' };
}

/* ================= helpers ================= */
function dtBook_() { return SpreadsheetApp.openById(DT_CFG.BOOK_ID); }
/** { th, pr, st } marks of a subject (exact name: ISLAMIYAT ELECTIVE is not ISLAMIYAT) */
function dtLimits_(sub) {
  var k = gtK_(sub), hit = DT_CFG.DEFAULT_SCHEME;
  Object.keys(DT_CFG.SCHEME).forEach(function (n) { if (gtK_(n) === k) hit = DT_CFG.SCHEME[n]; });
  return { th: hit[0], pr: hit[1], st: hit[0] + hit[1] };
}
function dtIsPractical_(sub) { return dtLimits_(sub).pr > 0; }
function dtScheme_() { var o = {}; Object.keys(DT_CFG.SCHEME).forEach(function (n) { o[n] = DT_CFG.SCHEME[n]; }); o['Other subjects'] = DT_CFG.DEFAULT_SCHEME; return o; }
/** 'A' / 'L' / number / '' ; null when not valid */
function dtVal_(v, max) {
  var s = gtS_(v).toUpperCase();
  if (s === '') return '';
  if (['A', 'AB', 'ABS', 'ABSENT'].indexOf(s) >= 0) return 'A';
  if (['L', 'LEAVE'].indexOf(s) >= 0) return 'L';
  if (/^\d+(\.\d+)?$/.test(s) && Number(s) <= max) return Number(s);
  return null;
}
/** total like the old portal: both absent / leave -> 'A'; otherwise the sum (A and L count 0) */
function dtTotal_(th, pr, practical) {
  var abs = function (x) { return x === 'A' || x === 'L'; };
  if (th === '' && pr === '') return '';
  if (!practical) return th === '' ? '' : abs(th) ? 'A' : th;
  if (abs(th) && (pr === '' || abs(pr))) return 'A';
  if (th === '' && abs(pr)) return 'A';
  return (abs(th) || th === '' ? 0 : th) + (abs(pr) || pr === '' ? 0 : pr);
}
function dtApp_() {
  var out = { opens: null };
  try {
    var sh = dtBook_().getSheetByName(DT_CFG.APP_SETTINGS);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var k = gtK_(r[0]), v = r[1];
      if (k === 'OPENS' && v) { var d = Object.prototype.toString.call(v) === '[object Date]' ? v : new Date(String(v)); if (!isNaN(d.getTime())) out.opens = d.getTime(); }
    });
  } catch (e) {}
  if (!out.opens) {                                       // default: 1 February of this session
    var now = new Date(), y = now.getMonth() >= 7 ? now.getFullYear() + 1 : now.getFullYear();
    out.opens = new Date(y, DT_CFG.OPENS_MONTH, 1, 0, 0, 0).getTime(); out.defaultOpens = true;
  }
  return out;
}
function dtDeadline_() {
  try {
    var sh = dtBook_().getSheetByName(DT_CFG.SETTINGS_SHEET); if (!sh) return null;
    var v = sh.getRange(DT_CFG.DEADLINE_CELL).getValue();
    var d = Object.prototype.toString.call(v) === '[object Date]' ? v : (v ? new Date(String(v)) : null);
    return d && !isNaN(d.getTime()) ? d.getTime() : null;
  } catch (e) { return null; }
}
function dtWindow_() {
  var a = dtApp_(), dl = dtDeadline_(), now = Date.now();
  return { opens: a.opens, defaultOpens: !!a.defaultOpens, deadline: dl, scheme: dtScheme_(),
           now: now, open: now >= a.opens && (!dl || now <= dl) };
}
function dtSetSettings_(user, req) {
  if (!gtAll_(user)) return { ok: false, error: 'ADMIN_ONLY' };
  var opens = req.opens ? new Date(Number(req.opens)) : null, dl = req.deadline ? new Date(Number(req.deadline)) : null;
  if (opens && dl && dl.getTime() <= opens.getTime()) return { ok: false, error: 'The deadline must be after the opening date.' };
  var book = dtBook_(), sh = book.getSheetByName(DT_CFG.APP_SETTINGS);
  if (!sh) { sh = book.insertSheet(DT_CFG.APP_SETTINGS); sh.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Note']]).setFontWeight('bold'); }
  var rows = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(function (r) { return gtK_(r[0]); });
  var put = function (key, label, val, note) {
    var i = rows.indexOf(key), row = i >= 0 ? i + 1 : sh.getLastRow() + 1;
    if (i < 0) { sh.getRange(row, 1).setValue(label); sh.getRange(row, 3).setValue(note); rows[row - 1] = key; }
    sh.getRange(row, 2).setValue(val);
  };
  put('OPENS', 'Opens', opens || '', 'Detention button opens for teachers (empty = 1 February)');
  var s2 = book.getSheetByName(DT_CFG.SETTINGS_SHEET) || book.insertSheet(DT_CFG.SETTINGS_SHEET);
  s2.getRange(DT_CFG.DEADLINE_CELL).setValue(dl || '');
  SpreadsheetApp.flush();
  var w = dtWindow_(); w.ok = true; w.all = true; w.saved = true;
  return w;
}

/** section sheet: header row (S.No), subject -> { theory, practical, total } columns */
function dtSheet_(sh) {
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return null;
  var v = sh.getRange(1, 1, lr, lc).getValues(), hi = -1;
  for (var i = 0; i < Math.min(v.length, 8); i++) { if (gtK_(v[i][0]) === 'SNO') { hi = i; break; } }
  if (hi < 1) return null;
  var head = v[hi].map(gtK_), subjRow = v[hi - 1], teachRow = hi >= 2 ? v[hi - 2] : [];
  var col = function (n) { for (var k = 0; k < n.length; k++) { var x = head.indexOf(n[k]); if (x >= 0) return x; } return -1; };
  var L = { sh: sh, v: v, hi: hi, lc: lc, sno: 0, enroll: col(['ENROLL']), name: col(['STUDENTNAME', 'NAME']), father: col(['FATHERNAME', 'FATHER']),
            cs: col(['CLASSSEC', 'CLASSSECTION']), subjects: {} };
  for (var c = 0; c < lc - 2; c++) {
    if (head[c] === 'THEORY' && head[c + 1] === 'PRACTICAL' && head[c + 2] === 'TOTAL' && gtS_(subjRow[c])) {
      L.subjects[gtK_(subjRow[c])] = { name: gtS_(subjRow[c]), theory: c, practical: c + 1, total: c + 2, teacher: gtS_(teachRow[c]) };
    }
  }
  return L.enroll < 0 ? null : L;
}
/** subject columns, also when the sheet writes the name a little differently (PUSHTO / PASHTO, ISLAMIYAT ELE ...) */
function dtSubject_(L, sub) {
  var k = gtK_(sub); if (L.subjects[k]) return L.subjects[k];
  var a = DT_CFG.ALIASES[k]; if (a && L.subjects[a]) return L.subjects[a];
  var keys = Object.keys(L.subjects).filter(function (x) { return x.indexOf(k) === 0 || k.indexOf(x) === 0; });
  return keys.length === 1 ? L.subjects[keys[0]] : null;
}
function dtLoad_(cls, sub) {
  var book = dtBook_(), sh = gtFindSheet_(book, cls);
  if (!sh) return { error: 'The Detention workbook has no sheet called "' + cls + '". Ask the admin to add it.' };
  var L = dtSheet_(sh);
  if (!L) return { error: 'Sheet "' + sh.getName() + '" has no S.No / Enroll# header row.' };
  var S = dtSubject_(L, sub);
  if (!S) return { error: 'Sheet "' + sh.getName() + '" has no THEORY / PRACTICAL / TOTAL columns for ' + sub + '. Ask the admin to add them.' };
  var taking = gtTaking_(cls, sub), struck = gtStruck_(L), seen = {}, students = [];
  for (var i = L.hi + 1; i < L.v.length; i++) {
    var r = L.v[i], roll = gtS_(r[L.enroll]), key = roll.toUpperCase();
    if (!roll || !taking[key] || seen[key]) continue;
    seen[key] = true;
    students.push({ roll: roll, name: gtS_(r[L.name]) || taking[key][1], father: L.father >= 0 ? gtS_(r[L.father]) : taking[key][4], cs: L.cs >= 0 ? gtS_(r[L.cs]) : taking[key][2],
                    th: gtS_(r[S.theory]), pr: gtS_(r[S.practical]), tot: gtS_(r[S.total]), row: i + 1, off: !!struck[i - L.hi - 1] });
  }
  Object.keys(taking).forEach(function (key) {
    if (seen[key]) return;
    var m = taking[key];
    students.push({ roll: m[0], name: m[1], father: m[4] || '', cs: m[2] || '', th: '', pr: '', tot: '', row: 0, missing: true });
  });
  students.sort(function (a, b) { return enrollNum_(a.roll) - enrollNum_(b.roll); });
  return { sh: sh, L: L, S: S, students: students };
}

/* ================= actions ================= */
function dtList_(user) {
  var items = getAssignments_(user).map(function (a) { return { cls: a.cls, sub: a.subject, teacher: a.teacher, practical: dtIsPractical_(a.subject) }; });
  items.sort(function (x, y) { return x.cls.localeCompare(y.cls, undefined, { numeric: true }) || x.sub.localeCompare(y.sub); });
  var w = dtWindow_();
  return { ok: true, user: String(user.username).toLowerCase(), items: items, all: gtAll_(user), opens: w.opens, deadline: w.deadline,
           scheme: w.scheme, now: w.now, session: gtSession_(), title: DT_CFG.TITLE, pass: DT_CFG.PASS };
}
function dtGet_(user, cls, sub) {
  var a = gtAssignment_(user, cls, sub);
  if (!a) return { ok: false, error: 'NOT_ALLOWED: ' + cls + ' ' + sub };
  var x = dtLoad_(cls, sub); if (x.error) return { ok: false, error: x.error };
  var w = dtWindow_();
  var lm = dtLimits_(a.subject);
  return { ok: true, cls: cls, sub: a.subject, teacher: a.teacher, sheet: x.sh.getName(), practical: lm.pr > 0,
           thMax: lm.th, prMax: lm.pr, subjectTotal: lm.st, opens: w.opens, deadline: w.deadline, now: w.now,
           students: x.students.map(function (s) { return { roll: s.roll, name: s.name, father: s.father, cs: s.cs, th: s.th, pr: s.pr, tot: s.tot, off: !!s.off, missing: !!s.missing }; }) };
}
/** req: { cls, sub, marks: { roll: { th, pr } }, ts } */
function dtSave_(user, req) {
  var a = gtAssignment_(user, req.cls, req.sub);
  if (!a) return { ok: false, error: 'NOT_ALLOWED: ' + req.cls + ' ' + req.sub };
  var w = dtWindow_(), all = gtAll_(user), ts = Math.min(Number(req.ts) || Date.now(), Date.now());
  var fmt = function (t) { return Utilities.formatDate(new Date(t), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a'); };
  if (!all && w.deadline && ts > w.deadline) return { ok: false, closed: true, error: 'The detention exam closed on ' + fmt(w.deadline) + '. Ask the admin to extend the deadline.' };
  if (!all && ts < w.opens) return { ok: false, closed: true, error: 'The detention exam opens on ' + fmt(w.opens) + '.' };
  var lm = dtLimits_(a.subject), practical = lm.pr > 0, thMax = lm.th, prMax = lm.pr;
  var marks = req.marks || {};
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var x = dtLoad_(req.cls, req.sub); if (x.error) return { ok: false, error: x.error };
    var sh = x.sh, L = x.L, S = x.S, bad = [], empty = [], vals = {};
    x.students.forEach(function (s) {
      if (s.off) return;
      var m = marks[s.roll] || {}, th = dtVal_(m.th, thMax), pr = practical ? dtVal_(m.pr, prMax) : '';
      if (th === null || pr === null) { bad.push(s.name); return; }
      if (th === '' || (practical && pr === '')) { empty.push(s.name); return; }
      vals[s.roll] = { th: th, pr: pr, tot: dtTotal_(th, pr, practical) };
    });
    if (bad.length) return { ok: false, error: 'Wrong marks for: ' + bad.slice(0, 5).join(', ') + (bad.length > 5 ? '…' : '') + '. Theory up to ' + thMax + (practical ? ', practical up to ' + prMax : '') + ', A for absent or L for leave.' };
    if (empty.length) return { ok: false, error: empty.length + ' student(s) have no marks: ' + empty.slice(0, 5).join(', ') + (empty.length > 5 ? '…' : '') + '.' };
    var added = 0, changes = [];
    x.students.forEach(function (s) {
      var v = vals[s.roll]; if (!v) return;
      if (s.missing) { s.row = bkAppend_(sh, L, [s.roll, s.name, s.cs, [], s.father]); added++; }
      var was = s.missing ? '' : s.tot;
      sh.getRange(s.row, S.theory + 1, 1, 3).setValues([[v.th, practical ? v.pr : (s.pr || ''), v.tot]]);
      if (String(was).toUpperCase() !== String(v.tot).toUpperCase() || String(s.th).toUpperCase() !== String(v.th).toUpperCase()) changes.push([s, s.th + (practical ? ' + ' + s.pr : ''), v.th + (practical ? ' + ' + v.pr : '') + ' = ' + v.tot]);
    });
    SpreadsheetApp.flush();
    if (changes.length) {
      try {
        var book = dtBook_(), lg = book.getSheetByName(DT_CFG.LOG_SHEET);
        if (!lg) { lg = book.insertSheet(DT_CFG.LOG_SHEET); lg.appendRow(['Timestamp', 'User', 'Enroll#', 'Student', 'Section', 'Subject', 'Teacher', 'Cell', 'Before', 'After']); }
        var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy hh:mm a');
        var out = changes.map(function (c) { return [when, user.username, c[0].roll, c[0].name, sh.getName(), S.name, a.teacher, sh.getName() + '!' + sh.getRange(c[0].row, S.total + 1).getA1Notation(), c[1], c[2]]; });
        lg.getRange(lg.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
      } catch (e) {}
    }
    try { CacheService.getScriptCache().remove('dt_progress_v1'); } catch (e) {}
    return { ok: true, saved: Object.keys(vals).length, changed: changes.length, added: added, late: !!(w.deadline && ts > w.deadline) };
  } finally { lock.releaseLock(); }
}
function dtPdf_(user, cls, sub) {
  var a = gtAssignment_(user, cls, sub);
  if (!a) return { ok: false, error: 'NOT_ALLOWED: ' + cls + ' ' + sub };
  var x = dtLoad_(cls, sub); if (x.error) return { ok: false, error: x.error };
  var lm = dtLimits_(a.subject), practical = lm.pr > 0, max = lm.st;
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var title = function (s) { return String(s).toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); }); };
  var list = x.students.filter(function (s) { return !s.missing; }), hi = 0, pres = 0, abs = 0, nr = 0, pass = 0, fail = 0;
  list.forEach(function (s) {
    var t = s.tot;
    if (t === '') nr++; else if (/^[AL]$/i.test(t)) abs++;
    else if (!isNaN(Number(t))) { pres++; var m = Number(t); if (m > hi) hi = m; if (m >= max * DT_CFG.PASS) pass++; else fail++; }
  });
  var rows = list.map(function (s, i) {
    var t = s.tot, c = '', show = t;
    if (t === '') { show = 'N/R'; c = 'nr'; } else if (/^A$/i.test(t)) { show = 'ABSENT'; c = 'ab'; } else if (/^L$/i.test(t)) { show = 'LEAVE'; c = 'nr'; }
    else if (Number(t) < max * DT_CFG.PASS) c = 'low'; else if (Number(t) === hi && hi) c = 'top';
    return '<tr><td>' + (i + 1) + '</td><td>' + esc(s.roll) + '</td><td class="l">' + esc(title(s.name)) + (s.off ? ' (cancelled)' : '') + '</td><td class="l">' + esc(title(s.father)) +
      '</td><td>' + esc(s.th) + '</td>' + (practical ? '<td>' + esc(s.pr) + '</td>' : '') + '<td class="' + c + '">' + esc(show) + '</td></tr>';
  }).join('');
  var html = '<html><head><style>@page{size:A4 portrait;margin:10mm}body{font-family:Arial,sans-serif;font-size:10px}h2,h4{text-align:center;margin:4px 0}h2{color:#6B3423}h4{color:#8A4A35}' +
    'table{border-collapse:collapse;width:100%;margin-top:6px}th,td{border:1px solid #8A4A35;padding:3px;text-align:center;font-size:12px}td.l{text-align:left}th{background:#6B3423;color:#fff}' +
    '.ab{background:#ffd6d6;color:#c00;font-weight:bold}.nr{background:#e0e0e0;color:#060;font-weight:bold}.top{background:#d4f8d4;font-weight:bold}.low{background:#fff0b3;color:#d9534f;font-weight:bold}' +
    '.sum{margin-top:10px;font-weight:bold;font-size:12px}.sig{margin-top:28px;text-align:right;font-weight:bold;font-size:11px}</style></head><body>' +
    '<h2>' + esc(DT_CFG.TITLE) + ' - ' + gtSession_() + '</h2><h4>Section: ' + esc(x.sh.getName()) + ' | Subject: ' + esc(x.S.name) + ' | Teacher: ' + esc(a.teacher) + ' | Out of ' + max + '</h4>' +
    '<table><tr><th>S.No</th><th>Enroll#</th><th>Name</th><th>Father Name</th><th>Theory</th>' + (practical ? '<th>Practical</th>' : '') + '<th>Total</th></tr>' + rows + '</table>' +
    '<div class="sum">Total students: ' + list.length + ' | Present: ' + pres + ' | Absent / Leave: ' + abs + ' | N/R: ' + nr + ' | Pass: ' + pass + ' | Fail: ' + fail + (hi ? ' | Highest: ' + hi : '') + '</div>' +
    '<div class="sig">___________________________<br>' + esc(a.teacher) + '</div></body></html>';
  var name = 'Detention_' + x.sh.getName() + '_' + x.S.name + '_' + a.teacher + '_' + gtSession_() + '.pdf';
  var blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(name);
  return { ok: true, name: name, pdf: Utilities.base64Encode(blob.getBytes()) };
}
function dtProgress_(user) {
  if (!gtAll_(user)) return { ok: false, error: 'ADMIN_ONLY' };
  var cache = CacheService.getScriptCache(), hit = cache.get('dt_progress_v1');
  if (hit) return JSON.parse(hit);
  var book = dtBook_(), sheets = {}, read = {}, out = [];
  book.getSheets().forEach(function (s) { if (!DT_CFG.SKIP.test(s.getName())) sheets[gtK_(s.getName())] = s; });
  getAssignments_(user).forEach(function (a) {
    var it = { cls: a.cls, sub: a.subject, teacher: a.teacher, students: 0, entered: 0, status: '' };
    var sh = sheets[gtK_(a.cls)];
    if (!sh) { it.status = 'nosheet'; out.push(it); return; }
    var L = read[sh.getName()] || (read[sh.getName()] = dtSheet_(sh) || { bad: true });
    var S = L.bad ? null : dtSubject_(L, a.subject);
    if (!S) { it.status = 'nocol'; out.push(it); return; }
    if (!L.struck) L.struck = gtStruck_(L);
    var taking = gtTaking_(a.cls, a.subject), seen = {};
    for (var i = L.hi + 1; i < L.v.length; i++) {
      var k = gtS_(L.v[i][L.enroll]).toUpperCase();
      if (!k || !taking[k] || seen[k]) continue;
      seen[k] = true;
      if (L.struck[i - L.hi - 1]) { delete taking[k]; continue; }
      if (gtS_(L.v[i][S.total]) !== '' || gtS_(L.v[i][S.theory]) !== '') it.entered++;
    }
    it.students = Object.keys(taking).length;
    it.status = it.students && it.entered >= it.students ? 'done' : it.entered ? 'partial' : 'none';
    out.push(it);
  });
  var res = { ok: true, items: out, at: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm') };
  try { cache.put('dt_progress_v1', JSON.stringify(res), 180); } catch (e) {}
  return res;
}
