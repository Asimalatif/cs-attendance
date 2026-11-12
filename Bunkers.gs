/**
 * Bunkers.gs — Daily bunkers report (page bunkers.html).
 *
 * A "bunker" is a student who was PRESENT in at least one lecture on a day but ABSENT (A) or BUNK (B)
 * in another lecture of the same day. Students a teacher marked "B" are listed too, even with no P that day.
 * Leave (L) is never counted as bunking.
 *
 * It reads the hidden "Period Data" sheet of the month workbook (the real mark of every period of every
 * subject), so nothing new has to be marked by the teachers. Nothing is written to the attendance sheets.
 *
 * Who can open it: admin / principal (whole college), Chief Proctors (whole college),
 *                  HOD (only lectures taught by teachers of her department).
 *
 * INSTALL:
 *   a) + > Script > name it "Bunkers" > paste this file.
 *   b) doPost, next to the rpt_ line:
 *        if (/^bnk_/.test(String(req.action || ''))) return json_(bnkRoute_(user, req));
 *   c) Deploy > Manage deployments > Edit > New version > Deploy.
 *   d) Optional: run bnkSetupDailyLog once — every day at about 4 pm the day's bunkers are also saved
 *      in the sheet "Bunkers Log" of the Users workbook (history that stays even after the month ends).
 */
var BNK_LOG_SHEET = 'Bunkers Log';
var BNK_LOG_HEAD = ['Date', 'Roll No', 'Student', 'Section', 'Pattern', 'Attended', 'Missed', 'Missed lectures', 'Attended lectures'];
var BNK_TOK_ = { 8: 'SEPT', 9: 'OCT', 10: 'NOV', 11: 'DEC', 0: 'JAN' };

function bnkRoute_(user, req) {
  try {
    var who = bnkWho_(user);
    if (!who) return { ok: false, error: 'Only the admin, principal, HOD or Chief Proctor can see the bunkers report.' };
    var a = String(req.action || '');
    if (a === 'bnk_day')   return bnkDay_(who, String(req.date || ''));
    if (a === 'bnk_month') return bnkMonth_(who, String(req.mon || ''));
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

/** { all: true } or { all: false, dept: 'COMPUTER SCIENCE' } or null */
function bnkWho_(user) {
  var s = scopeOf_(user);
  if (s && s.all) return { all: true, label: 'College' };
  var cp = false; try { cp = isCP_(user); } catch (e) {}
  if (cp) return { all: true, label: 'College (Chief Proctor)' };
  if (s) return { all: false, dept: s.dept, label: s.dept || 'Department' };
  return null;
}

/** "2026-11-12" -> { tok:'NOV', day:12, date:Date } (only Sept to Jan have attendance workbooks) */
function bnkDate_(txt) {
  var m = String(txt || '').match(/^(\d{4})-(\d{2})-(\d{2})$/), d = m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date();
  var tok = BNK_TOK_[d.getMonth()];
  if (!tok) throw new Error('Attendance is only kept from September to January.');
  return { tok: tok, day: d.getDate(), date: d, iso: Utilities.formatDate(d, TZ_(), 'yyyy-MM-dd') };
}

/** Everything one month needs: period marks of every sheet + sheet info + student names. */
function bnkLoad_(tok, who) {
  var book = bookFor_(tok), sh = book.getSheetByName(PD_SHEET), sheets = [];
  if (!sh || sh.getLastRow() < 2) return { sheets: [], names: {} };
  var idx = indexRows_(tok), depts = who.all ? null : teacherDepts_();
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    var name = String(r[0] || ''), raw;
    if (!name) return;
    try { raw = JSON.parse(r[1] || '{}'); } catch (e) { return; }
    var x = idx[name] || [], teacher = String(x[4] || '');
    var mine = who.all || String(depts[normalizeName(teacher)] || '').toUpperCase() === who.dept;
    sheets.push({ name: name, cls: String(x[2] || '') || name, sub: String(x[3] || ''), teacher: teacher, mine: mine, r: raw.r || [], d: raw.d || {} });
  });
  var names = {};
  getMaster_().rows.forEach(function (m) { names[enr_(m[0])] = { name: m[1], sec: m[2], father: m[4] }; });
  return { sheets: sheets, names: names };
}

/** roll -> [ { p, sub, teacher, cls, m, mine } ] for one day (only students with a mark that day) */
function bnkDayMarks_(L, day) {
  var out = {};
  L.sheets.forEach(function (s) {
    var dd = s.d[String(day)];
    if (!dd) return;
    Object.keys(dd).forEach(function (p) {
      var str = String(dd[p]);
      for (var i = 0; i < s.r.length; i++) {
        var c = str.charAt(i);
        if (!VALID[c]) continue;
        var k = enr_(s.r[i]);
        (out[k] = out[k] || { roll: String(s.r[i]), list: [] }).list.push({ p: Number(p) || 0, sub: s.sub, teacher: s.teacher, cls: s.cls, m: c, mine: s.mine });
      }
    });
  });
  return out;
}

/** One student's day -> null (not a bunker) or the bunk record. */
function bnkJudge_(rec, who) {
  // the same period in two sheets (guest, moved section, duplicate roll): a P there means she WAS in class
  var inClass = {};
  rec.list.forEach(function (x) { if (x.m === 'P' && x.p) inClass[x.p] = 1; });
  var L = rec.list.filter(function (x) { return x.m === 'P' || !x.p || !inClass[x.p]; }).sort(function (a, b) { return a.p - b.p; });
  var P = L.filter(function (x) { return x.m === 'P'; }), miss = L.filter(function (x) { return x.m === 'A' || x.m === 'B'; });
  var hasB = L.some(function (x) { return x.m === 'B'; });
  if (!miss.length || (!P.length && !hasB)) return null;
  if (!who.all && !miss.some(function (x) { return x.mine; })) return null;   // HOD: only her department's lectures
  var pattern = 'Marked B by the teacher';
  if (P.length) {
    var firstP = P[0].p, lastP = P[P.length - 1].p;
    var late = miss.some(function (x) { return x.p && x.p < firstP; }), early = miss.some(function (x) { return x.p && x.p > lastP; });
    var mid = miss.some(function (x) { return x.p > firstP && x.p < lastP; });
    pattern = early && !late && !mid ? 'Left early' : late && !early && !mid ? 'Came late' : mid && !late && !early ? 'Skipped in between' : 'Missed lectures in between';
  }
  return { attended: P.length, missed: miss.length, pattern: pattern,
    lectures: L.map(function (x) { return { p: x.p, sub: x.sub, teacher: x.teacher, m: x.m }; }) };
}

function bnkDay_(who, dateTxt) {
  var D = bnkDate_(dateTxt), L = bnkLoad_(D.tok, who), marks = bnkDayMarks_(L, D.day), rows = [];
  Object.keys(marks).forEach(function (k) {
    var j = bnkJudge_(marks[k], who);
    if (!j) return;
    var n = L.names[k] || {};
    j.roll = marks[k].roll; j.name = n.name || ''; j.father = n.father || '';
    j.sec = n.sec || (marks[k].list[0] && marks[k].list[0].cls) || '';
    rows.push(j);
  });
  rows.sort(function (a, b) { return a.sec.localeCompare(b.sec, undefined, { numeric: true }) || b.missed - a.missed || a.name.localeCompare(b.name); });
  // repeat bunkers: how many days this month (up to this day) each of them bunked
  var times = bnkRepeat_(L, who, D.day);
  rows.forEach(function (r) { r.days = times[enr_(r.roll)] || 1; });
  return { ok: true, date: D.iso, tok: D.tok, day: D.day, scope: who.label,
    marked: Object.keys(marks).length, rows: rows };
}

/** roll -> number of days (1..upTo) she bunked this month */
function bnkRepeat_(L, who, upTo) {
  var days = {}, out = {};
  L.sheets.forEach(function (s) { Object.keys(s.d).forEach(function (d) { if (Number(d) <= upTo) days[d] = 1; }); });
  Object.keys(days).forEach(function (d) {
    var M = bnkDayMarks_(L, d);
    Object.keys(M).forEach(function (k) { if (bnkJudge_(M[k], who)) out[k] = (out[k] || 0) + 1; });
  });
  return out;
}

/** The whole month: every student who bunked, with the days and the subjects she missed most. */
function bnkMonth_(who, tok) {
  tok = String(tok || '').toUpperCase() || currentSeasonToken_();
  if (!tok || SEASON.indexOf(tok) < 0) throw new Error('Choose a month from September to January.');
  var L = bnkLoad_(tok, who), days = {}, S = {};
  L.sheets.forEach(function (s) { Object.keys(s.d).forEach(function (d) { days[d] = 1; }); });
  var dayList = Object.keys(days).map(Number).sort(function (a, b) { return a - b; });
  dayList.forEach(function (d) {
    var M = bnkDayMarks_(L, d);
    Object.keys(M).forEach(function (k) {
      var j = bnkJudge_(M[k], who);
      if (!j) return;
      var n = L.names[k] || {}, s = S[k] = S[k] || { roll: M[k].roll, name: n.name || '', sec: n.sec || M[k].list[0].cls || '', days: [], missed: 0, subs: {}, patterns: {} };
      s.days.push(d); s.missed += j.missed; s.patterns[j.pattern] = (s.patterns[j.pattern] || 0) + 1;
      j.lectures.forEach(function (x) { if (x.m === 'A' || x.m === 'B') s.subs[x.sub] = (s.subs[x.sub] || 0) + 1; });
    });
  });
  var rows = Object.keys(S).map(function (k) {
    var s = S[k], top = Object.keys(s.subs).sort(function (a, b) { return s.subs[b] - s.subs[a]; });
    var pat = Object.keys(s.patterns).sort(function (a, b) { return s.patterns[b] - s.patterns[a]; })[0];
    return { roll: s.roll, name: s.name, sec: s.sec, days: s.days, count: s.days.length, missed: s.missed,
      subjects: top.slice(0, 3).map(function (x) { return x + ' (' + s.subs[x] + ')'; }), pattern: pat };
  }).sort(function (a, b) { return b.count - a.count || b.missed - a.missed || a.sec.localeCompare(b.sec, undefined, { numeric: true }); });
  var my = monthYear_(tok);
  return { ok: true, tok: tok, m: my.m, y: my.y, scope: who.label, markedDays: dayList.length, rows: rows };
}

/* ---------- optional daily log (history that stays) ---------- */
function bnkSetupDailyLog() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'bnkDailyLog') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('bnkDailyLog').timeBased().everyDays(1).atHour(16).create();
  Logger.log('Bunkers: the daily log runs every day at about 4 pm.');
}

function bnkDailyLog() {
  var now = new Date();
  if (now.getDay() === 0 || !BNK_TOK_[now.getMonth()]) return;          // Sunday or outside Sept-Jan
  var r = bnkDay_({ all: true, label: 'College' }, Utilities.formatDate(now, TZ_(), 'yyyy-MM-dd'));
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(BNK_LOG_SHEET);
  if (!sh) {
    sh = book.insertSheet(BNK_LOG_SHEET);
    sh.getRange(1, 1, 1, BNK_LOG_HEAD.length).setValues([BNK_LOG_HEAD]).setFontWeight('bold').setBackground('#6B3423').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  // the same date is written only once (a second run replaces it)
  if (sh.getLastRow() > 1) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    for (var i = v.length - 1; i >= 0; i--) if (v[i][0] === r.date) sh.deleteRow(i + 2);
  }
  if (!r.rows.length) return;
  var fmt = function (list, want) {
    return list.filter(function (x) { return want(x.m); }).map(function (x) { return 'P' + x.p + ' ' + x.sub + (x.m === 'B' ? ' (B)' : ''); }).join(', ');
  };
  var out = r.rows.map(function (x) {
    return [r.date, x.roll, x.name, x.sec, x.pattern, x.attended, x.missed,
      fmt(x.lectures, function (m) { return m === 'A' || m === 'B'; }), fmt(x.lectures, function (m) { return m === 'P'; })];
  });
  var at = sh.getLastRow() + 1;
  sh.getRange(at, 1, out.length, 1).setNumberFormat('@');
  sh.getRange(at, 1, out.length, BNK_LOG_HEAD.length).setValues(out);
}
