/**
 * Bunkers.gs — Daily, weekly and monthly bunkers report (page bunkers.html).
 *
 * It reads the hidden "Period Data" sheet of each month workbook (the real mark of every period of every subject),
 * so teachers mark nothing new, and nothing is written to the attendance sheets.
 *
 * Two kinds of bunkers on a day:
 *   CONFIRMED  a teacher marked her "B" (in college but missed the lecture).
 *              If her other lectures that day are all A, she came to college and missed every lecture.
 *   POSSIBLE   present (P) in one lecture but absent (A) in another lecture of the same day, and no B.
 *              The A may be a real bunk or a real reason (sent home, left with permission …), so it waits
 *              for a decision: "Confirm bunk" or "Excused". Decisions are kept in the sheet "Bunk Decisions"
 *              of the Users workbook (the attendance itself is never changed).
 * Leave (L) is never bunking. A P in the same period in another sheet (guest, moved section) cancels an A/B there.
 * Days typed straight into the sheets (no Period Data, e.g. September) are read from the day cells of the sheets;
 * their period comes from the timetable (that weekday), so "left early / came late" still works.
 *
 * Who can open it: admin / principal and Chief Proctors (whole college), HOD (lectures of her department's teachers),
 * every teacher (students who bunked HER lectures). The same people can confirm or excuse a possible bunk.
 * Week / month answers are kept in the cache for 10 minutes, so opening them again is quick.
 *
 * INSTALL:
 *   a) + > Script > name it "Bunkers" > paste this file (replace the old one completely).
 *   b) doPost, next to the rpt_ line (once):
 *        if (/^bnk_/.test(String(req.action || ''))) return json_(bnkRoute_(user, req));
 *   c) Deploy > Manage deployments > Edit > New version > Deploy.
 *   d) Optional: run bnkSetupDailyLog once — every day at about 4 pm the day's bunkers are saved in "Bunkers Log".
 */
var BNK_LOG_SHEET = 'Bunkers Log';
var BNK_LOG_HEAD = ['Date', 'Roll No', 'Student', 'Section', 'Kind', 'Pattern', 'Attended', 'Missed', 'Missed lectures', 'Attended lectures', 'Decision'];
var BNK_DEC_SHEET = 'Bunk Decisions';
var BNK_DEC_HEAD = ['Date', 'Roll No', 'Student', 'Section', 'Decision', 'Note', 'By', 'When'];
var BNK_TOK_ = { 8: 'SEPT', 9: 'OCT', 10: 'NOV', 11: 'DEC', 0: 'JAN' };
var BNK_MAX_DAYS = 62;

function bnkRoute_(user, req) {
  try {
    var who = bnkWho_(user);
    if (!who) return { ok: false, error: 'Your account has no teacher name, so there are no lectures to check.' };
    var a = String(req.action || '');
    if (a === 'bnk_day')    return bnkDay_(who, String(req.date || ''));
    if (a === 'bnk_range')  return bnkRange_(who, String(req.from || ''), String(req.to || ''));
    if (a === 'bnk_month')  return bnkMonthRange_(who, String(req.mon || ''));
    if (a === 'bnk_decide') return bnkDecide_(user, req);
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

/** { all: true } or { all: false, dept } or null */
function bnkWho_(user) {
  var s = scopeOf_(user);
  if (s && s.all) return { all: true, label: 'College' };
  var cp = false; try { cp = isCP_(user); } catch (e) {}
  if (cp) return { all: true, label: 'College (Chief Proctor)' };
  if (s) return { all: false, dept: s.dept, label: s.dept || 'Department' };
  var me = normalizeName(user.teacher || '');
  return me ? { all: false, me: me, label: 'Your lectures (' + (user.teacher || '') + ')' } : null;
}

var bnkIso_ = function (d) { return Utilities.formatDate(d, TZ_(), 'yyyy-MM-dd'); };
function bnkParse_(txt) {
  var m = String(txt || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}

/* ---------------- data ---------------- */

var BNK_MEMO_ = {};
/** One month: period marks of every sheet + sheet info. */
function bnkLoad_(tok, who) {
  var key = tok + '|' + bnkScopeKey_(who);
  if (BNK_MEMO_[key]) return BNK_MEMO_[key];
  var book = bookFor_(tok), sh = book.getSheetByName(PD_SHEET), sheets = [], TT = bnkTT_(tok);
  if (sh && sh.getLastRow() > 1) {
    var idx = indexRows_(tok), depts = who.all || who.me ? null : teacherDepts_();
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var name = String(r[0] || ''), raw;
      if (!name) return;
      try { raw = JSON.parse(r[1] || '{}'); } catch (e) { return; }
      var x = idx[name] || [], t = TT.sheets[name] || {}, teacher = t.teacher || String(x[4] || '');
      var mine = bnkMine_(who, teacher, depts);
      sheets.push({ name: name, cls: t.cls || String(x[2] || '') || name, sub: t.sub || String(x[3] || ''), teacher: teacher, mine: mine, r: raw.r || [], d: raw.d || {} });
    });
  }
  // days that only exist in the day cells of the sheets (typed in, or marked before Period Data existed)
  var pdDays = {};
  sheets.forEach(function (s) { pdDays[s.name] = s.d; });
  var C = bnkCells_(tok), idx2 = indexRows_(tok), my = monthYear_(tok), dep2 = who.all || who.me ? null : teacherDepts_();
  Object.keys(C).forEach(function (name) {
    var c = C[name], have = pdDays[name] || {}, x = idx2[name] || [], t = TT.sheets[name] || {};
    var cls = t.cls || String(x[2] || '') || c.cls, sub = t.sub || String(x[3] || '') || c.sub;
    var teacher = t.teacher || String(x[4] || '') || c.teacher, d = {}, any = false;
    Object.keys(c.d).forEach(function (day) {
      if (have[day]) return;                                  // Period Data is the real record when it exists
      var p = bnkPeriod_(cls, sub, new Date(my.y, my.m, Number(day)).getDay());
      d[day] = {}; d[day][p] = c.d[day]; any = true;
    });
    if (!any) return;
    var mine = bnkMine_(who, teacher, dep2);
    sheets.push({ name: name, cls: cls, sub: sub, teacher: teacher, mine: mine, r: c.r, d: d });
  });
  return (BNK_MEMO_[key] = { tok: tok, sheets: sheets });
}

/**
 * Day cells of every attendance sheet of a month: { name: { cls, sub, teacher, r:[rolls], d:{ day: 'PAB-L…' } } }
 * (one letter per student; "2P" counts as P). Kept in the cache for 10 minutes.
 */
function bnkCells_(tok) {
  var ck = 'bnk_cells_' + tok, hit = cacheGetBig_(ck);
  if (hit) return hit;
  var book = bookFor_(tok), names = book.getSheets().map(function (sh) { return sh.getName(); })
    .filter(function (n) { return n !== PD_SHEET && (' ' + n + ' ').indexOf(' ' + tok + ' ') >= 0; });
  var grids = {};
  try {                                                       // fast: one call (Sheets advanced service, if it is on)
    for (var k = 0; k < names.length; k += 100) {
      var part = names.slice(k, k + 100);
      var res = Sheets.Spreadsheets.Values.batchGet(book.getId(), { ranges: part.map(function (n) { return "'" + n.replace(/'/g, "''") + "'!A2:AZ"; }) });
      (res.valueRanges || []).forEach(function (vr, i) { grids[part[i]] = vr.values || []; });
    }
  } catch (e) {
    grids = {};
    names.forEach(function (n) {                              // slower: one sheet at a time
      var sh = book.getSheetByName(n), lr = sh ? sh.getLastRow() : 0;
      if (lr >= FIRST_ROW) grids[n] = sh.getRange(2, 1, lr - 1, Math.min(sh.getLastColumn(), 52)).getDisplayValues();
    });
  }
  var out = {};
  Object.keys(grids).forEach(function (n) {
    var g = grids[n]; if (g.length < 3) return;
    var hdr = g[HEADER_ROW - 2] || [], rc = -1, dc = {};
    hdr.forEach(function (h, i) { var t = norm_(h); if (t === 'ROLLNO' && rc < 0) rc = i; else if (/^\d+$/.test(t) && +t >= 1 && +t <= 31 && dc[+t] == null) dc[+t] = i; });
    if (rc < 0) return;
    var rows = g.slice(FIRST_ROW - 2).filter(function (r) { return String(r[rc] || '').trim(); });
    if (!rows.length) return;
    var d = {};
    Object.keys(dc).forEach(function (day) {
      var str = rows.map(function (r) { var m = String(r[dc[day]] || '').trim().toUpperCase().match(/^\d*([PALB])$/); return m ? m[1] : '-'; }).join('');
      if (/[PALB]/.test(str)) d[day] = str;
    });
    if (!Object.keys(d).length) return;
    var parts = n.split(' ' + tok + ' ');
    out[n] = { cls: parts[0] || n, sub: parts.slice(1).join(' ' + tok + ' '), teacher: String((g[0] || [])[0] || '').replace(/^Teacher:\s*/i, ''),
      r: rows.map(function (r) { return String(r[rc]).trim(); }), d: d };
  });
  cachePutBig_(ck, out);
  return out;
}

/** sheet name -> { cls, sub, teacher, program } exactly as written in the Timetable (the sheets are made from it). */
var BNK_TT_ = {};
function bnkTT_(tok) {
  if (BNK_TT_[tok]) return BNK_TT_[tok];
  var out = { sheets: {}, prog: {} };
  try {
    getTimetable_().rows.forEach(function (r) {
      var pcs = String(r[3] || '').trim(), sub = String(r[4] || '').trim();
      if (!pcs || !sub) return;
      var name = makeName_(pcs, tok, sub.toUpperCase());
      if (!out.sheets[name]) out.sheets[name] = { cls: pcs, sub: sub.toUpperCase(), teacher: String(r[0] || '').trim(), program: String(r[1] || '').trim() };
      var k = pcs.replace(/\s+/g, '').toUpperCase();
      if (!out.prog[k] && r[1]) out.prog[k] = String(r[1]).trim();
    });
  } catch (e) {}
  return (BNK_TT_[tok] = out);
}

/** The period of this lecture on that weekday, from the Timetable ('0' when the Timetable does not say). */
function bnkPeriod_(cls, sub, dow) {
  try { var sl = slotsOf_(cls, sub).filter(function (q) { return q.days[dow]; }); if (sl.length) return String(sl[0].p); } catch (e) {}
  return '0';
}

function bnkScopeKey_(who) { return who.all ? '*' : who.me ? 'T:' + who.me : 'D:' + who.dept; }
/** Is this lecture's teacher inside the user's view? (college / her department / herself) */
function bnkMine_(who, teacher, depts) {
  if (who.all) return true;
  if (who.me) return normalizeName(teacher) === who.me;
  return String(depts[normalizeName(teacher)] || '').toUpperCase() === who.dept;
}

var BNK_NAMES_ = null;
function bnkNames_() {
  if (BNK_NAMES_) return BNK_NAMES_;
  BNK_NAMES_ = {};
  getMaster_().rows.forEach(function (m) { BNK_NAMES_[enr_(m[0])] = { name: m[1], sec: m[2], father: m[4] }; });
  return BNK_NAMES_;
}

/** 'yyyy-MM-dd|ROLL' -> { decision, note, by, when } */
function bnkDecisions_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(BNK_DEC_SHEET), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, BNK_DEC_HEAD.length).getDisplayValues().forEach(function (r) {
    if (r[0] && r[1]) out[r[0] + '|' + enr_(r[1])] = { decision: r[4], note: r[5], by: r[6], when: r[7] };
  });
  return out;
}

/** roll -> { roll, list:[ { p, sub, teacher, cls, m, mine } ] } for one day of one month */
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

/** One student's day -> null or { kind: 'confirmed'|'possible', pattern, attended, missed, lectures } */
function bnkJudge_(rec, who) {
  // the same period in two sheets (guest, moved section, duplicate roll): a P there means she WAS in class
  var inClass = {};
  rec.list.forEach(function (x) { if (x.m === 'P' && x.p) inClass[x.p] = 1; });
  var L = rec.list.filter(function (x) { return x.m === 'P' || !x.p || !inClass[x.p]; }).sort(function (a, b) { return a.p - b.p; });
  var P = L.filter(function (x) { return x.m === 'P'; }), miss = L.filter(function (x) { return x.m === 'A' || x.m === 'B'; });
  var B = L.filter(function (x) { return x.m === 'B'; });
  if (!miss.length) return null;
  if (!B.length && !P.length) return null;                                   // absent the whole day = not a bunker
  if (!who.all && !miss.some(function (x) { return x.mine; })) return null;   // HOD: her department's lectures; teacher: her own lectures
  var kind = B.length ? 'confirmed' : 'possible', pattern;
  if (P.length && L.some(function (x) { return !x.p; })) pattern = 'Missed other lectures the same day';   // period not known
  else if (!P.length) pattern = miss.length > 1 ? 'In college, missed every lecture' : 'Marked B by the teacher';
  else {
    var firstP = P[0].p, lastP = P[P.length - 1].p;
    var late = miss.some(function (x) { return x.p && x.p < firstP; }), early = miss.some(function (x) { return x.p && x.p > lastP; });
    var mid = miss.some(function (x) { return x.p > firstP && x.p < lastP; });
    pattern = early && !late && !mid ? 'Left early' : late && !early && !mid ? 'Came late' : mid && !late && !early ? 'Skipped in between' : 'Missed lectures in between';
  }
  return { kind: kind, pattern: pattern, attended: P.length, missed: miss.length,
    lectures: L.map(function (x) { return { p: x.p, sub: x.sub, teacher: x.teacher, m: x.m }; }) };
}

/** A decision: "Confirmed" makes a possible bunk confirmed; "Excused" takes it out of every count. */
function bnkApply_(j, dec) {
  if (!dec) return j;
  j.decision = dec.decision; j.decNote = dec.note; j.decBy = dec.by;
  if (dec.decision === 'Excused') j.kind = 'excused';
  else if (dec.decision === 'Confirmed') j.kind = 'confirmed';
  return j;
}

/** All judged students of one calendar date: { iso, rows:[...], marked } */
function bnkDate_(who, d, DEC) {
  var tok = BNK_TOK_[d.getMonth()], iso = bnkIso_(d);
  if (!tok) return { iso: iso, rows: [], marked: 0 };
  var L = bnkLoad_(tok, who), M = bnkDayMarks_(L, d.getDate()), names = bnkNames_(), rows = [];
  Object.keys(M).forEach(function (k) {
    var j = bnkJudge_(M[k], who);
    if (!j) return;
    bnkApply_(j, DEC[iso + '|' + k]);
    var n = names[k] || {};
    j.key = k; j.roll = M[k].roll; j.name = n.name || ''; j.father = n.father || '';
    j.sec = n.sec || (M[k].list[0] && M[k].list[0].cls) || '';
    rows.push(j);
  });
  return { iso: iso, rows: rows, marked: Object.keys(M).length };
}

/* ---------------- day ---------------- */

function bnkDay_(who, dateTxt) {
  var d = bnkParse_(dateTxt) || new Date();
  if (!BNK_TOK_[d.getMonth()]) throw new Error('Attendance is only kept from September to January.');
  var DEC = bnkDecisions_(), R = bnkDate_(who, d, DEC), o = { confirmed: 0, possible: 1, excused: 2 };
  R.rows.sort(function (a, b) {
    return a.sec.localeCompare(b.sec, undefined, { numeric: true }) || o[a.kind] - o[b.kind] || b.missed - a.missed || a.name.localeCompare(b.name);
  });
  // days bunked this month up to this day (confirmed + possible, not excused)
  var cnt = {};
  for (var x = new Date(d.getFullYear(), d.getMonth(), 1); x <= d; x.setDate(x.getDate() + 1)) {
    if (x.getDay() === 0) continue;
    bnkDate_(who, new Date(x), DEC).rows.forEach(function (r) {
      if (r.kind === 'excused') return;
      var c = cnt[r.key] = cnt[r.key] || { c: 0, p: 0 };
      if (r.kind === 'confirmed') c.c++; else c.p++;
    });
  }
  R.rows.forEach(function (r) { var c = cnt[r.key] || { c: 0, p: 0 }; r.days = c.c + c.p; r.daysC = c.c; r.daysP = c.p; });
  return { ok: true, date: R.iso, scope: who.label, marked: R.marked, rows: R.rows };
}

/* ---------------- week / month (any range) ---------------- */

function bnkMonthRange_(who, tok) {
  tok = String(tok || '').toUpperCase() || currentSeasonToken_();
  if (!tok || SEASON.indexOf(tok) < 0) throw new Error('Choose a month from September to January.');
  var my = monthYear_(tok);
  var from = new Date(my.y, my.m, 1), to = new Date(my.y, my.m + 1, 0), today = new Date();
  if (to > today) to = today;
  return bnkRange_(who, bnkIso_(from), bnkIso_(to));
}

function bnkRange_(who, fromTxt, toTxt) {
  var from = bnkParse_(fromTxt), to = bnkParse_(toTxt);
  if (!from || !to || to < from) throw new Error('Choose a correct period.');
  if ((to - from) / 864e5 > BNK_MAX_DAYS) throw new Error('Choose at most ' + BNK_MAX_DAYS + ' days.');
  var ver = PropertiesService.getScriptProperties().getProperty('BNK_DEC_VER') || '0';
  var ck = 'bnk_r_' + ver.slice(-6) + bnkScopeKey_(who).replace(/[^\w:*]/g, '').slice(0, 40) + '_' + bnkIso_(from) + '_' + bnkIso_(to), hit = cacheGetBig_(ck);
  if (hit) { hit.cached = true; return hit; }
  var res = bnkRangeBuild_(who, from, to);
  delete res.rows;                                   // the page builds the student list from the events (smaller answer)
  cachePutBig_(ck, res);
  return res;
}

function bnkRangeBuild_(who, from, to) {
  var DEC = bnkDecisions_(), S = {}, days = 0, perDay = [], ev = [], st = {}, dict = [], dIx = {};
  var di = function (t) { t = String(t || ''); if (dIx[t] == null) { dIx[t] = dict.length; dict.push(t); } return dIx[t]; };   // each subject / teacher / pattern once
  for (var x = new Date(from); x <= to; x.setDate(x.getDate() + 1)) {
    if (x.getDay() === 0 || !BNK_TOK_[x.getMonth()]) continue;
    var R = bnkDate_(who, new Date(x), DEC), c = 0, p = 0;
    if (R.marked) days++;
    R.rows.forEach(function (j) {
      if (j.kind === 'excused') return;
      if (j.kind === 'confirmed') c++; else p++;
      // compact event list for the section / program / class / subject / teacher / weekday / period tables
      if (!st[j.key]) st[j.key] = [j.roll, j.name, j.father, j.sec, bnkProgram_(j.sec, x)];
      ev.push([j.key, R.iso, j.kind === 'confirmed' ? 'C' : 'P', di(j.pattern),
        j.lectures.filter(function (l) { return l.m === 'A' || l.m === 'B'; }).map(function (l) { return [l.p, di(l.sub), di(l.teacher), l.m]; }),
        j.lectures.filter(function (l) { return l.m === 'P' || l.m === 'L'; }).map(function (l) { return [l.p, di(l.sub), l.m]; })]);
      var s = S[j.key] = S[j.key] || { roll: j.roll, name: j.name, father: j.father, sec: j.sec, dates: [], confirmed: 0, possible: 0, missed: 0, subs: {}, patterns: {}, teachers: {} };
      s.dates.push({ d: R.iso, k: j.kind === 'confirmed' ? 'C' : 'P' });
      if (j.kind === 'confirmed') s.confirmed++; else s.possible++;
      s.missed += j.missed; s.patterns[j.pattern] = (s.patterns[j.pattern] || 0) + 1;
      j.lectures.forEach(function (l) {
        if (l.m !== 'A' && l.m !== 'B') return;
        s.subs[l.sub] = (s.subs[l.sub] || 0) + 1;
        if (l.teacher) s.teachers[l.teacher] = 1;
      });
    });
    if (R.marked) perDay.push({ d: R.iso, confirmed: c, possible: p, marked: R.marked });
  }
  var top = function (o, n) { return Object.keys(o).sort(function (a, b) { return o[b] - o[a]; }).slice(0, n); };
  var rows = Object.keys(S).map(function (k) {
    var s = S[k];
    return { roll: s.roll, name: s.name, father: s.father, sec: s.sec, dates: s.dates, count: s.dates.length,
      confirmed: s.confirmed, possible: s.possible, missed: s.missed,
      subjects: top(s.subs, 3).map(function (x) { return x + ' (' + s.subs[x] + ')'; }), pattern: top(s.patterns, 1)[0] || '',
      teachers: Object.keys(s.teachers) };
  }).sort(function (a, b) { return b.confirmed - a.confirmed || b.count - a.count || b.missed - a.missed || a.sec.localeCompare(b.sec, undefined, { numeric: true }); });
  return { ok: true, from: bnkIso_(from), to: bnkIso_(to), scope: who.label, markedDays: days, perDay: perDay, rows: rows, ev: ev, st: st, dict: dict };
}

/** Program of a section, as written in the Timetable (e.g. "FCS-1 D2" -> "FCS (JCW)"). */
function bnkProgram_(sec, d) {
  var tok = BNK_TOK_[d.getMonth()], k = String(sec || '').replace(/\s+/g, '').toUpperCase();
  return (tok && bnkTT_(tok).prog[k]) || '';
}

/* ---------------- decisions ---------------- */

/** req: { date, roll, decision: 'Confirmed' | 'Excused' | '' (clear), note } */
function bnkDecide_(user, req) {
  var d = bnkParse_(req.date), roll = String(req.roll || '').trim(), dec = String(req.decision || '');
  if (!d || !roll) return { ok: false, error: 'Missing date or roll number.' };
  if (dec && dec !== 'Confirmed' && dec !== 'Excused') return { ok: false, error: 'Unknown decision.' };
  var iso = bnkIso_(d), lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(BNK_DEC_SHEET);
    if (!sh) {
      sh = book.insertSheet(BNK_DEC_SHEET);
      sh.getRange(1, 1, 1, BNK_DEC_HEAD.length).setValues([BNK_DEC_HEAD]).setFontWeight('bold').setBackground('#6B3423').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    if (sh.getLastRow() > 1) {                                   // one decision per student per day
      var v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
      for (var i = v.length - 1; i >= 0; i--) if (v[i][0] === iso && enr_(v[i][1]) === enr_(roll)) sh.deleteRow(i + 2);
    }
    PropertiesService.getScriptProperties().setProperty('BNK_DEC_VER', String(Date.now()));   // week / month answers are made again
    if (dec) {
      var n = bnkNames_()[enr_(roll)] || {};
      var at = sh.getLastRow() + 1;
      sh.getRange(at, 1, 1, 2).setNumberFormat('@');
      sh.getRange(at, 1, 1, BNK_DEC_HEAD.length).setValues([[iso, roll, n.name || '', n.sec || '', dec, String(req.note || '').slice(0, 300), user.teacher || user.username, new Date()]]);
    }
    return { ok: true, date: iso, roll: roll, decision: dec };
  } finally { lock.releaseLock(); }
}

/* ---------------- optional daily log (history that stays) ---------------- */
function bnkSetupDailyLog() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'bnkDailyLog') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('bnkDailyLog').timeBased().everyDays(1).atHour(16).create();
  Logger.log('Bunkers: the daily log runs every day at about 4 pm.');
}

function bnkDailyLog() {
  var now = new Date();
  if (now.getDay() === 0 || !BNK_TOK_[now.getMonth()]) return;
  var R = bnkDate_({ all: true, label: 'College' }, now, bnkDecisions_());
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(BNK_LOG_SHEET);
  if (!sh) {
    sh = book.insertSheet(BNK_LOG_SHEET);
    sh.getRange(1, 1, 1, BNK_LOG_HEAD.length).setValues([BNK_LOG_HEAD]).setFontWeight('bold').setBackground('#6B3423').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() > 1) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    for (var i = v.length - 1; i >= 0; i--) if (v[i][0] === R.iso) sh.deleteRow(i + 2);
  }
  if (!R.rows.length) return;
  var fmt = function (list, want) {
    return list.filter(function (x) { return want(x.m); }).map(function (x) { return 'P' + x.p + ' ' + x.sub + (x.m === 'B' ? ' (B)' : ''); }).join(', ');
  };
  var KIND = { confirmed: 'Confirmed (B)', possible: 'Possible', excused: 'Excused' };
  var out = R.rows.map(function (x) {
    return [R.iso, x.roll, x.name, x.sec, KIND[x.kind], x.pattern, x.attended, x.missed,
      fmt(x.lectures, function (m) { return m === 'A' || m === 'B'; }), fmt(x.lectures, function (m) { return m === 'P'; }), x.decision || ''];
  });
  var at = sh.getLastRow() + 1;
  sh.getRange(at, 1, out.length, 2).setNumberFormat('@');
  sh.getRange(at, 1, out.length, BNK_LOG_HEAD.length).setValues(out);
}
