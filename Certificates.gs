/**
 * Certificates.gs — Provisional cum Character Certificates inside the attendance app (admin only).
 *
 * ONE GO: after the gazette PDF is saved by board.html ("HSSC 2nd year 2026" …), pc_prepare does, for the
 * Jinnah College rows only (other colleges are never written):
 *   1. Left Students  -> row struck through, "Matched by" = LEFT STUDENT-…, no certificate
 *                        (Enroll#, REG# or exact Name + Father Name; never when the same student is still in Transform)
 *   2. Transform-ALL-COLLEGE-STRENGTH -> Enroll#, SName, SFname, Program, Class, Section, PCS, Group
 *                        (existing Enroll#, then exact Name + Father Name; the app then offers 96%+ matches)
 *   3. Portal backup  -> DOB (01/01/2000 and empty left out)
 *   4. Grade          -> Grade column of the gazette; else a grade in Remarks; else worked out from Marks
 *   5. Session (e.g. 2025-2026) and Distinction_won (board position 1-20) when empty
 * Nothing already typed by hand is overwritten (Enroll#, DOB, Group, Session, Distinction_won, Grade).
 * The same columns are used by the "Certificates" menu script in the workbook, so both keep working.
 *
 * INSTALL: add this file, and in doPost (next to the br_ line):
 *     if (/^pc_/.test(String(req.action || ''))) return json_(pcRoute_(user, req));
 * then Deploy → Manage deployments → Edit → New version.
 * For the PDFs keep jcw.png and madam_sign.png in the Drive of the account that runs the script.
 */
var PC_CFG = {
  BOOK: '1C2tdr2yXAja0_C5Nk4HIk2m-71Kc4I-FS5VQUhjUtHQ',          // Provisional Certificates workbook
  USERS: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',         // Users workbook
  STRENGTH: 'Transform-ALL-COLLEGE-STRENGTH',
  LEFT: 'Left Students',
  PORTAL: 'Portal backup',
  COLLEGE: /JINNAH COLLEGE FOR WOMEN/i,
  PASS: ['A1', 'A', 'B', 'C', 'D', 'E'],
  FILL: ['Enroll#', 'SName', 'SFname', 'Program', 'Class', 'Section', 'PCS', 'Group', 'Session', 'Distinction_won', 'DOB', 'Grade', 'Matched by'],
  EXTRA: ['Total Marks', 'Conduct'],                              // made only when the admin types a value
  DEF: { conduct: 'GOOD', distinctions: 'NIL' },
  TOP: 20,
  TITLE: 'PROVISIONAL CUM CHARACTER CERTIFICATE',
  COLLEGE_NAME: 'JINNAH COLLEGE FOR WOMEN',
  UNIVERSITY: 'UNIVERSITY OF PESHAWAR',
  GAZETTE: 'has been declared successful according to the Result Gazette, of Board of Intermediate and Secondary Education, Peshawar.',
  SIGN: ['Principal,', 'Jinnah College for Women,', 'University of Peshawar.'],
  LOGO: 'jcw.png', SIGN_IMG: 'madam_sign.png', FOLDER: 'Certificates (all)'
};

/** Run this ONCE from the Apps Script editor (select pcAuthorize ▸ Run ▸ Allow), then deploy a new version.
 *  Certificates.gs uses Drive (logo, signature, PDFs) and the Users / Provisional Certificates workbooks;
 *  until those permissions are allowed every call from the app fails and the page shows a network error. */
function pcAuthorize() {
  SpreadsheetApp.openById(PC_CFG.BOOK).getName();
  SpreadsheetApp.openById(PC_CFG.USERS).getName();
  DriveApp.getFilesByName(PC_CFG.LOGO).hasNext();
  CacheService.getScriptCache().get('x');
  LockService.getScriptLock().tryLock(1);
  Logger.log('Certificates: permissions OK. Now Deploy → Manage deployments → Edit → New version.');
}

function pcRoute_(user, req) {
  try {
    if (!isAdmin_(user)) return { ok: false, error: 'Only the admin can make certificates.' };
    switch (String(req.action || '')) {
      case 'pc_sheets':  return pcSheets_();
      case 'pc_prepare': return pcPrepare_(req.sheet, req.strength);
      case 'pc_load':    return pcLoad_(req.sheet, null, req.strength);
      case 'pc_link':    return pcLink_(req);
      case 'pc_save':    return pcSave_(req);
      case 'pc_strike':  return pcStrike_(req);
      case 'pc_assets':  return pcAssets_();
      case 'pc_pdf':     return pcPdf_(req);
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/* ============================================================ helpers */
function pcClean_(v) { return String(v == null ? '' : v).replace(/[​-‏⁠﻿]/g, '').replace(/[ \s]+/g, ' ').trim(); }
function pcHead_(v) { return pcClean_(v).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function pcLetters_(v) { return pcClean_(v).toUpperCase().replace(/[^A-Z]/g, ''); }
function pcCol_(H, names) { for (var i = 0; i < names.length; i++) { var k = H.indexOf(pcHead_(names[i])); if (k >= 0) return k; } return -1; }
function pcSheet_(name) {
  var sh = SpreadsheetApp.openById(PC_CFG.BOOK).getSheetByName(String(name || ''));
  if (!sh) throw new Error('Sheet "' + name + '" not found in Provisional Certificates. Save the gazette first (Board results).');
  return sh;
}
function pcPart_(name) { return /2nd|part.?ii\b|part.?2/i.test(name) ? 2 : 1; }
function pcYear_(name) { var m = String(name).match(/(20\d\d)/); return m ? Number(m[1]) : new Date().getFullYear(); }
function pcTotal_(name) { return pcPart_(name) === 2 ? 1200 : 600; }
function pcOrd_(n) { var s = ['th', 'st', 'nd', 'rd'], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); }
function pcGradeOf_(marks, total) {
  var m = Number(String(marks).replace(/,/g, '')); if (!isFinite(m) || !total || String(marks).trim() === '') return '';
  var p = m / total * 100;
  return p >= 80 ? 'A1' : p >= 70 ? 'A' : p >= 60 ? 'B' : p >= 50 ? 'C' : p >= 40 ? 'D' : p >= 33 ? 'E' : 'F';
}
function pcPlaceholderDob_(s) {
  s = String(s).trim(); var m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/))) return +m[1] === 1 && +m[2] === 1 && +m[3] === 2000;
  if ((m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/))) return +m[1] === 2000 && +m[2] === 1 && +m[3] === 1;
  return false;
}

/** columns of an HSSC sheet (0-based); FILL columns are added at the right when missing */
function pcCols_(sh, make) {
  var lastCol = sh.getLastColumn(), H = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(pcHead_);
  var need = function (title) {
    var c = H.indexOf(pcHead_(title));
    if (c < 0 && make) { lastCol++; sh.getRange(1, lastCol).setValue(title).setFontWeight('bold').setBackground('#E7B7AE'); c = lastCol - 1; H[c] = pcHead_(title); }
    return c;
  };
  var C = {
    roll: pcCol_(H, ['Roll No', 'RollNo']), name: pcCol_(H, ['Name']), father: pcCol_(H, ['Father Name']),
    reg: pcCol_(H, ['Registration #', 'Registration', 'Registration No']), marks: pcCol_(H, ['Marks']),
    remarks: pcCol_(H, ['Remarks']), college: pcCol_(H, ['College Name']), pos: pcCol_(H, ['Position'])
  };
  if (C.name < 0 || C.father < 0) throw new Error('"' + sh.getName() + '" needs Name and Father Name columns.');
  C.f = {}; PC_CFG.FILL.forEach(function (t) { C.f[t] = need(t); });
  PC_CFG.EXTRA.forEach(function (t) { C.f[t] = H.indexOf(pcHead_(t)); });
  C.addExtra = function (t) { if (C.f[t] < 0) { make = true; C.f[t] = need(t); C.last = lastCol; } return C.f[t]; };
  C.last = lastCol;
  return C;
}

/* ============================================================ sources (Users workbook) */
/** Students from Transform-ALL-COLLEGE-STRENGTH and its older copies ("Transform-ALL-COLLEGE-STRENGTH 2025" …),
 *  so a 2nd-year gazette also finds students who are no longer in the current sheet.
 *  The current sheet wins when an Enroll# is in both. name = one sheet only (the box on the page). */
function pcStrength_(name) {
  var book = SpreadsheetApp.openById(PC_CFG.USERS), names;
  name = pcClean_(name);
  if (name) {
    if (!book.getSheetByName(name)) throw new Error('Sheet "' + name + '" not found in the Users workbook.');
    names = [name];
  } else {
    var base = pcHead_(PC_CFG.STRENGTH);
    names = book.getSheets().map(function (x) { return x.getName(); })
      .filter(function (n) { return pcHead_(n).indexOf(base) === 0; })
      .sort(function (a, b) { return pcHead_(a) === base ? -1 : pcHead_(b) === base ? 1 : b.localeCompare(a); });
    if (!names.length) throw new Error('Sheet "' + PC_CFG.STRENGTH + '" not found in the Users workbook.');
  }
  var S = { rows: [], byEnroll: {}, byName: {}, cur: { byEnroll: {}, byName: {} }, sheets: names };
  names.forEach(function (sn, si) {
    var v = book.getSheetByName(sn).getDataRange().getDisplayValues(), H = (v[0] || []).map(pcHead_);
    var c = { e: pcCol_(H, ['Enroll#', 'Enrollment#', 'Enrollment No', 'REG#']), n: pcCol_(H, ['Student Name', 'Name']), f: pcCol_(H, ['Father Name', 'Father']),
      p: pcCol_(H, ['Program', 'Program Name']), c: pcCol_(H, ['Class']), s: pcCol_(H, ['Section']), x: pcCol_(H, ['PCS']), a: pcCol_(H, ['Appeared in']),
      st: pcCol_(H, ['Portal Status', 'Student Status', 'Enrollment Status', 'Status']) };
    if (c.e < 0 || c.n < 0) { if (si === 0) throw new Error(sn + ' needs Enroll# and Student Name columns.'); return; }
    var g = function (r, i) { return i >= 0 ? pcClean_(r[i]) : ''; };
    for (var i = 1; i < v.length; i++) {
      var r = v[i], o = { enroll: g(r, c.e), name: g(r, c.n), father: g(r, c.f), program: g(r, c.p), cls: g(r, c.c), section: g(r, c.s), pcs: g(r, c.x), group: g(r, c.a), status: g(r, c.st) };
      var k = pcHead_(o.enroll); if (!k || S.byEnroll[k]) continue;
      var nk = pcLetters_(o.name) + '|' + pcLetters_(o.father);
      S.rows.push(o); S.byEnroll[k] = o;
      if (nk !== '|') (S.byName[nk] = S.byName[nk] || []).push(o);
      if (si === 0) { S.cur.byEnroll[k] = o; if (nk !== '|') (S.cur.byName[nk] = S.cur.byName[nk] || []).push(o); }
    }
  });
  return S;
}
function pcLeft_() {
  var book = SpreadsheetApp.openById(PC_CFG.USERS), base = pcHead_(PC_CFG.LEFT);
  var out = { byEnroll: {}, byReg: {}, byName: {}, byOnly: {}, list: [], count: 0, sheets: [] };
  book.getSheets().forEach(function (sh) {
    if (pcHead_(sh.getName()).indexOf(base) !== 0 || sh.getLastRow() < 2) return;
    var v = sh.getDataRange().getDisplayValues(), at = -1, ix = null;
    for (var h = 0; h < Math.min(10, v.length); h++) {             // the header may sit on row 2
      var H = v[h].map(pcHead_), x = { e: pcCol_(H, ['Enroll#', 'Enrollment#', 'Enrollment No', 'Enroll No']), n: pcCol_(H, ['Student Name', 'Name']), f: pcCol_(H, ['Father Name', 'Father']),
        r: pcCol_(H, ['REG#', 'Registration #', 'Registration No']), st: pcCol_(H, ['Portal Status', 'Status']) };
      if (x.n >= 0 && (x.e >= 0 || x.r >= 0 || x.f >= 0)) { at = h; ix = x; break; }
    }
    if (at < 0) return;
    out.sheets.push(sh.getName());
    for (var i = at + 1; i < v.length; i++) {
      var r = v[i], g = function (k) { return k >= 0 ? pcClean_(r[k]) : ''; };
      var rec = { enroll: g(ix.e), reg: g(ix.r), name: g(ix.n), father: g(ix.f), status: g(ix.st) };
      if (!rec.name && !rec.enroll && !rec.reg) continue;
      out.count++; rec.nl = pcLetters_(rec.name); rec.fl = pcLetters_(rec.father); out.list.push(rec);
      if (rec.enroll) out.byEnroll[pcHead_(rec.enroll)] = out.byEnroll[pcHead_(rec.enroll)] || rec;
      if (rec.reg) out.byReg[pcHead_(rec.reg)] = out.byReg[pcHead_(rec.reg)] || rec;
      if (rec.name && rec.father) out.byName[pcLetters_(rec.name) + '|' + pcLetters_(rec.father)] = rec;
      if (rec.name) (out.byOnly[pcLetters_(rec.name)] = out.byOnly[pcLetters_(rec.name)] || []).push(rec);
    }
  });
  return out;
}
/** 0..1 similarity of two letter strings (1 - edit distance / length) */
function pcSim_(a, b) {
  if (a === b) return 1; if (!a || !b) return 0;
  var m = a.length, n = b.length, p = [], c, i, j;
  for (j = 0; j <= n; j++) p[j] = j;
  for (i = 1; i <= m; i++) { c = [i]; for (j = 1; j <= n; j++) c[j] = Math.min(p[j] + 1, c[j - 1] + 1, p[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); p = c; }
  return 1 - p[n] / Math.max(m, n);
}
/** the closest Left Students record by name (+ father); {rec, ns, fs, sc} or null */
function pcLeftNear_(L, name, father) {
  var nl = pcLetters_(name), fl = pcLetters_(father), best = null, second = 0;
  if (!nl) return null;
  L.list.forEach(function (rec) {
    if (!rec.nl || Math.abs(rec.nl.length - nl.length) > 10) return;
    var ns = pcSim_(nl, rec.nl);
    if (rec.nl.indexOf(nl) === 0 || nl.indexOf(rec.nl) === 0) ns = Math.max(ns, 0.85);   // "KOMAL" / "KOMAL KHAN"
    if (ns < 0.75) return;
    var fs = fl && rec.fl ? pcSim_(fl, rec.fl) : 0.5, sc = 0.6 * ns + 0.4 * fs;
    if (!best || sc > best.sc) { if (best) second = best.sc; best = { rec: rec, ns: ns, fs: fs, sc: sc }; } else if (sc > second) second = sc;
  });
  if (best) best.unique = best.sc - second >= 0.03;
  return best;
}
function pcLeftOf_(L, S, name, father, enroll, reg) {
  var ek = pcHead_(enroll), rk = pcHead_(reg), nk = pcLetters_(name) + '|' + pcLetters_(father);
  if (ek && L.byEnroll[ek] && !S.cur.byEnroll[ek]) return L.byEnroll[ek];
  if (rk && L.byReg[rk]) return L.byReg[rk];
  if ((S.cur.byName[nk] || []).length || (ek && S.cur.byEnroll[ek])) return null;           // still studying -> not left
  if (name && father && L.byName[nk]) return L.byName[nk];
  // same name, father written a little differently ("M. Yameen" / "Muhammad Yameen"): one such record, same last word of father
  var one = L.byOnly[pcLetters_(name)] || [], lw = function (s) { var w = pcClean_(s).toUpperCase().replace(/[^A-Z ]/g, ' ').trim().split(/\s+/); return w[w.length - 1] || ''; };
  if (one.length === 1 && father && one[0].father && lw(one[0].father) === lw(father)) return one[0];
  // spelling differences in both names ("Wareesha"/"Warisha", "Mehmood"/"Mahmood"): very close and only one such record
  var nr = pcLeftNear_(L, name, father);
  if (nr && nr.unique && ((nr.ns >= 0.88 && nr.fs >= 0.8) || (nr.ns >= 0.8 && nr.fs >= 0.95))) return nr.rec;
  return null;
}
/** The sheet row of this student: the given row if it still holds her, else her Roll No, else her unique Name + Father Name
 *  (rows can move when the sheet is sorted or filtered). 0 = not found. */
function pcLocate_(sh, C, row, roll, name, father) {
  var n = sh.getLastRow() - 1; if (n < 1) return 0;
  var col = function (k) { return k >= 0 ? sh.getRange(2, k + 1, n, 1).getDisplayValues().map(function (x) { return pcClean_(x[0]); }) : null; };
  var N = col(C.name), Fa = col(C.father), R = col(C.roll), same = function (i) { return pcLetters_(N[i]) === pcLetters_(name) && pcLetters_(Fa[i]) === pcLetters_(father); };
  roll = pcClean_(roll); row = Number(row) || 0;
  if (row >= 2 && row - 2 < n && (roll && R ? R[row - 2] === roll : same(row - 2))) return row;
  var i, hit = [];
  if (roll && R) for (i = 0; i < n; i++) if (R[i] === roll) hit.push(i);
  if (hit.length === 1) return hit[0] + 2;
  hit = []; for (i = 0; i < n; i++) if (same(i)) hit.push(i);
  return hit.length === 1 ? hit[0] + 2 : 0;
}
function pcLeftNote_(rec) { return 'LEFT STUDENT-' + (pcClean_(rec.status).toUpperCase() || 'LEFT'); }
/** Portal backup: DOB and portal status (Enrolled, Migrated, Cancelled …) by Enroll# */
function pcPortal_() {
  var sh = SpreadsheetApp.openById(PC_CFG.USERS).getSheetByName(PC_CFG.PORTAL), out = { dob: {}, status: {} };
  if (!sh) return out;
  var v = sh.getDataRange().getDisplayValues(), H = (v[0] || []).map(pcHead_), e = pcCol_(H, ['Enrollment#', 'Enrollment', 'Enroll#']),
    d = pcCol_(H, ['DOB', 'Date of Birth']), t = pcCol_(H, ['Portal Status', 'Student Status', 'Enrollment Status', 'Status']);
  if (e < 0) return out;
  for (var i = 1; i < v.length; i++) {
    var k = pcHead_(v[i][e]); if (!k) continue;
    var x = d >= 0 ? pcClean_(v[i][d]) : ''; if (x && !pcPlaceholderDob_(x) && !out.dob[k]) out.dob[k] = x;
    var y = t >= 0 ? pcClean_(v[i][t]) : ''; if (y && !out.status[k]) out.status[k] = y;
  }
  return out;
}
function pcPortalDob_() { return pcPortal_().dob; }
/** every status except Enrolled (and passed / promoted) means she left: Migrated, Cancelled, Disenrolled, Struck off … */
/** only a status that clearly says she left: Migrated, Cancelled, Disenrolled, Struck off, Left, Withdrawn, Expelled */
function pcLeftStatus_(s) { return /MIGRAT|CANCEL|DISENROL|STRUCK|^LEFT|WITHDR|EXPEL|RUSTICAT/.test(pcClean_(s).toUpperCase()); }
function pcStatusOf_(P, S, enroll) { var k = pcHead_(enroll); if (!k) return ''; return (P && P.status[k]) || (S.byEnroll[k] && S.byEnroll[k].status) || ''; }

/* ============================================================ list of sheets */
function pcSheets_() {
  var book = SpreadsheetApp.openById(PC_CFG.BOOK);
  return { ok: true, url: book.getUrl(), sheets: book.getSheets().map(function (s) { return s.getName(); })
    .filter(function (n) { return /^HSSC\s/i.test(n) && !/Top\s*20/i.test(n); }) };
}

/* ============================================================ ONE GO */
function pcPrepare_(name, strength) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = pcSheet_(name), C = pcCols_(sh, true), n = sh.getLastRow() - 1;
    if (n < 1) throw new Error('The sheet has no students.');
    var S = pcStrength_(strength), L = pcLeft_(), P = pcPortal_(), DOB = P.dob;
    var all = sh.getRange(2, 1, n, C.last).getDisplayValues();
    var cols = PC_CFG.FILL.map(function (t) { return sh.getRange(2, C.f[t] + 1, n, 1).getValues(); });
    var F = {}; PC_CFG.FILL.forEach(function (t, x) { F[t] = cols[x]; });
    var total = pcTotal_(sh.getName()), year = pcYear_(sh.getName()), session = (year - 1) + '-' + String(year).slice(2), longSession = (year - 1) + '-' + year;
    var used = {}, strike = [], st = { jcw: 0, left: 0, enroll: 0, name: 0, open: 0, dup: 0, dob: 0, ready: 0, fail: 0 };
    var set = function (t, i, v, keep) { if (keep && pcClean_(F[t][i][0]) !== '') return; F[t][i][0] = v; };

    // Enroll# already in use (kept)
    all.forEach(function (r, i) { if (C.college < 0 || PC_CFG.COLLEGE.test(r[C.college])) { var k = pcHead_(F['Enroll#'][i][0]); if (k) used[k] = i; } });

    // rows this tool struck before are checked again from scratch (strike removed below if she is not left);
    // rows struck by hand ("Left / no match (struck)") are kept
    var was = {};
    all.forEach(function (r, i) { if (/^LEFT STUDENT/i.test(pcClean_(F['Matched by'][i][0]))) { was[i + 2] = true; F['Matched by'][i][0] = ''; } });

    all.forEach(function (r, i) {
      if (C.college >= 0 && !PC_CFG.COLLEGE.test(r[C.college])) return;
      st.jcw++;
      var nm = pcClean_(r[C.name]), fa = pcClean_(r[C.father]), en = pcClean_(F['Enroll#'][i][0]), rg = C.reg >= 0 ? pcClean_(r[C.reg]) : '';
      if (!nm && !fa) return;
      var ps0 = pcStatusOf_(P, S, en);
      if (pcLeftStatus_(ps0)) { F['Matched by'][i][0] = 'LEFT STUDENT-' + ps0.toUpperCase(); strike.push(i + 2); st.left++; return; }
      var lf = pcLeftOf_(L, S, nm, fa, en, rg);
      if (lf) { F['Matched by'][i][0] = pcLeftNote_(lf); strike.push(i + 2); st.left++; return; }
      if (/^Left \/ no match/i.test(pcClean_(F['Matched by'][i][0]))) return;           // struck by hand: keep

      var t = en ? S.byEnroll[pcHead_(en)] : null, how = '';
      if (t) how = pcClean_(F['Matched by'][i][0]) || 'Existing Enroll# - Transform';
      if (!t && !en) {
        var list = (S.byName[pcLetters_(nm) + '|' + pcLetters_(fa)] || []).filter(function (o) { var u = used[pcHead_(o.enroll)]; return u === undefined || u === i; });
        if (list.length === 1) { t = list[0]; how = 'Name + Father Name - Transform'; }
        else if (list.length > 1) { F['Matched by'][i][0] = 'Duplicate name - choose in the app'; st.dup++; }
      }
      if (t) {
        var ps = pcStatusOf_(P, S, t.enroll);
        if (pcLeftStatus_(ps)) { F['Enroll#'][i][0] = t.enroll; F['SName'][i][0] = t.name; F['SFname'][i][0] = t.father; F['Matched by'][i][0] = 'LEFT STUDENT-' + ps.toUpperCase(); strike.push(i + 2); st.left++; return; }
        var lf2 = pcLeftOf_(L, S, nm, fa, t.enroll, rg);
        if (lf2) { F['Enroll#'][i][0] = t.enroll; F['Matched by'][i][0] = pcLeftNote_(lf2); strike.push(i + 2); st.left++; return; }
        used[pcHead_(t.enroll)] = i;
        set('Enroll#', i, t.enroll); set('SName', i, t.name); set('SFname', i, t.father); set('Program', i, t.program);
        set('Class', i, t.cls); set('Section', i, t.section); set('PCS', i, t.pcs); set('Group', i, t.group, true);
        F['Matched by'][i][0] = how;
        if (how.indexOf('Existing') === 0) st.enroll++; else st.name++;
      } else if (!pcClean_(F['Matched by'][i][0]) || /^Not found/i.test(F['Matched by'][i][0])) { F['Matched by'][i][0] = 'Not found in Transform'; st.open++; }
      else st.open++;

      var k = pcHead_(F['Enroll#'][i][0]);
      if (k && DOB[k] && !pcClean_(F['DOB'][i][0])) { F['DOB'][i][0] = DOB[k]; }
      if (pcClean_(F['DOB'][i][0])) st.dob++;

      if (!pcClean_(F['Grade'][i][0])) {                         // the gazette's own Grade is kept
        var rem = C.remarks >= 0 ? pcClean_(r[C.remarks]).toUpperCase() : '';
        F['Grade'][i][0] = PC_CFG.PASS.indexOf(rem) >= 0 ? rem
          : /FAIL|ABSENT|\bABS\b|UFM|CANCEL|COMP|SUPP|R\.?\s?L|WITHHELD|\bF\b/.test(rem) ? 'F'
          : pcGradeOf_(r[C.marks], total);
      }
      if (pcClean_(F['Session'][i][0]) === longSession) F['Session'][i][0] = session;   // 2025-2026 -> 2025-26
      set('Session', i, session, true);
      var pos = C.pos >= 0 ? Number(r[C.pos]) : 0;
      if (pos >= 1 && pos <= PC_CFG.TOP) set('Distinction_won', i, pcOrd_(pos) + ' Position in BISE Peshawar', true);
      if (PC_CFG.PASS.indexOf(pcClean_(F['Grade'][i][0]).toUpperCase()) >= 0 && pcClean_(r[C.roll])) st.ready++; else st.fail++;
    });

    PC_CFG.FILL.forEach(function (t) {
      var rg = sh.getRange(2, C.f[t] + 1, n, 1);
      if (t === 'DOB' || t === 'Enroll#') rg.setNumberFormat('@');
      rg.setValues(F[t]);
    });
    var blocks = function (list, line) {
      list.sort(function (a, b) { return a - b; });
      for (var a = 0; a < list.length;) {
        var b = a; while (b + 1 < list.length && list[b + 1] === list[b] + 1) b++;
        sh.getRange(list[a], 1, list[b] - list[a] + 1, C.last).setFontLine(line);
        a = b + 1;
      }
    };
    var isStruck = {}; strike.forEach(function (r) { isStruck[r] = true; });
    blocks(Object.keys(was).map(Number).filter(function (r) { return !isStruck[r]; }), 'none');
    blocks(strike, 'line-through');
    SpreadsheetApp.flush();
    var out = pcLoad_(name, S);
    out.stats = st; out.stats.sheets = S.sheets; out.stats.leftRead = L.count; out.stats.leftSheets = L.sheets;
    return out;
  } finally { lock.releaseLock(); }
}

/** a possible Left Students record for a row that was NOT struck, with the reason it was not struck */
function pcLeftHint_(L, S, name, father, enroll) {
  var ek = pcHead_(enroll), rec = ek && L.byEnroll[ek], why = '';
  var nr = rec ? null : pcLeftNear_(L, name, father);
  if (!rec && nr && nr.ns >= 0.8) rec = nr.rec;
  if (!rec) return '';
  var nk = pcLetters_(name) + '|' + pcLetters_(father);
  if ((S.cur.byName[nk] || []).length || (ek && S.cur.byEnroll[ek])) why = 'she is still in the current student list';
  else if (nr) why = 'names only ' + Math.round(nr.ns * 100) + '% / father ' + Math.round(nr.fs * 100) + '% alike';
  else why = 'press Prepare again';
  return rec.name + ' D/o ' + rec.father + (rec.enroll ? ' · ' + rec.enroll : '') + (rec.status ? ' · ' + rec.status : '') + ' — not struck: ' + why;
}

/* ============================================================ data for the app */
function pcLoad_(name, S, strength) {
  S = S || pcStrength_(strength);
  var P = pcPortal_(), L = pcLeft_();
  var sh = pcSheet_(name), C = pcCols_(sh, true), n = sh.getLastRow() - 1, rows = [];
  if (n > 0) {
    var v = sh.getRange(2, 1, n, C.last).getDisplayValues(), lines = sh.getRange(2, C.name + 1, n, 1).getFontLines();
    var g = function (r, k) { return k >= 0 ? pcClean_(r[k]) : ''; };
    v.forEach(function (r, i) {
      if (C.college >= 0 && pcClean_(r[C.college]) && !PC_CFG.COLLEGE.test(r[C.college])) return;
      if (!g(r, C.name) && !g(r, C.roll)) return;
      rows.push({ row: i + 2, roll: g(r, C.roll), name: g(r, C.name), father: g(r, C.father), reg: g(r, C.reg), marks: g(r, C.marks),
        remarks: g(r, C.remarks), pos: g(r, C.pos), enroll: g(r, C.f['Enroll#']), sname: g(r, C.f['SName']), program: g(r, C.f['Program']),
        cls: g(r, C.f['Class']), section: g(r, C.f['Section']), pcs: g(r, C.f['PCS']), group: g(r, C.f['Group']), session: g(r, C.f['Session']),
        dist: g(r, C.f['Distinction_won']), dob: g(r, C.f['DOB']), grade: g(r, C.f['Grade']), how: g(r, C.f['Matched by']),
        total: g(r, C.f['Total Marks']), conduct: g(r, C.f['Conduct']), struck: lines[i][0] === 'line-through',
        pstatus: pcStatusOf_(P, S, g(r, C.f['Enroll#'])), leftHint: lines[i][0] === 'line-through' ? '' : pcLeftHint_(L, S, g(r, C.name), g(r, C.father), g(r, C.f['Enroll#'])) });
    });
  }
  return { ok: true, sheet: sh.getName(), url: SpreadsheetApp.openById(PC_CFG.BOOK).getUrl() + '#gid=' + sh.getSheetId(),
    total: pcTotal_(sh.getName()), rows: rows,
    source: S.rows.map(function (o) { return [o.enroll, o.name, o.father, o.cls, o.pcs]; }) };
}

/* ============================================================ link rows to Transform students (manual / 96%+) */
function pcLink_(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { return pcLinkCore_(req); } finally { lock.releaseLock(); }
}
function pcLinkCore_(req) {
    var sh = pcSheet_(req.sheet), C = pcCols_(sh, true), n = sh.getLastRow() - 1, S = pcStrength_(req.strength), DOB = pcPortalDob_();
    var names = sh.getRange(2, C.name + 1, n, 1).getDisplayValues(), fathers = sh.getRange(2, C.father + 1, n, 1).getDisplayValues();
    var F = {}; PC_CFG.FILL.forEach(function (t) { F[t] = sh.getRange(2, C.f[t] + 1, n, 1).getValues(); });
    var used = {}; F['Enroll#'].forEach(function (x, i) { var k = pcHead_(x[0]); if (k) used[k] = i + 2; });
    var done = [], skipped = [];
    (req.items || []).forEach(function (it) {
      var at = pcLocate_(sh, C, it.row, it.roll, it.name, it.father), i = at - 2;
      if (!at) { skipped.push({ row: it.row, msg: (it.roll || it.name) + ' not found in the sheet any more' }); return; }
      it.row = at;
      if (!it.enroll) {                                           // unlink
        ['Enroll#', 'SName', 'SFname', 'Program', 'Class', 'Section', 'PCS'].forEach(function (t) { F[t][i][0] = ''; });
        F['Matched by'][i][0] = 'Not found in Transform'; done.push(Number(it.row)); return;
      }
      var t = S.byEnroll[pcHead_(it.enroll)];
      if (!t) { skipped.push({ row: it.row, msg: it.enroll + ' is not in ' + S.sheets.join(' / ') }); return; }
      var other = used[pcHead_(it.enroll)];
      if (other && other !== Number(it.row)) { skipped.push({ row: it.row, msg: it.enroll + ' is already used in row ' + other }); return; }
      var old = pcHead_(F['Enroll#'][i][0]); if (old && used[old] === Number(it.row)) delete used[old];
      F['Enroll#'][i][0] = t.enroll; F['SName'][i][0] = t.name; F['SFname'][i][0] = t.father; F['Program'][i][0] = t.program;
      F['Class'][i][0] = t.cls; F['Section'][i][0] = t.section; F['PCS'][i][0] = t.pcs;
      if (!pcClean_(F['Group'][i][0])) F['Group'][i][0] = t.group;
      if (!pcClean_(F['DOB'][i][0]) && DOB[pcHead_(t.enroll)]) F['DOB'][i][0] = DOB[pcHead_(t.enroll)];
      F['Matched by'][i][0] = it.by || 'Chosen by hand (app)';
      used[pcHead_(t.enroll)] = Number(it.row);
      done.push(Number(it.row));
    });
    if (done.length) {
      PC_CFG.FILL.forEach(function (t) { var rg = sh.getRange(2, C.f[t] + 1, n, 1); if (t === 'DOB' || t === 'Enroll#') rg.setNumberFormat('@'); rg.setValues(F[t]); });
      SpreadsheetApp.flush();
    }
    return { ok: true, done: done, skipped: skipped, rows: done.length ? pcLoad_(req.sheet, S).rows.filter(function (r) { return done.indexOf(r.row) >= 0; }) : [] };
}

/* ============================================================ edit one student (or add one not in the gazette) */
var PC_EDIT = { name: 'name', father: 'father', roll: 'roll', reg: 'reg', marks: 'marks',
  grade: 'Grade', dob: 'DOB', group: 'Group', session: 'Session', dist: 'Distinction_won', total: 'Total Marks', conduct: 'Conduct' };
function pcSave_(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = pcSheet_(req.sheet), C = pcCols_(sh, true), v = req.values || {}, row = Number(req.row) || 0;
    if (!pcClean_(v.name) || !pcClean_(v.father)) return { ok: false, error: 'Name and Father name are needed.' };
    var asked = row;
    if (row) {
      var X = req.expect || {};
      row = pcLocate_(sh, C, row, X.roll, X.name, X.father);
      if (!row) return { ok: false, error: 'Roll No ' + (X.roll || '-') + ' / ' + X.name + ' was not found in "' + sh.getName() + '" any more. Press Reload.' };
    }
    var roll = pcClean_(v.roll);
    if (roll && C.roll >= 0) {
      var rolls = sh.getRange(2, C.roll + 1, Math.max(1, sh.getLastRow() - 1), 1).getDisplayValues();
      for (var i = 0; i < rolls.length; i++) if (pcClean_(rolls[i][0]) === roll && i + 2 !== row) return { ok: false, error: 'Roll No ' + roll + ' is already in row ' + (i + 2) + '.' };
    }
    if (!row) {
      row = sh.getLastRow() + 1;
      if (C.college >= 0) sh.getRange(row, C.college + 1).setValue('JINNAH COLLEGE FOR WOMEN UNIVERSITY OF PESHAWAR');
      sh.getRange(row, C.f['Matched by'] + 1).setValue('Added in the app');
    }
    Object.keys(PC_EDIT).forEach(function (k) {
      if (!(k in v)) return;
      var val = pcClean_(v[k]), t = PC_EDIT[k], col;
      if (k === 'conduct' && (val === '' || val.toUpperCase() === PC_CFG.DEF.conduct)) val = '';
      if (k === 'dist' && val.toUpperCase() === 'NIL') val = '';
      if (typeof C[t] === 'number') col = C[t];                // gazette column
      else col = C.f[t] >= 0 ? C.f[t] : (val ? C.addExtra(t) : -1);
      if (col < 0) return;
      var cell = sh.getRange(row, col + 1);
      if (k === 'roll' || k === 'reg' || k === 'dob') cell.setNumberFormat('@');
      cell.setValue(k === 'marks' && /^\d+(\.\d+)?$/.test(val) ? Number(val) : val);
    });
    if (v.enroll !== undefined) {
      SpreadsheetApp.flush();
      var r2 = pcLinkCore_({ sheet: req.sheet, strength: req.strength, items: [{ row: row, roll: v.roll || (req.expect && req.expect.roll), name: v.name, father: v.father, enroll: pcClean_(v.enroll), by: 'Chosen by hand (app)' }] });
      if (r2.skipped && r2.skipped.length && pcClean_(v.enroll)) return { ok: false, error: 'Saved, but Enroll#: ' + r2.skipped[0].msg };
    }
    SpreadsheetApp.flush();
    var one = pcLoad_(req.sheet, null, req.strength).rows.filter(function (r) { return r.row === row; })[0];
    return { ok: true, row: one, moved: !!asked && asked !== row };
  } finally { lock.releaseLock(); }
}

/** strike / un-strike a row (no certificate) */
function pcStrike_(req) {
  var sh = pcSheet_(req.sheet), C = pcCols_(sh, true), row = pcLocate_(sh, C, req.row, req.roll, req.name, req.father);
  if (!row) return { ok: false, error: (req.roll || req.name) + ' was not found in the sheet any more. Press Reload.' };
  var v = sh.getRange(row, 1, 1, C.last).getDisplayValues()[0];
  sh.getRange(row, 1, 1, C.last).setFontLine(req.on ? 'line-through' : 'none');
  var hc = sh.getRange(row, C.f['Matched by'] + 1), cur = pcClean_(hc.getDisplayValue());
  if (req.on) hc.setValue('Left / no match (struck)'); else if (/^Left \/ no match|^LEFT STUDENT/.test(cur)) hc.setValue(pcClean_(v[C.f['Enroll#']]) ? 'Chosen by hand (app)' : 'Not found in Transform');
  return { ok: true };
}

/* ============================================================ pictures + PDF */
function pcImg_(name) {
  try { var it = DriveApp.getFilesByName(name); if (!it.hasNext()) return ''; var b = it.next().getBlob(); return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); }
  catch (e) { return ''; }
}
function pcAssets_() {
  var c = CacheService.getScriptCache(), k = 'pc_assets_v1', hit = c.get(k);
  if (hit) return JSON.parse(hit);
  var out = { ok: true, logo: pcImg_(PC_CFG.LOGO), sign: pcImg_(PC_CFG.SIGN_IMG), cfg: { title: PC_CFG.TITLE, college: PC_CFG.COLLEGE_NAME, university: PC_CFG.UNIVERSITY, gazette: PC_CFG.GAZETTE, sign: PC_CFG.SIGN } };
  try { var s = JSON.stringify(out); if (s.length < 95000) c.put(k, s, 21600); } catch (e) {}
  return out;
}
function pcEsc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function pcRowHtml_(parts) {
  return '<table style="width:100%;border-collapse:collapse;margin-top:16px;"><tr>' + parts.map(function (p) {
    return '<td style="white-space:nowrap;padding:0 6px 0 0;">' + pcEsc_(p[0]) + '</td><td style="width:' + p[2] + '%;border-bottom:1px solid #000;text-align:center;font-weight:bold;">' + (pcEsc_(p[1]) || '&nbsp;') + '</td>';
  }).join('<td style="width:10px"></td>') + '</tr></table>';
}
function pcPage_(v, logo, sign, first) {
  var head = logo ? '<div style="text-align:center;"><img width="470" src="' + logo + '"></div>' : '<div style="text-align:center;font-size:30px;font-weight:bold;color:#1F4E8C;">' + PC_CFG.COLLEGE_NAME + '</div>';
  return '<div style="padding:20px 40px;color:#000;' + (first ? '' : 'page-break-before:always;') + '">' + head +
    '<div style="text-align:center;font-size:20px;font-weight:bold;color:#1F4E8C;margin-top:4px;">' + PC_CFG.UNIVERSITY + '</div>' +
    '<div style="text-align:center;font-size:14px;font-weight:bold;text-decoration:underline;margin:22px 0 26px;">' + PC_CFG.TITLE + '</div>' +
    pcRowHtml_([['This is to certify that Miss.', v.name, 70]]) + pcRowHtml_([['Daughter of Mr.', v.father, 70]]) +
    pcRowHtml_([['Appeared in the', v.group, 30], ['Examination held in', v.session, 25]]) +
    '<div style="margin-top:18px;line-height:1.5;">' + PC_CFG.GAZETTE + '</div><div style="margin-top:14px;">Other Particulars are: -</div>' +
    pcRowHtml_([['Roll No.', v.roll, 30], ['Registration No.', v.reg, 40]]) +
    pcRowHtml_([['Marks obtained', v.marks, 15], ['out of', v.total, 15], ['Grade/Division', v.grade, 15]]) +
    pcRowHtml_([['Date of Birth', v.dob, 70]]) + pcRowHtml_([['Distinctions Won', v.dist, 70]]) + pcRowHtml_([['Conduct and Moral Character', v.conduct, 55]]) +
    '<div style="margin-top:40px;text-align:right;"><div style="display:inline-block;width:230px;text-align:center;"><div style="height:60px;">' +
    (sign ? '<img height="55" src="' + sign + '">' : '') + '</div>' + PC_CFG.SIGN.map(pcEsc_).join('<br>') + '</div></div></div>';
}
/** { sheet, rows:[row numbers] (empty = every ready student) } -> one PDF in Drive */
function pcPdf_(req) {
  var d = pcLoad_(req.sheet, null, req.strength), want = (req.rows || []).map(Number), A = pcAssets_();
  var list = d.rows.filter(function (r) {
    if (want.length) return want.indexOf(r.row) >= 0;
    return !r.struck && r.roll && PC_CFG.PASS.indexOf(String(r.grade).toUpperCase()) >= 0;
  }).map(function (r) {
    return { name: r.name, father: r.father, group: r.group || (pcPart_(d.sheet) === 2 ? 'H.S.S.C (Part-II)' : 'H.S.S.C (Part-I)'), session: r.session,
      roll: r.roll, reg: r.reg, marks: r.marks, total: r.total || d.total, grade: r.grade, dob: r.dob, dist: r.dist || PC_CFG.DEF.distinctions, conduct: r.conduct || PC_CFG.DEF.conduct };
  });
  if (!list.length) return { ok: false, error: 'No students to print.' };
  list.sort(function (a, b) { var x = Number(a.roll), y = Number(b.roll); return (isNaN(x) || isNaN(y)) ? String(a.roll).localeCompare(String(b.roll)) : x - y; });
  var html = '<html><head><meta charset="utf-8"></head><body style="font-family:\'Times New Roman\',serif;font-size:15px;margin:0;">' +
    list.map(function (v, i) { return pcPage_(v, A.logo, A.sign, i === 0); }).join('') + '</body></html>';
  var tz = Session.getScriptTimeZone(), nm = list.length === 1 ? pcClean_(list[0].roll + ' - ' + list[0].name).replace(/[\\/:*?"<>|]/g, '') : 'Provisional Certificates ' + d.sheet + ' ' + Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HHmm');
  var it = DriveApp.getFoldersByName(PC_CFG.FOLDER), folder = it.hasNext() ? it.next() : DriveApp.createFolder(PC_CFG.FOLDER);
  var file = folder.createFile(Utilities.newBlob(html, 'text/html', nm + '.html').getAs('application/pdf').setName(nm + '.pdf'));
  return { ok: true, url: file.getUrl(), id: file.getId(), n: list.length };
}
