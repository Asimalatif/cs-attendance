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
function pcStrength_(name) {
  name = pcClean_(name) || PC_CFG.STRENGTH;
  var sh = SpreadsheetApp.openById(PC_CFG.USERS).getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found in the Users workbook.');
  var v = sh.getDataRange().getDisplayValues(), H = (v[0] || []).map(pcHead_);
  var c = { e: pcCol_(H, ['Enroll#', 'Enrollment#', 'Enrollment No', 'REG#']), n: pcCol_(H, ['Student Name', 'Name']), f: pcCol_(H, ['Father Name', 'Father']),
    p: pcCol_(H, ['Program', 'Program Name']), c: pcCol_(H, ['Class']), s: pcCol_(H, ['Section']), x: pcCol_(H, ['PCS']), a: pcCol_(H, ['Appeared in']) };
  if (c.e < 0 || c.n < 0) throw new Error(PC_CFG.STRENGTH + ' needs Enroll# and Student Name columns.');
  var g = function (r, i) { return i >= 0 ? pcClean_(r[i]) : ''; };
  var S = { rows: [], byEnroll: {}, byName: {} };
  for (var i = 1; i < v.length; i++) {
    var r = v[i], o = { enroll: g(r, c.e), name: g(r, c.n), father: g(r, c.f), program: g(r, c.p), cls: g(r, c.c), section: g(r, c.s), pcs: g(r, c.x), group: g(r, c.a) };
    var k = pcHead_(o.enroll); if (!k || S.byEnroll[k]) continue;
    S.rows.push(o); S.byEnroll[k] = o;
    var nk = pcLetters_(o.name) + '|' + pcLetters_(o.father); if (nk !== '|') (S.byName[nk] = S.byName[nk] || []).push(o);
  }
  return S;
}
function pcLeft_() {
  var sh = SpreadsheetApp.openById(PC_CFG.USERS).getSheetByName(PC_CFG.LEFT), out = { byEnroll: {}, byReg: {}, byName: {} };
  if (!sh || sh.getLastRow() < 2) return out;
  var v = sh.getDataRange().getDisplayValues(), at = -1, ix = null;
  for (var h = 0; h < Math.min(10, v.length); h++) {             // the header may sit on row 2
    var H = v[h].map(pcHead_), x = { e: pcCol_(H, ['Enroll#', 'Enrollment#', 'Enrollment No']), n: pcCol_(H, ['Student Name', 'Name']), f: pcCol_(H, ['Father Name', 'Father']),
      r: pcCol_(H, ['REG#', 'Registration #', 'Registration No']), st: pcCol_(H, ['Portal Status', 'Status']) };
    if (x.n >= 0 && x.f >= 0 && (x.e >= 0 || x.r >= 0)) { at = h; ix = x; break; }
  }
  if (at < 0) return out;
  for (var i = at + 1; i < v.length; i++) {
    var r = v[i], g = function (k) { return k >= 0 ? pcClean_(r[k]) : ''; };
    var rec = { enroll: g(ix.e), reg: g(ix.r), name: g(ix.n), father: g(ix.f), status: g(ix.st) };
    if (rec.enroll) out.byEnroll[pcHead_(rec.enroll)] = out.byEnroll[pcHead_(rec.enroll)] || rec;
    if (rec.reg) out.byReg[pcHead_(rec.reg)] = out.byReg[pcHead_(rec.reg)] || rec;
    if (rec.name && rec.father) out.byName[pcLetters_(rec.name) + '|' + pcLetters_(rec.father)] = rec;
  }
  return out;
}
function pcLeftOf_(L, S, name, father, enroll, reg) {
  var ek = pcHead_(enroll), rk = pcHead_(reg), nk = pcLetters_(name) + '|' + pcLetters_(father);
  if (ek && L.byEnroll[ek] && !S.byEnroll[ek]) return L.byEnroll[ek];
  if (rk && L.byReg[rk]) return L.byReg[rk];
  if (name && father && L.byName[nk] && !(S.byName[nk] || []).length) return L.byName[nk];   // still studying -> not left
  return null;
}
function pcLeftNote_(rec) { return 'LEFT STUDENT-' + (pcClean_(rec.status).toUpperCase() || 'LEFT'); }
function pcPortalDob_() {
  var sh = SpreadsheetApp.openById(PC_CFG.USERS).getSheetByName(PC_CFG.PORTAL), out = {};
  if (!sh) return out;
  var v = sh.getDataRange().getDisplayValues(), H = (v[0] || []).map(pcHead_), e = pcCol_(H, ['Enrollment#', 'Enrollment', 'Enroll#']), d = pcCol_(H, ['DOB', 'Date of Birth']);
  if (e < 0 || d < 0) return out;
  for (var i = 1; i < v.length; i++) { var k = pcHead_(v[i][e]), x = pcClean_(v[i][d]); if (k && x && !pcPlaceholderDob_(x) && !out[k]) out[k] = x; }
  return out;
}

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
    var S = pcStrength_(strength), L = pcLeft_(), DOB = pcPortalDob_();
    var all = sh.getRange(2, 1, n, C.last).getDisplayValues();
    var cols = PC_CFG.FILL.map(function (t) { return sh.getRange(2, C.f[t] + 1, n, 1).getValues(); });
    var F = {}; PC_CFG.FILL.forEach(function (t, x) { F[t] = cols[x]; });
    var total = pcTotal_(sh.getName()), year = pcYear_(sh.getName()), session = (year - 1) + '-' + year;
    var used = {}, strike = [], st = { jcw: 0, left: 0, enroll: 0, name: 0, open: 0, dup: 0, dob: 0, ready: 0, fail: 0 };
    var set = function (t, i, v, keep) { if (keep && pcClean_(F[t][i][0]) !== '') return; F[t][i][0] = v; };

    // Enroll# already in use (kept)
    all.forEach(function (r, i) { if (C.college < 0 || PC_CFG.COLLEGE.test(r[C.college])) { var k = pcHead_(F['Enroll#'][i][0]); if (k) used[k] = i; } });

    all.forEach(function (r, i) {
      if (C.college >= 0 && !PC_CFG.COLLEGE.test(r[C.college])) return;
      st.jcw++;
      var nm = pcClean_(r[C.name]), fa = pcClean_(r[C.father]), en = pcClean_(F['Enroll#'][i][0]), rg = C.reg >= 0 ? pcClean_(r[C.reg]) : '';
      if (!nm && !fa) return;
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
    strike.sort(function (a, b) { return a - b; });
    for (var a = 0; a < strike.length;) {
      var b = a; while (b + 1 < strike.length && strike[b + 1] === strike[b] + 1) b++;
      sh.getRange(strike[a], 1, strike[b] - strike[a] + 1, C.last).setFontLine('line-through');
      a = b + 1;
    }
    SpreadsheetApp.flush();
    var out = pcLoad_(name, S);
    out.stats = st;
    return out;
  } finally { lock.releaseLock(); }
}

/* ============================================================ data for the app */
function pcLoad_(name, S, strength) {
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
        total: g(r, C.f['Total Marks']), conduct: g(r, C.f['Conduct']), struck: lines[i][0] === 'line-through' });
    });
  }
  S = S || pcStrength_(strength);
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
      var i = Number(it.row) - 2;
      if (!(i >= 0 && i < n)) { skipped.push({ row: it.row, msg: 'row not found' }); return; }
      if (pcClean_(names[i][0]) !== pcClean_(it.name) || pcClean_(fathers[i][0]) !== pcClean_(it.father)) { skipped.push({ row: it.row, msg: 'row changed; reload' }); return; }
      if (!it.enroll) {                                           // unlink
        ['Enroll#', 'SName', 'SFname', 'Program', 'Class', 'Section', 'PCS'].forEach(function (t) { F[t][i][0] = ''; });
        F['Matched by'][i][0] = 'Not found in Transform'; done.push(Number(it.row)); return;
      }
      var t = S.byEnroll[pcHead_(it.enroll)];
      if (!t) { skipped.push({ row: it.row, msg: it.enroll + ' is not in ' + PC_CFG.STRENGTH }); return; }
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
    if (row) {
      var cur = sh.getRange(row, 1, 1, C.last).getDisplayValues()[0];
      if (pcClean_(cur[C.name]) !== pcClean_(req.expect && req.expect.name) || pcClean_(cur[C.father]) !== pcClean_(req.expect && req.expect.father))
        return { ok: false, error: 'Row ' + row + ' has changed since it was loaded. Reload the list.' };
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
      var r2 = pcLinkCore_({ sheet: req.sheet, strength: req.strength, items: [{ row: row, name: v.name, father: v.father, enroll: pcClean_(v.enroll), by: 'Chosen by hand (app)' }] });
      if (r2.skipped && r2.skipped.length && pcClean_(v.enroll)) return { ok: false, error: 'Saved, but Enroll#: ' + r2.skipped[0].msg };
    }
    SpreadsheetApp.flush();
    var one = pcLoad_(req.sheet, null, req.strength).rows.filter(function (r) { return r.row === row; })[0];
    return { ok: true, row: one };
  } finally { lock.releaseLock(); }
}

/** strike / un-strike a row (no certificate) */
function pcStrike_(req) {
  var sh = pcSheet_(req.sheet), C = pcCols_(sh, true), row = Number(req.row);
  if (!row || row < 2 || row > sh.getLastRow()) return { ok: false, error: 'Row not found.' };
  var v = sh.getRange(row, 1, 1, C.last).getDisplayValues()[0];
  if (pcClean_(v[C.name]) !== pcClean_(req.name) || pcClean_(v[C.father]) !== pcClean_(req.father)) return { ok: false, error: 'Row changed; reload.' };
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
