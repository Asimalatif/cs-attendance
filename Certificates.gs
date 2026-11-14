/**
 * Certificates.gs — Provisional cum Character Certificates inside the attendance app (admin only).
 *
 * ONE GO (pc_prepare), Jinnah College rows only (other colleges are never written). Only Student Name + Father Name
 * are compared, never Enroll# or Registration #:
 *   1. Left Students  -> the same name + father (and not in the current Transform sheet): row struck through
 *   2. Transform-ALL-COLLEGE-STRENGTH (+ older copies) -> 100% same name + father: Enroll#, SName, SFname, Program,
 *                        Class, Section, PCS, Group are filled. Everything else is listed in the app; the admin ticks
 *                        the right students and they are written in ONE go (pc_link).
 *   3. Portal backup  -> DOB (by name + father when Portal backup has those columns, else by the linked student)
 *   4. Grade (gazette, Remarks, or from Marks) and Session (2025-26). Distinction_won is typed by the admin.
 * Speed: the sheet is read once and only changed cells are written. Turn on Services > Google Sheets API in the
 * Apps Script editor and all changes go to the sheet in one request.
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
      case 'pc_strike':  return pcStrikes_({ sheet: req.sheet, items: [{ row: req.row, roll: req.roll, name: req.name, father: req.father, on: req.on !== false }], on: req.on !== false });
      case 'pc_strikes': return pcStrikes_(req);
      case 'pc_assets':  return pcAssets_();
      case 'pc_pdf':     return pcPdf_(req);
      case 'pc_pdf_get': return pcPdfGet_(req);
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

/* ============================================================ the result sheet in memory (one read, few writes) */
/** columns of an HSSC sheet (0-based); FILL columns are added at the right when missing */
function pcCols_(sh) {
  var lastCol = sh.getLastColumn(), H = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(pcHead_);
  var add = function (title) {
    var c = H.indexOf(pcHead_(title));
    if (c < 0) { lastCol++; sh.getRange(1, lastCol).setValue(title).setFontWeight('bold').setBackground('#E7B7AE'); c = lastCol - 1; H[c] = pcHead_(title); }
    return c;
  };
  var C = {
    roll: pcCol_(H, ['Roll No', 'RollNo']), name: pcCol_(H, ['Name']), father: pcCol_(H, ['Father Name']),
    reg: pcCol_(H, ['Registration #', 'Registration', 'Registration No']), marks: pcCol_(H, ['Marks']),
    remarks: pcCol_(H, ['Remarks']), college: pcCol_(H, ['College Name']), pos: pcCol_(H, ['Position'])
  };
  if (C.name < 0 || C.father < 0) throw new Error('"' + sh.getName() + '" needs Name and Father Name columns.');
  C.f = {}; PC_CFG.FILL.forEach(function (t) { C.f[t] = add(t); });
  PC_CFG.EXTRA.forEach(function (t) { C.f[t] = H.indexOf(pcHead_(t)); });
  C.addExtra = function (t) { if (C.f[t] < 0) { C.f[t] = add(t); C.last = lastCol; } return C.f[t]; };
  C.last = lastCol;
  return C;
}
/** reads the whole sheet ONCE: values of every column + strike state of the Name column */
function pcCtx_(name) {
  var sh = pcSheet_(name), C = pcCols_(sh), n = Math.max(0, sh.getLastRow() - 1);
  var V = n ? sh.getRange(2, 1, n, C.last).getDisplayValues() : [];
  var lines = n ? sh.getRange(2, C.name + 1, n, 1).getFontLines() : [];
  var jcw = [];
  for (var i = 0; i < n; i++) {
    var col = C.college >= 0 ? pcClean_(V[i][C.college]) : '';
    if (col && !PC_CFG.COLLEGE.test(col)) continue;                         // other colleges: never touched
    if (!pcClean_(V[i][C.name]) && !(C.roll >= 0 && pcClean_(V[i][C.roll]))) continue;
    jcw.push(i);
  }
  return { sh: sh, C: C, n: n, V: V, struck: lines.map(function (x) { return x[0] === 'line-through'; }), jcw: jcw, changed: {}, strikeOn: [], strikeOff: [] };
}
function pcKey_(name, father) { return pcLetters_(name) + '|' + pcLetters_(father); }
/** set one cell in memory and remember it for the write */
function pcSet_(X, i, col, val) {
  if (col < 0) return;
  val = val == null ? '' : String(val);
  while (X.V[i].length <= col) X.V[i].push('');
  if (pcClean_(X.V[i][col]) === pcClean_(val)) return;
  X.V[i][col] = val;
  (X.changed[i] = X.changed[i] || {})[col] = true;
}
function pcF_(X, i, t) { var c = X.C.f[t]; return c >= 0 ? pcClean_(X.V[i][c]) : ''; }
function pcSetF_(X, i, t, v) { pcSet_(X, i, X.C.f[t], v); }

/** writes only the changed cells. With the Google Sheets advanced service (Services > Google Sheets API) every
 *  change goes in ONE request; without it, changed cells are written in as few calls as possible. */
function pcFlush_(X) {
  var sh = X.sh, rows = Object.keys(X.changed).map(Number).sort(function (a, b) { return a - b; });
  var A1 = function (r, c) { var s = '', n = c + 1; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s + (r + 2); };
  var runs = [];
  rows.forEach(function (i) {
    var cols = Object.keys(X.changed[i]).map(Number).sort(function (a, b) { return a - b; });
    for (var a = 0; a < cols.length;) { var b = a; while (b + 1 < cols.length && cols[b + 1] === cols[b] + 1) b++; runs.push({ i: i, c0: cols[a], c1: cols[b] }); a = b + 1; }
  });
  if (runs.length) {
    var vals = function (r) { var out = []; for (var c = r.c0; c <= r.c1; c++) out.push(X.V[r.i][c] == null ? '' : String(X.V[r.i][c])); return out; };
    var done = false;
    if (runs.length > 20 && typeof Sheets !== 'undefined') {
      try {
        var q = "'" + sh.getName().replace(/'/g, "''") + "'!";
        for (var k = 0; k < runs.length; k += 3000) {
          Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'RAW', data: runs.slice(k, k + 3000).map(function (r) {
            return { range: q + A1(r.i, r.c0) + ':' + A1(r.i, r.c1), values: [vals(r)] };
          }) }, PC_CFG.BOOK);
        }
        done = true;
      } catch (e) { Logger.log('Sheets API write failed, using normal writes: ' + e); }
    }
    if (!done) {
      var textCols = [X.C.f['DOB'], X.C.f['Enroll#'], X.C.roll, X.C.reg];
      if (runs.length > 300) {                                              // many rows: one block per changed column
        var cols = {}; runs.forEach(function (r) { for (var c = r.c0; c <= r.c1; c++) cols[c] = true; });
        Object.keys(cols).map(Number).forEach(function (c) {
          var rg = sh.getRange(2, c + 1, X.n, 1);
          if (textCols.indexOf(c) >= 0) rg.setNumberFormat('@');
          rg.setValues(X.V.map(function (r) { return [r[c] == null ? '' : r[c]]; }));
        });
      } else {
        runs.forEach(function (r) {
          var rg = sh.getRange(r.i + 2, r.c0 + 1, 1, r.c1 - r.c0 + 1);
          if (textCols.some(function (c) { return c >= r.c0 && c <= r.c1; })) rg.setNumberFormat('@');
          rg.setValues([vals(r)]);
        });
      }
    }
  }
  var lines = function (list, kind) {
    if (!list.length) return;
    var a1 = list.map(function (i) { return 'A' + (i + 2) + ':' + A1(i, X.C.last - 1).replace(/\d+$/, '') + (i + 2); });
    for (var k = 0; k < a1.length; k += 500) sh.getRangeList(a1.slice(k, k + 500)).setFontLine(kind);
  };
  lines(X.strikeOff, 'none'); lines(X.strikeOn, 'line-through');
  X.strikeOff.forEach(function (i) { X.struck[i] = false; }); X.strikeOn.forEach(function (i) { X.struck[i] = true; });
  X.changed = {}; X.strikeOn = []; X.strikeOff = [];
  SpreadsheetApp.flush();
}
/** row object for the app */
function pcRowOf_(X, i) {
  var V = X.V[i], C = X.C, g = function (k) { return k >= 0 ? pcClean_(V[k]) : ''; }, f = function (t) { return pcF_(X, i, t); };
  return { row: i + 2, roll: g(C.roll), name: g(C.name), father: g(C.father), reg: g(C.reg), marks: g(C.marks), remarks: g(C.remarks), pos: g(C.pos),
    enroll: f('Enroll#'), sname: f('SName'), sfather: f('SFname'), program: f('Program'), cls: f('Class'), section: f('Section'), pcs: f('PCS'),
    group: f('Group'), session: f('Session'), dist: f('Distinction_won'), dob: f('DOB'), grade: f('Grade'), how: f('Matched by'),
    total: f('Total Marks'), conduct: f('Conduct'), struck: !!X.struck[i] };
}
/** a row by number; if rows moved (sorting), by Roll No, else by unique Name + Father Name */
function pcFind_(X, row, roll, name, father) {
  var i = Number(row) - 2, k = pcKey_(name, father), C = X.C, R = function (j) { return C.roll >= 0 ? pcClean_(X.V[j][C.roll]) : ''; };
  if (i >= 0 && i < X.n && (roll ? R(i) === pcClean_(roll) : pcKey_(X.V[i][C.name], X.V[i][C.father]) === k)) return i;
  var hit = X.jcw.filter(function (j) { return roll && R(j) === pcClean_(roll); });
  if (hit.length === 1) return hit[0];
  hit = X.jcw.filter(function (j) { return pcKey_(X.V[j][C.name], X.V[j][C.father]) === k; });
  return hit.length === 1 ? hit[0] : -1;
}

/* ============================================================ sources (Users workbook) */
/** Students from Transform-ALL-COLLEGE-STRENGTH and its older copies ("Transform-ALL-COLLEGE-STRENGTH 2025" …).
 *  name = one sheet only (the box on the page). Only Student Name + Father Name are compared. */
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
  var S = { rows: [], byId: {}, byName: {}, cur: {}, sheets: names };
  names.forEach(function (sn, si) {
    var v = book.getSheetByName(sn).getDataRange().getDisplayValues(), H = (v[0] || []).map(pcHead_);
    var c = { e: pcCol_(H, ['Enroll#', 'Enrollment#', 'Enrollment No', 'REG#']), n: pcCol_(H, ['Student Name', 'Name']), f: pcCol_(H, ['Father Name', 'Father']),
      p: pcCol_(H, ['Program', 'Program Name']), c: pcCol_(H, ['Class']), s: pcCol_(H, ['Section']), x: pcCol_(H, ['PCS']), a: pcCol_(H, ['Appeared in']),
      d: pcCol_(H, ['DOB', 'Date of Birth']) };
    if (c.n < 0) { if (si === 0) throw new Error(sn + ' needs a Student Name column.'); return; }
    var g = function (r, i) { return i >= 0 ? pcClean_(r[i]) : ''; };
    for (var i = 1; i < v.length; i++) {
      var r = v[i], o = { enroll: g(r, c.e), name: g(r, c.n), father: g(r, c.f), program: g(r, c.p), cls: g(r, c.c), section: g(r, c.s), pcs: g(r, c.x), group: g(r, c.a), dob: g(r, c.d) };
      if (!o.name) continue;
      var id = o.enroll || (sn + '#' + i);                                 // only an id for the app's choice list
      if (S.byId[pcHead_(id)]) continue;                                   // the current sheet wins
      o.id = id; S.rows.push(o); S.byId[pcHead_(id)] = o;
      var nk = pcKey_(o.name, o.father);
      (S.byName[nk] = S.byName[nk] || []).push(o);
      if (si === 0) S.cur[nk] = true;
    }
  });
  return S;
}
/** every sheet whose name starts with "Left Students"; only Student Name + Father Name are used */
function pcLeft_() {
  var book = SpreadsheetApp.openById(PC_CFG.USERS), base = pcHead_(PC_CFG.LEFT), out = { byName: {}, list: [], count: 0, sheets: [] };
  book.getSheets().forEach(function (sh) {
    if (pcHead_(sh.getName()).indexOf(base) !== 0 || sh.getLastRow() < 2) return;
    var v = sh.getDataRange().getDisplayValues(), at = -1, ix = null;
    for (var h = 0; h < Math.min(10, v.length); h++) {
      var H = v[h].map(pcHead_), x = { n: pcCol_(H, ['Student Name', 'Name']), f: pcCol_(H, ['Father Name', 'Father']), e: pcCol_(H, ['Enroll#', 'Enrollment#']), st: pcCol_(H, ['Portal Status', 'Status']) };
      if (x.n >= 0 && x.f >= 0) { at = h; ix = x; break; }
    }
    if (at < 0) return;
    out.sheets.push(sh.getName());
    for (var i = at + 1; i < v.length; i++) {
      var r = v[i], g = function (k) { return k >= 0 ? pcClean_(r[k]) : ''; };
      var rec = { name: g(ix.n), father: g(ix.f), enroll: g(ix.e), status: g(ix.st) };
      if (!rec.name) continue;
      out.count++; out.list.push(rec);
      if (rec.father) out.byName[pcKey_(rec.name, rec.father)] = rec;
    }
  });
  return out;
}
/** Left student = same Student Name + Father Name in Left Students, and not the same name + father in the current Transform */
function pcLeftOf_(L, S, name, father) {
  if (!name || !father) return null;
  var k = pcKey_(name, father);
  return S.cur[k] ? null : (L.byName[k] || null);
}
function pcLeftNote_(rec) { return 'LEFT STUDENT-' + (pcClean_(rec.status).toUpperCase() || 'LEFT'); }
/** Portal backup DOB. By Student Name + Father Name when the sheet has those columns, else by the linked student's Enroll# */
function pcPortal_() {
  var sh = SpreadsheetApp.openById(PC_CFG.USERS).getSheetByName(PC_CFG.PORTAL), out = { byName: {}, byId: {}, named: false };
  if (!sh) return out;
  var v = sh.getDataRange().getDisplayValues(), H = (v[0] || []).map(pcHead_);
  var e = pcCol_(H, ['Enrollment#', 'Enrollment', 'Enroll#']), d = pcCol_(H, ['DOB', 'Date of Birth']),
    n = pcCol_(H, ['Student Name', 'Name']), f = pcCol_(H, ['Father Name', 'Father']);
  if (d < 0) return out;
  out.named = n >= 0 && f >= 0;
  var dup = {};
  for (var i = 1; i < v.length; i++) {
    var x = pcClean_(v[i][d]); if (!x || pcPlaceholderDob_(x)) continue;
    if (out.named) { var k = pcKey_(v[i][n], v[i][f]); if (k !== '|') { if (out.byName[k] && out.byName[k] !== x) dup[k] = true; out.byName[k] = x; } }
    if (e >= 0) { var id = pcHead_(v[i][e]); if (id && !out.byId[id]) out.byId[id] = x; }
  }
  Object.keys(dup).forEach(function (k) { delete out.byName[k]; });   // two different DOBs for the same name: not used
  return out;
}
function pcDobFor_(P, name, father, enroll) { return (P.named && P.byName[pcKey_(name, father)]) || (enroll && P.byId[pcHead_(enroll)]) || ''; }

/* ============================================================ list of sheets */
function pcSheets_() {
  var book = SpreadsheetApp.openById(PC_CFG.BOOK);
  return { ok: true, url: book.getUrl(), sheets: book.getSheets().map(function (s) { return s.getName(); })
    .filter(function (n) { return /^HSSC\s/i.test(n) && !/Top\s*20/i.test(n); }) };
}

/** copy one Transform student into row i */
function pcFill_(X, i, t, how, P) {
  ['Enroll#', 'SName', 'SFname', 'Program', 'Class', 'Section', 'PCS'].forEach(function (f, k) {
    pcSetF_(X, i, f, [t.enroll, t.name, t.father, t.program, t.cls, t.section, t.pcs][k]);
  });
  if (!pcF_(X, i, 'Group') && t.group) pcSetF_(X, i, 'Group', t.group);
  if (!pcF_(X, i, 'DOB')) { var d = pcDobFor_(P, X.V[i][X.C.name], X.V[i][X.C.father], t.enroll); if (d) pcSetF_(X, i, 'DOB', d); }
  pcSetF_(X, i, 'Matched by', how);
}
function pcUnfill_(X, i, how) {
  ['Enroll#', 'SName', 'SFname', 'Program', 'Class', 'Section', 'PCS'].forEach(function (f) { pcSetF_(X, i, f, ''); });
  pcSetF_(X, i, 'Matched by', how);
}
var PC_HAND = /^Chosen by hand|^Auto 96|^Added in the app/i;      // links the admin confirmed: kept by Prepare

/* ============================================================ ONE GO: only 100% Name + Father Name matches */
function pcPrepare_(name, strength) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var X = pcCtx_(name), C = X.C, S = pcStrength_(strength), L = pcLeft_(), P = pcPortal_();
    if (!X.jcw.length) throw new Error('No Jinnah College rows in "' + name + '".');
    var total = pcTotal_(X.sh.getName()), year = pcYear_(X.sh.getName()), session = (year - 1) + '-' + String(year).slice(2), longSession = (year - 1) + '-' + year;
    var st = { jcw: X.jcw.length, left: 0, linked: 0, kept: 0, open: 0, dup: 0, dob: 0, ready: 0, fail: 0 }, used = {};
    X.jcw.forEach(function (i) { if (PC_HAND.test(pcF_(X, i, 'Matched by')) && pcF_(X, i, 'Enroll#')) used[pcHead_(pcF_(X, i, 'Enroll#'))] = i; });

    X.jcw.forEach(function (i) {
      var nm = pcClean_(X.V[i][C.name]), fa = pcClean_(X.V[i][C.father]), how0 = pcF_(X, i, 'Matched by');
      var wasAuto = /^LEFT STUDENT/i.test(how0);
      if (/^Left \/ no match/i.test(how0)) return;                                       // struck by the admin: kept
      var lf = pcLeftOf_(L, S, nm, fa);
      if (lf) { pcSetF_(X, i, 'Matched by', pcLeftNote_(lf)); if (!X.struck[i]) X.strikeOn.push(i); st.left++; return; }
      if (wasAuto) { pcSetF_(X, i, 'Matched by', ''); how0 = ''; if (X.struck[i]) X.strikeOff.push(i); }

      if (PC_HAND.test(how0) && pcF_(X, i, 'Enroll#')) { st.kept++; }
      else {
        var list = (S.byName[pcKey_(nm, fa)] || []).filter(function (o) { var u = used[pcHead_(o.id)]; return u === undefined || u === i; });
        if (list.length === 1) { used[pcHead_(list[0].id)] = i; pcFill_(X, i, list[0], 'Name + Father Name - Transform', P); st.linked++; }
        else { pcUnfill_(X, i, list.length > 1 ? 'Same name twice - choose in the app' : 'Not found in Transform'); st.open++; if (list.length > 1) st.dup++; }
      }
      if (!pcF_(X, i, 'DOB')) { var d = pcDobFor_(P, nm, fa, pcF_(X, i, 'Enroll#')); if (d) pcSetF_(X, i, 'DOB', d); }
      if (pcF_(X, i, 'DOB')) st.dob++;

      if (!pcF_(X, i, 'Grade')) {                                                         // the gazette's own Grade is kept
        var rem = C.remarks >= 0 ? pcClean_(X.V[i][C.remarks]).toUpperCase() : '';
        pcSetF_(X, i, 'Grade', PC_CFG.PASS.indexOf(rem) >= 0 ? rem
          : /FAIL|ABSENT|\bABS\b|UFM|CANCEL|COMP|SUPP|R\.?\s?L|WITHHELD|\bF\b/.test(rem) ? 'F'
          : pcGradeOf_(X.V[i][C.marks], total));
      }
      if (pcF_(X, i, 'Session') === longSession || !pcF_(X, i, 'Session')) pcSetF_(X, i, 'Session', session);
      if (/^\d+(st|nd|rd|th) Position in BISE Peshawar$/i.test(pcF_(X, i, 'Distinction_won'))) pcSetF_(X, i, 'Distinction_won', '');
      if (PC_CFG.PASS.indexOf(pcF_(X, i, 'Grade').toUpperCase()) >= 0 && C.roll >= 0 && pcClean_(X.V[i][C.roll])) st.ready++; else st.fail++;
    });
    st.cells = Object.keys(X.changed).length;
    pcFlush_(X);
    var out = pcPayload_(X, S, L);
    st.sheets = S.sheets; st.leftRead = L.count; st.leftSheets = L.sheets; st.fast = typeof Sheets !== 'undefined';
    out.stats = st;
    return out;
  } finally { lock.releaseLock(); }
}

/* ============================================================ data for the app */
function pcPayload_(X, S, L) {
  return { ok: true, sheet: X.sh.getName(), url: SpreadsheetApp.openById(PC_CFG.BOOK).getUrl() + '#gid=' + X.sh.getSheetId(),
    total: pcTotal_(X.sh.getName()), rows: X.jcw.map(function (i) { return pcRowOf_(X, i); }),
    source: S.rows.map(function (o) { return [o.id, o.name, o.father, o.cls, o.pcs, S.cur[pcKey_(o.name, o.father)] ? 1 : 0]; }),
    left: L.list.map(function (r) { return [r.name, r.father, r.enroll, r.status]; }) };
}
function pcLoad_(name, S, strength) {
  var X = pcCtx_(name);
  return pcPayload_(X, S || pcStrength_(strength), pcLeft_());
}

/* ============================================================ the ticked list: link many rows in ONE write */
// items: [{row, roll, name, father, id}]  (id = the Transform student the admin picked; '' = unlink)
function pcLink_(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var X = pcCtx_(req.sheet), S = pcStrength_(req.strength), P = pcPortal_(), used = {}, done = [], skipped = [];
    X.jcw.forEach(function (i) { var e = pcF_(X, i, 'Enroll#'); if (e) used[pcHead_(e)] = i; });
    (req.items || []).forEach(function (it) {
      if (it.id === undefined && it.enroll !== undefined) it.id = it.enroll;          // older app page sends "enroll"
      var i = pcFind_(X, it.row, it.roll, it.name, it.father);
      if (i < 0) { skipped.push({ row: it.row, msg: (it.name || it.roll) + ' not found in the sheet' }); return; }
      if (!it.id) { pcUnfill_(X, i, 'Not found in Transform'); done.push(i); return; }
      var t = S.byId[pcHead_(it.id)];
      if (!t) { skipped.push({ row: it.row, msg: it.id + ' is not in the student list' }); return; }
      var u = t.enroll ? used[pcHead_(t.enroll)] : undefined;
      if (u !== undefined && u !== i) { skipped.push({ row: it.row, msg: t.name + ' (' + t.enroll + ') is already linked to row ' + (u + 2) }); return; }
      var old = pcF_(X, i, 'Enroll#'); if (old && used[pcHead_(old)] === i) delete used[pcHead_(old)];
      pcFill_(X, i, t, it.by || 'Chosen by hand (app)', P);
      if (t.enroll) used[pcHead_(t.enroll)] = i;
      done.push(i);
    });
    pcFlush_(X);
    return { ok: true, done: done.map(function (i) { return i + 2; }), skipped: skipped, rows: done.map(function (i) { return pcRowOf_(X, i); }) };
  } finally { lock.releaseLock(); }
}

/* ============================================================ edit one student (or add one not in the gazette) */
var PC_EDIT = { name: 'name', father: 'father', roll: 'roll', reg: 'reg', marks: 'marks',
  grade: 'Grade', dob: 'DOB', group: 'Group', session: 'Session', dist: 'Distinction_won', total: 'Total Marks', conduct: 'Conduct' };
function pcSave_(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var X = pcCtx_(req.sheet), C = X.C, v = req.values || {}, Xp = req.expect || {}, i;
    if (!('id' in v) && 'enroll' in v) v.id = v.enroll;                                 // older app page sends "enroll"
    if (!pcClean_(v.name) && !Xp.name) return { ok: false, error: 'Name and Father name are needed.' };
    if (Number(req.row)) {
      i = pcFind_(X, req.row, Xp.roll, Xp.name, Xp.father);
      if (i < 0) return { ok: false, error: (Xp.name || Xp.roll) + ' was not found in "' + X.sh.getName() + '" any more. Press Reload.' };
    } else {
      i = X.n; X.n++; X.V.push(new Array(C.last).fill('')); X.struck.push(false); X.jcw.push(i);
      if (C.college >= 0) pcSet_(X, i, C.college, 'JINNAH COLLEGE FOR WOMEN UNIVERSITY OF PESHAWAR');
      pcSetF_(X, i, 'Matched by', 'Added in the app');
    }
    var roll = pcClean_(v.roll);
    if (roll && C.roll >= 0) for (var j = 0; j < X.n; j++) if (j !== i && pcClean_(X.V[j][C.roll]) === roll) return { ok: false, error: 'Roll No ' + roll + ' is already in row ' + (j + 2) + '.' };
    Object.keys(PC_EDIT).forEach(function (k) {
      if (!(k in v)) return;
      var val = pcClean_(v[k]), t = PC_EDIT[k];
      if (k === 'conduct' && val.toUpperCase() === PC_CFG.DEF.conduct) val = '';
      if (k === 'dist' && val.toUpperCase() === 'NIL') val = '';
      var col = typeof C[t] === 'number' ? C[t] : (C.f[t] >= 0 ? C.f[t] : (val ? C.addExtra(t) : -1));
      pcSet_(X, i, col, val);
    });
    if ('id' in v) {
      if (!v.id) pcUnfill_(X, i, 'Not found in Transform');
      else {
        var S = pcStrength_(req.strength), t = S.byId[pcHead_(v.id)];
        if (!t) return { ok: false, error: v.id + ' is not in the student list.' };
        var e = pcHead_(t.enroll);
        for (var k = 0; k < X.jcw.length; k++) { var o = X.jcw[k]; if (o !== i && e && pcHead_(pcF_(X, o, 'Enroll#')) === e) return { ok: false, error: t.name + ' (' + t.enroll + ') is already linked to row ' + (o + 2) + '.' }; }
        pcFill_(X, i, t, 'Chosen by hand (app)', pcPortal_());
      }
    }
    pcFlush_(X);
    return { ok: true, row: pcRowOf_(X, i), moved: Number(req.row) && Number(req.row) !== i + 2 };
  } finally { lock.releaseLock(); }
}

/** strike / un-strike rows (no certificate). items: [{row, roll, name, father}] */
function pcStrikes_(req) {
  var X = pcCtx_(req.sheet), done = [];
  (req.items || [req]).forEach(function (it) {
    var i = pcFind_(X, it.row, it.roll, it.name, it.father); if (i < 0) return;
    var on = req.on !== false && it.on !== false, cur = pcF_(X, i, 'Matched by');
    if (on) { if (!X.struck[i]) X.strikeOn.push(i); pcSetF_(X, i, 'Matched by', 'Left / no match (struck)'); }
    else { if (X.struck[i]) X.strikeOff.push(i); if (/^Left \/ no match|^LEFT STUDENT/.test(cur)) pcSetF_(X, i, 'Matched by', pcF_(X, i, 'Enroll#') ? 'Chosen by hand (app)' : 'Not found in Transform'); }
    done.push(i);
  });
  pcFlush_(X);
  return { ok: true, done: done.map(function (i) { return i + 2; }), rows: done.map(function (i) { return pcRowOf_(X, i); }) };
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
    // p[2] = 0: the line starts straight after the label and runs to the right edge (Conduct and Moral Character)
    return '<td style="white-space:nowrap;padding:0 6px 0 0;' + (p[2] ? '' : 'width:1%;') + '">' + pcEsc_(p[0]) + '</td><td style="' + (p[2] ? 'width:' + p[2] + '%;' : '') + 'border-bottom:1px solid #000;text-align:center;font-weight:bold;">' + (pcEsc_(p[1]) || '&nbsp;') + '</td>';
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
    pcRowHtml_([['Date of Birth', v.dob, 70]]) + pcRowHtml_([['Distinctions Won', v.dist, 70]]) + pcRowHtml_([['Conduct and Moral Character', v.conduct, 0]]) +
    '<div style="margin-top:40px;text-align:right;"><div style="display:inline-block;width:230px;text-align:center;"><div style="height:60px;">' +
    (sign ? '<img height="55" src="' + sign + '">' : '') + '</div>' + PC_CFG.SIGN.map(pcEsc_).join('<br>') + '</div></div></div>';
}
/** { sheet, rows:[row numbers] (empty = every ready student) } -> one PDF in Drive */
/** v7.5: the bytes of a certificates PDF made by pcPdf_ (to download or share it as a file on the phone) */
function pcPdfGet_(req) {
  var f = DriveApp.getFileById(String(req.id || ''));
  if (f.getMimeType() !== 'application/pdf') return { ok: false, error: 'Not a PDF.' };
  if (f.getSize() > 40 * 1024 * 1024) return { ok: false, error: 'The PDF is too big to send (' + Math.round(f.getSize() / 1048576) + ' MB). Open it in Drive instead.' };
  return { ok: true, name: f.getName(), b64: Utilities.base64Encode(f.getBlob().getBytes()) };
}
function pcPdf_(req) {
  var X = pcCtx_(req.sheet), d = { sheet: X.sh.getName(), total: pcTotal_(X.sh.getName()), rows: X.jcw.map(function (i) { return pcRowOf_(X, i); }) }, want = (req.rows || []).map(Number), A = pcAssets_();
  var list = d.rows.filter(function (r) {
    if (want.length) return want.indexOf(r.row) >= 0;
    return !r.struck && r.roll && PC_CFG.PASS.indexOf(String(r.grade).toUpperCase()) >= 0;
  }).map(function (r) {
    return { name: r.name, father: r.father, group: r.group || (pcPart_(d.sheet) === 2 ? 'H.S.S.C (Part-II)' : 'H.S.S.C (Part-I)'), session: r.session,
      roll: r.roll, reg: r.reg, marks: r.marks, total: r.total || d.total, grade: r.grade, dob: r.dob, dist: r.dist || PC_CFG.DEF.distinctions, conduct: r.conduct || PC_CFG.DEF.conduct };
  });
  if (!list.length) return { ok: false, error: 'No students to print.' };
  list.sort(function (a, b) { var x = Number(a.roll), y = Number(b.roll); return (isNaN(x) || isNaN(y)) ? String(a.roll).localeCompare(String(b.roll)) : x - y; });
  var mk = function (L) { return '<html><head><meta charset="utf-8"></head><body style="font-family:\'Times New Roman\',serif;font-size:15px;margin:0;">' +
    L.map(function (v, i) { return pcPage_(v, A.logo, A.sign, i === 0); }).join('') + '</body></html>'; };
  // v7.6: big lists (1,000+ certificates) are too much for one HTML -> PDF conversion: the page asks for parts of 40 and joins them
  if (req.chunk) {
    var from = Math.max(0, Number(req.from) || 0), n = Math.min(60, Number(req.n) || 40), part = list.slice(from, from + n);
    if (!part.length) return { ok: true, total: list.length, from: from, n: 0, b64: '' };
    var h = mk(part), pdf = Utilities.newBlob(typeof inkHtml_ === 'function' ? inkHtml_(h) : h, 'text/html', 'part.html').getAs('application/pdf');
    return { ok: true, total: list.length, from: from, n: part.length, b64: Utilities.base64Encode(pdf.getBytes()), sheet: d.sheet };
  }
  if (list.length > 60) return { ok: false, error: 'TOO_BIG', total: list.length };
  var html = mk(list);
  var tz = Session.getScriptTimeZone(), nm = list.length === 1 ? pcClean_(list[0].roll + ' - ' + list[0].name).replace(/[\\/:*?"<>|]/g, '') : 'Provisional Certificates ' + d.sheet + ' ' + Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HHmm');
  var it = DriveApp.getFoldersByName(PC_CFG.FOLDER), folder = it.hasNext() ? it.next() : DriveApp.createFolder(PC_CFG.FOLDER);
  var file = folder.createFile(Utilities.newBlob(typeof inkHtml_ === 'function' ? inkHtml_(html) : html, 'text/html', nm + '.html').getAs('application/pdf').setName(nm + '.pdf'));
  return { ok: true, url: file.getUrl(), id: file.getId(), n: list.length };
}
