/**
 * ================================================================
 * HSSC FILL  -  complete the Jinnah College rows of a board result sheet made by the app
 *               ("HSSC 1st year 2026", "HSSC 2nd year 2026", ...) in the Provisional Certificates workbook
 * ================================================================
 * Same idea as "1st Year Fill", for any HSSC sheet. Only rows whose College Name is JINNAH COLLEGE FOR WOMEN
 * are touched (the other colleges' rows are never written).
 *
 * Columns added at the right (made by themselves):
 *   Enroll# | SName | SFname | Program | Class | Section | PCS | Group | Session | Distinction_won | DOB | Matched by
 *   Group / Session / Distinction_won come from the old 2nd year list ("Data" / "boardResult2NDYEAR") when the
 *   student is there, else Group = "Appeared in" of the strength sheet. They are what the certificate prints.
 *
 * MENU "HSSC Fill"
 *  0) Choose the sheet                       (asked once; shown in the menu messages)
 *  1) Fill automatically                     a) Roll No or Registration # -> Enroll# (boardResult1STYEAR for 1st year;
 *                                               boardResult2NDYEAR and Data for 2nd year)
 *                                            b) Name + Father Name, the same letters
 *                                            then the DOB (step 3)
 *  2) Match the rest (window)                same window as 1st Year Fill
 *  3) Fill DOB from Portal Data              DOB of the admission portal ("Portal Data" sheet of the Users workbook),
 *                                            exactly as the portal shows it (format not changed); 01/01/2000 and empty
 *                                            are left out; a student not in Portal Data gets the "Date of Birth" of the
 *                                            old 2nd year list, if she is there.
 *
 * Safe: rows that already have an Enroll# keep it; nothing outside the columns above is ever written.
 *
 * INSTALL (Provisional Certificates workbook > Extensions > Apps Script):
 *   + > Script > "HSSC Fill" > paste this file. Then add this line inside your existing onOpen():   hfMenu_();
 *   (do NOT add a second onOpen)
 * ================================================================
 */
var HF_CFG = {
  BOOK_ID: '1C2tdr2yXAja0_C5Nk4HIk2m-71Kc4I-FS5VQUhjUtHQ',            // Provisional Certificates
  COLLEGE: /JINNAH COLLEGE FOR WOMEN/i,
  LINKS_1: ['boardResult1STYEAR'],                                      // Roll No / Registration -> Enroll#
  LINKS_2: ['boardResult2NDYEAR', 'Data'],
  OLD_2: ['boardResult2NDYEAR', 'Data'],                                // 2nd year students who already left: also a source
  STRENGTH_ID: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',         // Users workbook
  STRENGTH_SHEET: 'Transform-ALL-COLLEGE-STRENGTH',
  LEFT_SHEET: 'Left Students',
  PORTAL_SHEET: 'Portal Data',                                          // admission portal data (DOB)
  FILL: ['Enroll#', 'SName', 'SFname', 'Program', 'Class', 'Section', 'PCS', 'Group', 'Session', 'Distinction_won'],
  DOB: 'DOB',
  HOW: 'Matched by',
  NO_MATCH_COLOR: '#FFF3C4',
  AUTO_MIN: 0.96
};

function hfMenu_() {
  SpreadsheetApp.getUi().createMenu('HSSC Fill')
    .addItem('0) Choose the sheet', 'hfChoose')
    .addItem('1) Fill automatically (Roll No, Registration, exact name) + DOB', 'hfFillAuto')
    .addItem('2) Match the rest (window)', 'hfOpen')
    .addItem('3) Fill DOB from Portal Data', 'hfFillDob')
    .addToUi();
}

/* ===================================================================== helpers */
function hfHead_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function hfLetters_(s) { return String(s == null ? '' : s).toUpperCase().replace(/[^A-Z]/g, ''); }
function hfTrim_(s) { return String(s == null ? '' : s).replace(/[​-‏⁠﻿]/g, '').replace(/\s+/g, ' ').trim(); }
function hfToast_(msg) { Logger.log(msg); try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'HSSC Fill', 20); } catch (e) {} }

function hfSheetName_() {
  var n = PropertiesService.getDocumentProperties().getProperty('HF_SHEET');
  if (n) return n;
  var hs = SpreadsheetApp.openById(HF_CFG.BOOK_ID).getSheets().map(function (s) { return s.getName(); }).filter(function (x) { return /^HSSC /i.test(x) && !/top 20/i.test(x); });
  if (hs.length === 1) return hs[0];
  throw new Error('Choose the sheet first: menu HSSC Fill > 0) Choose the sheet.');
}
function hfChoose() {
  var ui = SpreadsheetApp.getUi(), book = SpreadsheetApp.openById(HF_CFG.BOOK_ID);
  var hs = book.getSheets().map(function (s) { return s.getName(); }).filter(function (x) { return /^HSSC /i.test(x) && !/top 20/i.test(x); });
  var cur = PropertiesService.getDocumentProperties().getProperty('HF_SHEET') || '';
  var r = ui.prompt('HSSC Fill', 'Sheet to fill (type the name exactly):\n' + hs.join('\n') + (cur ? '\n\nNow: ' + cur : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var n = r.getResponseText().trim();
  if (!book.getSheetByName(n)) { ui.alert('There is no sheet "' + n + '".'); return; }
  PropertiesService.getDocumentProperties().setProperty('HF_SHEET', n);
  hfToast_('Sheet chosen: ' + n);
}
function hfPart_(name) { return /2nd|part.?ii\b|part.?2/i.test(name) ? 2 : 1; }
function hfSheet_() {
  var n = hfSheetName_(), sh = SpreadsheetApp.openById(HF_CFG.BOOK_ID).getSheetByName(n);
  if (!sh) throw new Error('Sheet "' + n + '" not found. Menu HSSC Fill > 0) Choose the sheet.');
  return sh;
}

/** students: strength sheet, Left Students and (2nd year sheets) the old 2nd year board list.
 *  { rows:[{ enroll, name, father, cls, pcs, left, vals:{FILL name: value}, dob }], byEnroll, byName } */
function hfStrength_(part) {
  var S = { rows: [], byEnroll: {}, byName: {} };
  var add = function (o) {
    var en = hfHead_(o.enroll); if (!en || S.byEnroll[en]) return;
    S.rows.push(o); S.byEnroll[en] = o;
    var k = hfLetters_(o.name) + '|' + hfLetters_(o.father); if (k !== '|') (S.byName[k] = S.byName[k] || []).push(o);
  };
  var book = SpreadsheetApp.openById(HF_CFG.STRENGTH_ID);
  [[HF_CFG.STRENGTH_SHEET, false], [HF_CFG.LEFT_SHEET, true]].forEach(function (src) {
    var sh = book.getSheetByName(src[0]);
    if (!sh) { if (!src[1]) throw new Error('Sheet "' + src[0] + '" not found in the Users workbook.'); return; }
    var tv = sh.getDataRange().getDisplayValues(); if (tv.length < 2) return;
    var th = tv[0].map(hfHead_), c = function (n) { return th.indexOf(hfHead_(n)); };
    var cE = c('Enroll#'), cN = c('Student Name'), cF = c('Father Name'), cP = c('Program'), cC = c('Class'), cS = c('Section'), cX = c('PCS'), cA = c('Appeared in');
    if (cE < 0 || cN < 0) throw new Error('Sheet "' + src[0] + '" needs Enroll# and Student Name.');
    tv.slice(1).forEach(function (r) {
      var g = function (i) { return i >= 0 ? hfTrim_(r[i]) : ''; };
      add({ enroll: g(cE), name: g(cN), father: g(cF), cls: g(cC), pcs: g(cX), left: src[1],
            vals: { 'Enroll#': g(cE), SName: g(cN), SFname: g(cF), Program: g(cP), Class: g(cC), Section: g(cS), PCS: g(cX), Group: g(cA), Session: '', Distinction_won: '' } });
    });
  });
  // the 2nd year board list of the college (girls who already left after the 2nd year): Class No = Enroll#
  var old = {};
  if (part === 2) {
    var pb = SpreadsheetApp.openById(HF_CFG.BOOK_ID);
    HF_CFG.OLD_2.forEach(function (n) {
      var sh = pb.getSheetByName(n); if (!sh || sh.getLastRow() < 2) return;
      var v = sh.getDataRange().getDisplayValues(), h = v[0].map(hfHead_), c = function (x) { return h.indexOf(hfHead_(x)); };
      var cE = c('Class No'), cN = c('Name'), cF = c('Father Name'), cG = c('Group'), cC = c('Class'), cS = c('Section'), cY = c('Session'), cD = c('Distinction_won'), cB = c('Date of Birth');
      if (cE < 0 || cN < 0) return;
      v.slice(1).forEach(function (r) {
        var g = function (i) { return i >= 0 ? hfTrim_(r[i]) : ''; }, en = hfHead_(r[cE]); if (!en) return;
        var o = old[en] = old[en] || { group: g(cG), session: g(cY), dist: g(cD), dob: g(cB) };
        if (!o.dob) o.dob = g(cB); if (!o.dist) o.dist = g(cD); if (!o.session) o.session = g(cY); if (!o.group) o.group = g(cG);
        if (!S.byEnroll[en]) add({ enroll: g(cE), name: g(cN), father: g(cF), cls: g(cC) || '2nd Year', pcs: '', left: true, old: true,
          vals: { 'Enroll#': g(cE), SName: g(cN), SFname: g(cF), Program: g(cG), Class: g(cC) || '2nd Year', Section: g(cS), PCS: '', Group: g(cG), Session: g(cY), Distinction_won: g(cD) } });
      });
    });
    // a student still in the strength sheet: Group / Session / Distinction_won of the 2nd year list win
    Object.keys(old).forEach(function (en) { var o = S.byEnroll[en]; if (!o) return; if (old[en].group) o.vals.Group = old[en].group; o.vals.Session = old[en].session; o.vals.Distinction_won = old[en].dist; });
  }
  S.old = old;
  return S;
}

/** 0-based columns of the HSSC sheet; missing FILL / DOB / HOW columns are added at the right */
function hfCols_(sh) {
  var lastCol = sh.getLastColumn(), hh = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(hfHead_);
  var colOf = function (title) {
    var c = hh.indexOf(hfHead_(title));
    if (c < 0) { lastCol++; sh.getRange(1, lastCol).setValue(title).setFontWeight('bold').setBackground('#E7B7AE'); c = lastCol - 1; hh[c] = hfHead_(title); }
    return c;
  };
  var C = { fill: HF_CFG.FILL.map(colOf), dob: colOf(HF_CFG.DOB), how: colOf(HF_CFG.HOW),
            roll: hh.indexOf('ROLLNO'), name: hh.indexOf('NAME'), father: hh.indexOf('FATHERNAME'), marks: hh.indexOf('MARKS'),
            reg: hh.indexOf('REGISTRATION') >= 0 ? hh.indexOf('REGISTRATION') : hh.indexOf('REGISTRATIONNO'), college: hh.indexOf('COLLEGENAME') };
  C.lastCol = lastCol;
  if (C.name < 0 || C.father < 0) throw new Error('"' + sh.getName() + '" needs Name and Father Name columns.');
  return C;
}
/** 0-based data rows of the sheet that belong to the college */
function hfMine_(rows, C) {
  var out = [];
  rows.forEach(function (r, i) { if (C.college < 0 || HF_CFG.COLLEGE.test(r[C.college])) out.push(i); });
  return out;
}
/** Roll No -> Enroll#, Registration -> Enroll# from the link sheets */
function hfLinks_(part) {
  var book = SpreadsheetApp.openById(HF_CFG.BOOK_ID), L = { roll: {}, reg: {} };
  (part === 2 ? HF_CFG.LINKS_2 : HF_CFG.LINKS_1).forEach(function (n) {
    var sh = book.getSheetByName(n); if (!sh || sh.getLastRow() < 2) return;
    var v = sh.getDataRange().getDisplayValues(), h = v[0].map(hfHead_);
    var e = h.indexOf('ENROLL'); if (e < 0) e = h.indexOf('CLASSNO');
    var r = h.indexOf('ROLLNO'), g = h.indexOf('REGISTRATION');
    if (e < 0) return;
    for (var j = 1; j < v.length; j++) {
      var en = hfTrim_(v[j][e]); if (!en) continue;
      if (r >= 0 && hfTrim_(v[j][r]) && !L.roll[hfTrim_(v[j][r])]) L.roll[hfTrim_(v[j][r])] = en;
      if (g >= 0 && hfTrim_(v[j][g]) && !L.reg[hfTrim_(v[j][g])]) L.reg[hfTrim_(v[j][g])] = en;
    }
  });
  return L;
}

/* ===================================================================== 1) FILL AUTOMATICALLY */
function hfFillAuto() {
  var sh = hfSheet_(), part = hfPart_(sh.getName()), S = hfStrength_(part), L = hfLinks_(part), C = hfCols_(sh), n = sh.getLastRow() - 1;
  if (n < 1) return;
  var rows = sh.getRange(2, 1, n, C.lastCol).getDisplayValues(), mine = hfMine_(rows, C);
  var vals = C.fill.map(function (c) { return sh.getRange(2, c + 1, n, 1).getValues(); });
  var bgs = C.fill.map(function (c) { return sh.getRange(2, c + 1, n, 1).getBackgrounds(); });
  var how = sh.getRange(2, C.how + 1, n, 1).getValues();
  var stat = { roll: 0, reg: 0, name: 0, kept: 0, left: 0, none: 0, two: 0 };

  mine.forEach(function (i) {
    var r = rows[i], cur = hfTrim_(r[C.fill[0]]);
    if (cur && cur !== '2 MATCHES') { stat.kept++; return; }
    if (!hfTrim_(r[C.name])) return;
    var fill = function (t, by) {
      C.fill.forEach(function (c, x) { vals[x][i][0] = t.vals[HF_CFG.FILL[x]] || ''; bgs[x][i][0] = null; });
      how[i][0] = by + (t.old ? ' · 2nd year list' : t.left ? ' · Left Students' : '');
    };
    var mark = function (enroll, by) {
      C.fill.forEach(function (c, x) { vals[x][i][0] = ''; bgs[x][i][0] = HF_CFG.NO_MATCH_COLOR; });
      vals[0][i][0] = enroll; how[i][0] = by;
    };
    var e1 = C.roll >= 0 ? L.roll[hfTrim_(r[C.roll])] || '' : '', e2 = C.reg >= 0 ? L.reg[hfTrim_(r[C.reg])] || '' : '';
    if (e1 && S.byEnroll[hfHead_(e1)]) { fill(S.byEnroll[hfHead_(e1)], 'Roll No'); stat.roll++; return; }
    if (e2 && S.byEnroll[hfHead_(e2)]) { fill(S.byEnroll[hfHead_(e2)], 'Registration #'); stat.reg++; return; }
    var hits = S.byName[hfLetters_(r[C.name]) + '|' + hfLetters_(r[C.father])] || [];
    if (hits.length > 1) {
      var stay = hits.filter(function (t) { return !t.left; }); if (stay.length) hits = stay;
      var second = hits.filter(function (t) { return /2nd/i.test(t.cls); }); if (second.length) hits = second;
    }
    if (hits.length === 1) { fill(hits[0], 'Name + Father Name'); stat.name++; return; }
    if (hits.length > 1) { mark('2 MATCHES', 'Same name twice: fill by hand'); stat.two++; }
    else if (e1 || e2) { mark(e1 || e2, 'Enroll# not in strength sheet / Left Students / 2nd year list'); stat.left++; }
    else { mark('', 'Not found (spelling?)'); stat.none++; }
  });

  C.fill.forEach(function (c, x) { sh.getRange(2, c + 1, n, 1).setValues(vals[x]).setBackgrounds(bgs[x]); });
  sh.getRange(2, C.how + 1, n, 1).setValues(how);
  SpreadsheetApp.flush();
  var d = hfFillDob_(sh, S);
  hfToast_(sh.getName() + ' — Jinnah College rows: ' + mine.length + ' | by Roll No: ' + stat.roll + ' | by Registration: ' + stat.reg + ' | by name: ' + stat.name +
           ' | already filled: ' + stat.kept + ' | Enroll# only: ' + stat.left + ' | not found: ' + stat.none + ' | same name twice: ' + stat.two +
           (stat.none + stat.two ? '  -> menu 2) Match the rest' : '') + '  ||  ' + d);
}

/* ===================================================================== 3) DOB FROM PORTAL DATA */
function hfFillDob() { var sh = hfSheet_(); hfToast_(hfFillDob_(sh, hfStrength_(hfPart_(sh.getName())))); }
function hfFillDob_(sh, S) {
  var isPlaceholder = function (s) {
    s = String(s).trim(); var m;
    if ((m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/))) return +m[1] === 1 && +m[2] === 1 && +m[3] === 2000;
    if ((m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/))) return +m[1] === 2000 && +m[2] === 1 && +m[3] === 1;
    return false;
  };
  var src = SpreadsheetApp.openById(HF_CFG.STRENGTH_ID).getSheetByName(HF_CFG.PORTAL_SHEET);
  if (!src) throw new Error('Sheet "' + HF_CFG.PORTAL_SHEET + '" not found in the Users workbook.');
  var pv = src.getDataRange().getDisplayValues(), ph = pv[0].map(hfHead_);
  var pe = ph.indexOf('ENROLLMENT'), pd = ph.indexOf('DOB');
  if (pe < 0 || pd < 0) throw new Error('"' + HF_CFG.PORTAL_SHEET + '" needs Enrollment# and DOB columns.');
  var portal = {};
  for (var i = 1; i < pv.length; i++) { var k = hfHead_(pv[i][pe]), d = String(pv[i][pd]).trim(); if (k && (portal[k] === undefined || (!portal[k] && d))) portal[k] = d; }
  var C = hfCols_(sh), n = sh.getLastRow() - 1; if (n < 1) return 'DOB: no rows';
  var rows = sh.getRange(2, 1, n, C.lastCol).getDisplayValues(), mine = hfMine_(rows, C);
  var out = sh.getRange(2, C.dob + 1, n, 1).getDisplayValues().map(function (x) { return [x[0]]; });
  var st = { portal: 0, list: 0, none: 0 };
  mine.forEach(function (i) {
    var k = hfHead_(rows[i][C.fill[0]]); if (!k || k === '2MATCHES') return;
    var d = portal[k];
    if (d && !isPlaceholder(d)) { out[i][0] = d; st.portal++; return; }
    var o = S && S.old && S.old[k];
    if (o && o.dob && !isPlaceholder(o.dob)) { out[i][0] = o.dob; st.list++; return; }
    out[i][0] = ''; st.none++;
  });
  var rg = sh.getRange(2, C.dob + 1, n, 1); rg.setNumberFormat('@'); rg.setValues(out);
  return 'DOB from Portal Data: ' + st.portal + ' | from the 2nd year list: ' + st.list + ' | no DOB: ' + st.none;
}

/* ===================================================================== 2) MATCH THE REST: window */
function hfOpen() {
  var sh = hfSheet_();
  var html = HtmlService.createHtmlOutput(hfHtml_(sh.getName())).setWidth(1040).setHeight(660);
  SpreadsheetApp.getUi().showModelessDialog(html, 'HSSC Fill - ' + sh.getName() + ' - match the rest');
}
function hfGetData() {
  var sh = hfSheet_(), C = hfCols_(sh), n = sh.getLastRow() - 1, S = hfStrength_(hfPart_(sh.getName())), targets = [], used = {};
  if (n > 0) {
    var v = sh.getRange(2, 1, n, C.lastCol).getDisplayValues(), lines = sh.getRange(2, C.name + 1, n, 1).getFontLines();
    hfMine_(v, C).forEach(function (i) {
      var r = v[i], name = hfTrim_(r[C.name]), father = hfTrim_(r[C.father]), enroll = hfTrim_(r[C.fill[0]]);
      if (!name && !father) return;
      var struck = lines[i][0] === 'line-through';
      if (enroll && enroll !== '2 MATCHES' && !struck) used[hfHead_(enroll)] = i + 2;
      targets.push({ row: i + 2, name: name, father: father, roll: C.roll >= 0 ? hfTrim_(r[C.roll]) : '', marks: C.marks >= 0 ? hfTrim_(r[C.marks]) : '',
                     enroll: enroll, how: hfTrim_(r[C.how]), struck: struck, open: !struck && (!enroll || enroll === '2 MATCHES' || !S.byEnroll[hfHead_(enroll)]) });
    });
  }
  var source = S.rows.map(function (r) { return { enroll: r.enroll, name: r.name, father: r.father, cls: r.cls, pcs: r.pcs, left: r.left, usedRow: used[hfHead_(r.enroll)] || 0 }; });
  return { targets: targets, source: source, autoMin: HF_CFG.AUTO_MIN };
}
function hfApply(row, expected, enroll, by) { return hfApplyMany([{ row: row, expected: expected, enroll: enroll, by: by || 'Chosen by hand (window)' }]); }
function hfApplyMany(items) {
  var sh = hfSheet_(), C = hfCols_(sh), n = sh.getLastRow() - 1, S = hfStrength_(hfPart_(sh.getName()));
  if (n < 1) return { ok: false, msg: 'No rows.' };
  var names = sh.getRange(2, C.name + 1, n, 1).getDisplayValues(), fathers = sh.getRange(2, C.father + 1, n, 1).getDisplayValues();
  var vals = C.fill.map(function (c) { return sh.getRange(2, c + 1, n, 1).getValues(); });
  var bgs = C.fill.map(function (c) { return sh.getRange(2, c + 1, n, 1).getBackgrounds(); });
  var how = sh.getRange(2, C.how + 1, n, 1).getValues(), usedBy = {}, done = [], skipped = [];
  vals[0].forEach(function (x, i) { var k = hfHead_(x[0]); if (k && k !== '2MATCHES') usedBy[k] = i + 2; });
  items.forEach(function (it) {
    var i = Number(it.row) - 2;
    if (!(i >= 0 && i < n)) { skipped.push({ row: it.row, msg: 'row not found' }); return; }
    if (hfTrim_(names[i][0]) !== hfTrim_(it.expected.name) || hfTrim_(fathers[i][0]) !== hfTrim_(it.expected.father)) { skipped.push({ row: it.row, msg: 'row changed since the window was opened' }); return; }
    var t = S.byEnroll[hfHead_(it.enroll)];
    if (!t) { skipped.push({ row: it.row, msg: it.enroll + ' is not in the student lists' }); return; }
    var other = usedBy[hfHead_(it.enroll)];
    if (other && other !== Number(it.row)) { skipped.push({ row: it.row, msg: it.enroll + ' is already used in row ' + other }); return; }
    C.fill.forEach(function (c, x) { vals[x][i][0] = t.vals[HF_CFG.FILL[x]] || ''; bgs[x][i][0] = null; });
    how[i][0] = (it.by || 'Chosen by hand (window)') + (t.old ? ' · 2nd year list' : t.left ? ' · Left Students' : '');
    usedBy[hfHead_(it.enroll)] = Number(it.row);
    done.push({ row: Number(it.row), vals: HF_CFG.FILL.map(function (f) { return t.vals[f] || ''; }) });
  });
  if (done.length) {
    C.fill.forEach(function (c, x) { sh.getRange(2, c + 1, n, 1).setValues(vals[x]).setBackgrounds(bgs[x]); });
    sh.getRange(2, C.how + 1, n, 1).setValues(how);
    SpreadsheetApp.flush();
    hfFillDob_(sh, S);
  }
  return { ok: true, done: done, skipped: skipped,
           msg: done.length ? (done.length === 1 ? 'Row ' + done[0].row + ' filled.' : done.length + ' rows filled.') : (skipped[0] ? 'Nothing saved: ' + skipped[0].msg + '.' : 'Nothing saved.') };
}
function hfStrike(row, expected, on) {
  row = Number(row);
  var sh = hfSheet_(), C = hfCols_(sh);
  if (!row || row < 2 || row > sh.getLastRow()) return { ok: false, msg: 'Row not found. Close and reopen the window.' };
  var rg = sh.getRange(row, 1, 1, C.lastCol), v = rg.getDisplayValues()[0];
  if (hfTrim_(v[C.name]) !== hfTrim_(expected.name) || hfTrim_(v[C.father]) !== hfTrim_(expected.father)) return { ok: false, msg: 'Row ' + row + ' has changed since the window was opened. Nothing saved.' };
  rg.setFontLine(on ? 'line-through' : 'none');
  var hc = sh.getRange(row, C.how + 1), cur = hfTrim_(hc.getDisplayValue());
  if (on) hc.setValue('Left / no match (struck)'); else if (cur === 'Left / no match (struck)') hc.setValue('');
  return { ok: true, msg: 'Row ' + row + (on ? ' struck through.' : ': strike-through removed.') };
}

/* ===================================================================== window page (same as 1st Year Fill) */
function hfHtml_(sheetName) {
  return `<!DOCTYPE html><html><head><base target="_top"><style>
*{box-sizing:border-box}
body{font-family:Arial,sans-serif;font-size:13px;margin:0;color:#222}
#wrap{display:flex;height:100vh}
#left{width:37%;border-right:1px solid #ccd;padding:10px;display:flex;flex-direction:column}
#right{width:63%;padding:10px;overflow:auto}
h3{margin:0 0 6px;color:#1F4E8C;font-size:14px}
input[type=text]{width:100%;padding:7px;font-size:13px;border:1px solid #9bb;border-radius:5px;margin-bottom:6px}
label.chk{display:flex;align-items:center;gap:6px;margin:2px 0 6px;font-size:12px}
#list{flex:1;overflow:auto}
.item{border:1px solid #ddd;border-radius:6px;padding:5px 8px;margin-bottom:4px;cursor:pointer}
.item:hover{background:#eef3ff}.item.sel{border-color:#1F4E8C;background:#e3ecff}
.item small,.cand small{color:#555;display:block}
.badge{float:right;font-size:11px;padding:1px 6px;border-radius:10px;color:#fff}
.b-ok{background:#2e7d32}.b-check{background:#e65100}.b-struck{background:#757575}
.item.struck b,.item.struck small{text-decoration:line-through;color:#888}
#cur{border:2px solid #1F4E8C;border-radius:8px;padding:8px;background:#f5f8ff;margin-bottom:8px}
#cur .nm{font-weight:bold;color:#1F4E8C;font-size:14px}
.cand{border:1px solid #ddd;border-radius:6px;padding:6px 9px;margin-bottom:5px;cursor:pointer}
.cand:hover{background:#eef3ff}.cand.sel{border:2px solid #1F4E8C;background:#e3ecff}
.pct{float:right;font-weight:bold;padding:1px 7px;border-radius:10px;color:#fff;font-size:12px}
.p-hi{background:#2e7d32}.p-mid{background:#f9a825;color:#222}.p-lo{background:#757575}
.tag{font-size:11px;color:#b71c1c}
#change{border:1px solid #1F4E8C;border-radius:8px;padding:8px;margin-top:8px;background:#fafcff}
.crow{margin:3px 0}.new{color:#1b5e20;font-weight:bold}
button{padding:7px 12px;margin:4px 6px 0 0;border:0;border-radius:5px;cursor:pointer;font-size:12px;background:#e8eefb;color:#1F4E8C}
button.main{background:#1F4E8C;color:#fff}button.red{background:#b71c1c;color:#fff}button.green{background:#2e7d32;color:#fff}
#status{margin-top:8px;padding:6px 8px;border-radius:5px;background:#f4f4f4;font-size:12px;min-height:16px}
#status.ok{background:#e6f4ea;color:#1b5e20}#status.bad{background:#fdecea;color:#b71c1c}
.muted{color:#666;font-size:12px}
#auto{border:2px solid #2e7d32;border-radius:8px;padding:8px;margin-bottom:8px;background:#f1f8f2}
#autolog{max-height:140px;overflow:auto;font-size:12px;margin-top:6px;border-top:1px solid #cdd;padding-top:4px}
</style></head><body>
<div id="wrap">
 <div id="left">
  <h3>JCW rows of ${sheetName}</h3>
  <input type="text" id="q" placeholder="Search name, father, roll no or enroll#">
  <label class="chk"><input type="checkbox" id="only" checked> Only rows still without a student (no Enroll#)</label>
  <div class="muted" id="count"></div>
  <div id="list"></div>
 </div>
 <div id="right">
  <div id="auto" style="display:none"></div>
  <div id="autolog" style="display:none"></div>
  <div id="cur" style="display:none"><div class="nm" id="cur_nm"></div><div id="cur_info"></div>
   <button class="red" id="strikeBtn">Strike through (no match)</button></div>
  <div id="candbox" style="display:none">
   <h3>Closest students (strength sheet, Left Students, 2nd year board list)</h3>
   <label class="chk"><input type="checkbox" id="y2" checked> Only 2nd Year students (and the 2nd year board list)</label>
   <label class="chk"><input type="checkbox" id="lft" checked> Also students in "Left Students"</label>
   <label class="chk"><input type="checkbox" id="hide" checked> Hide students whose Enroll# is already used in the sheet</label>
   <input type="text" id="mq" placeholder="Not here? Search the strength sheet by name, father, Enroll# or PCS">
   <div id="cands"></div>
  </div>
  <div id="change" style="display:none"></div>
  <div id="status">Loading data...</div>
 </div>
</div>
<script>
var T=[],S=[],selT=null,selC=null,CANDS=[],AUTO_MIN=0.96,AUTOPAIRS=[],USED={};
function $(id){return document.getElementById(id)}
function say(m,c){var s=$("status");s.textContent=m;s.className=c||""}
function fail(e){say("Error: "+(e&&e.message?e.message:e),"bad")}

/* ---------- similarity (same as the Name Fixer) ---------- */
function tokens(s){return String(s||"").toUpperCase().replace(/[^A-Z ]/g," ").split(" ").filter(function(x){return x})}
function lev(a,b){
  var m=a.length,n=b.length,p=[],c=[],i,j,t;
  if(!m)return n; if(!n)return m;
  for(j=0;j<=n;j++)p[j]=j;
  for(i=1;i<=m;i++){c[0]=i;for(j=1;j<=n;j++){c[j]=Math.min(p[j]+1,c[j-1]+1,p[j-1]+(a.charAt(i-1)===b.charAt(j-1)?0:1))}t=p;p=c;c=t}
  return p[n];
}
function ratio(a,b){var L=Math.max(a.length,b.length);return L?1-lev(a,b)/L:1}
function collapse(s){var o="",i;for(i=0;i<s.length;i++){if(i===0||s.charAt(i)!==s.charAt(i-1))o+=s.charAt(i)}return o}
function skel(s){var o=s.charAt(0),i;for(i=1;i<s.length;i++){if("AEIOUY".indexOf(s.charAt(i))===-1)o+=s.charAt(i)}return collapse(o)}
function tokSim(a,b){
  if(a===b)return 1;
  var r=ratio(collapse(a),collapse(b)),sa=skel(a),sb=skel(b);
  if(sa.length>1&&sb.length>1)r=Math.max(r,0.8*ratio(sa,sb));
  if((a.length===1||b.length===1)&&a.charAt(0)===b.charAt(0))r=Math.max(r,0.75);
  return r;
}
function dirSim(A,B){var tot=0,w=0;A.forEach(function(x){var best=0;B.forEach(function(y){var v=tokSim(x,y);if(v>best)best=v});tot+=best*x.length;w+=x.length});return tot/w}
function textSim(A,B){if(!A.length||!B.length)return null;var s=(dirSim(A,B)+dirSim(B,A))/2,f=ratio(collapse(A.join("")),collapse(B.join("")));return Math.max(s,f)}
function score(t,s){
  var ns=textSim(t.nt,s.nt),fs=textSim(t.ft,s.ft),sc;
  if(ns===null)return {sc:0,ns:0,fs:fs};
  sc=(fs===null)?ns*0.85:0.6*ns+0.4*fs;
  return {sc:sc,ns:ns,fs:fs};
}
function sortedJoin(a){return collapse(a.slice().sort().join(""))}

/* ---------- load ---------- */
function prep(x){x.nt=tokens(x.name);x.ft=tokens(x.father)}
function setStatus(t){t.status=t.struck?"struck":(t.open?"check":"ok")}
function rebuildUsed(){USED={};T.forEach(function(t){if(!t.struck&&t.enroll&&t.enroll!=="2 MATCHES")USED[t.enroll.toUpperCase()]=t})}
function taken(s,t){var o=USED[s.enroll.toUpperCase()];return !!(o&&o!==t)}
function init(d){
  T=d.targets;S=d.source;AUTO_MIN=d.autoMin||0.96;
  S.forEach(function(s){prep(s);s.sj=sortedJoin(s.nt)});
  T.forEach(function(t){prep(t);setStatus(t)});
  rebuildUsed();renderList();
  var need=T.filter(function(t){return t.status==="check"}).length;
  say(T.length+" rows, "+S.length+" students in the strength sheet. "+need+" rows still without a student. Pick one on the left.","");
  startAuto();
}
google.script.run.withSuccessHandler(init).withFailureHandler(fail).hfGetData();

/* ---------- left list ---------- */
function filtered(){
  var q=$("q").value.trim().toLowerCase(),words=q?q.split(" "):[],only=$("only").checked;
  return T.filter(function(t){
    if(only&&t.status!=="check")return false;
    var hay=(t.name+" "+t.father+" "+t.roll+" "+t.enroll).toLowerCase();
    for(var i=0;i<words.length;i++){if(hay.indexOf(words[i])===-1)return false}
    return true;
  });
}
function renderList(){
  var lst=filtered(),box=$("list");box.textContent="";
  $("count").textContent=lst.length+" shown"+(lst.length>150?" (first 150)":"");
  lst.slice(0,150).forEach(function(t){
    var d=document.createElement("div");d.className="item"+(t===selT?" sel":"")+(t.struck?" struck":"");
    var b=document.createElement("span");b.className="badge "+(t.status==="ok"?"b-ok":t.status==="struck"?"b-struck":"b-check");
    b.textContent=t.status==="ok"?"filled":t.status==="struck"?"struck":"check";
    var n=document.createElement("b");n.textContent=t.name||"(no name)";
    var sm=document.createElement("small");
    sm.textContent="D/o "+(t.father||"-")+(t.roll?" | Roll No "+t.roll:"")+(t.marks?" | Marks "+t.marks:"")+(t.enroll?" | "+t.enroll:"");
    d.appendChild(b);d.appendChild(n);d.appendChild(sm);
    d.onclick=function(){selectT(t)};box.appendChild(d);
  });
}
$("q").addEventListener("input",renderList);
$("only").addEventListener("change",renderList);

/* ---------- right side ---------- */
function okSource(s,t){
  if(s.left&&!$("lft").checked)return false;
  if($("y2").checked&&!/2nd/i.test(s.cls))return false;
  if($("hide").checked&&taken(s,t))return false;
  return true;
}
function selectT(t){
  selT=t;selC=null;
  $("strikeBtn").textContent=t.struck?"Remove strike-through":"Strike through (no match)";
  $("cur_nm").textContent=t.name;
  $("cur_info").textContent="D/o "+(t.father||"-")+(t.roll?" | Roll No "+t.roll:"")+(t.marks?" | Marks "+t.marks:"")+(t.how?" | "+t.how:"");
  $("cur").style.display="block";$("candbox").style.display="block";
  $("mq").value="";$("change").style.display="none";
  var r=[];
  S.forEach(function(s){if(!okSource(s,t))return;var x=score(t,s);if(x.sc>=0.4)r.push({s:s,sc:x.sc,ns:x.ns,fs:x.fs})});
  r.sort(function(a,b){return b.sc-a.sc});
  CANDS=r.slice(0,8);renderCands();renderList();
  say(CANDS.length?"Click the right student.":"No close student. Try the search box, untick 'Only 2nd Year', or strike this row through.","");
}
function renderCands(){
  var box=$("cands");box.textContent="";
  CANDS.forEach(function(c){
    var s=c.s,d=document.createElement("div");d.className="cand"+(c===selC?" sel":"");
    var p=document.createElement("span"),pc=Math.round(c.sc*100);
    p.className="pct "+(pc>=85?"p-hi":pc>=65?"p-mid":"p-lo");p.textContent=pc+"%";
    var n=document.createElement("b");n.textContent=s.name;
    var f=document.createElement("small");f.textContent="D/o "+(s.father||"-")+" | "+s.enroll+" | "+(s.pcs||"-")+" | "+(s.cls||"-");
    d.appendChild(p);d.appendChild(n);
    if(s.left){var lt=document.createElement("span");lt.className="tag";lt.textContent="  Left Students";n.appendChild(lt)}
    d.appendChild(f);
    if(c.fs!==null&&c.fs!==undefined){var m=document.createElement("small");m.textContent="Name "+Math.round(c.ns*100)+"% | Father "+Math.round(c.fs*100)+"%";d.appendChild(m)}
    if(taken(s,selT)){var tg=document.createElement("small");tg.className="tag";tg.textContent="This Enroll# is already used in row "+USED[s.enroll.toUpperCase()].row;d.appendChild(tg)}
    d.onclick=function(){selC=c;renderCands();renderChange()};
    box.appendChild(d);
  });
}
$("mq").addEventListener("input",function(){
  if(!selT)return;
  var q=$("mq").value.trim().toLowerCase();
  if(!q){selectT(selT);return}
  var words=q.split(" "),r=[];
  for(var i=0;i<S.length&&r.length<200;i++){
    var s=S[i],hay=(s.name+" "+s.father+" "+s.enroll+" "+s.pcs).toLowerCase(),ok=true;
    for(var j=0;j<words.length;j++){if(hay.indexOf(words[j])===-1){ok=false;break}}
    if(ok&&!($("hide").checked&&taken(s,selT))){var x=score(selT,s);r.push({s:s,sc:x.sc,ns:x.ns,fs:x.fs})}
  }
  r.sort(function(a,b){return b.sc-a.sc});
  CANDS=r.slice(0,20);selC=null;renderCands();$("change").style.display="none";
});
$("hide").addEventListener("change",function(){if(selT)selectT(selT)});
$("y2").addEventListener("change",function(){if(selT)selectT(selT)});
$("lft").addEventListener("change",function(){if(selT)selectT(selT)});

/* ---------- change panel ---------- */
function renderChange(){
  var box=$("change"),s=selC.s;box.textContent="";box.style.display="block";
  var h=document.createElement("b");h.textContent="Will be filled in row "+selT.row;box.appendChild(h);
  [["Enroll#",s.enroll],["SName",s.name],["SFname",s.father],["PCS",s.pcs],["Class",s.cls]].forEach(function(x){
    var row=document.createElement("div");row.className="crow";
    var l=document.createElement("span");l.textContent=x[0]+": ";var v=document.createElement("span");v.className="new";v.textContent=x[1]||"-";
    row.appendChild(l);row.appendChild(v);box.appendChild(row);
  });
  var note=document.createElement("div");note.className="muted";note.textContent="Program and Section are filled too, from the strength sheet.";box.appendChild(note);
  var bA=document.createElement("button");bA.className="main";bA.textContent="Apply to sheet";bA.onclick=apply;
  var bS=document.createElement("button");bS.textContent="Skip to next";bS.onclick=function(){goNext(selT)};
  box.appendChild(bA);box.appendChild(bS);
}
function nextAfter(t){var lst=filtered(),i=lst.indexOf(t);return i>=0?lst[i+1]:null}
function goNext(t){var nx=nextAfter(t);if(nx)selectT(nx);else say("No more rows in this list.","")}
function markFilled(t,s,how){t.enroll=s.enroll;t.open=false;t.how=how;setStatus(t)}
function apply(){
  if(!selT||!selC)return;
  say("Saving...");
  var t=selT,s=selC.s,nx=nextAfter(t);
  google.script.run.withSuccessHandler(function(r){
    if(!r.ok||!r.done.length){say(r.msg,"bad");return}
    markFilled(t,s,"Chosen by hand (window)");rebuildUsed();updateAuto();
    $("change").style.display="none";selC=null;
    if(nx){selectT(nx);say(r.msg+" Next row loaded.","ok")}
    else{selT=null;$("cur").style.display="none";$("candbox").style.display="none";renderList();say(r.msg+" No more rows in this list.","ok")}
  }).withFailureHandler(fail).hfApply(t.row,{name:t.name,father:t.father},s.enroll,"Chosen by hand (window)");
}

/* ---------- strike through (left the college / no student) ---------- */
function strike(){
  if(!selT)return;
  var t=selT,on=!t.struck,nx=on?nextAfter(t):null;
  function undo(){t.struck=!on;setStatus(t);rebuildUsed();updateAuto();renderList()}
  t.struck=on;setStatus(t);rebuildUsed();updateAuto();
  if(on&&nx)selectT(nx);else selectT(t);
  say("Saving row "+t.row+"...","");
  google.script.run.withSuccessHandler(function(r){if(!r.ok){undo();say(r.msg,"bad");return}say(r.msg,"ok")})
    .withFailureHandler(function(e){undo();fail(e)}).hfStrike(t.row,{name:t.name,father:t.father},on);
}
$("strikeBtn").addEventListener("click",strike);

/* ---------- auto (best match 96% or more, clearly better than the next) ---------- */
function pending(){return AUTOPAIRS.filter(function(p){return p.t.open&&!p.t.struck&&!taken(p.s,p.t)})}
function startAuto(){
  var box=$("auto"),i=0,pairs=[],open=T.filter(function(t){return t.open&&!t.struck&&t.nt.length});
  box.style.display="block";box.textContent="Looking for matches of "+Math.round(AUTO_MIN*100)+"% or more...";
  var pool=S.filter(function(s){return /2nd/i.test(s.cls)||s.left});
  function chunk(){
    var end=Math.min(i+20,open.length);
    for(;i<end;i++){
      var t=open[i],tj=sortedJoin(t.nt),best=null,second=0;
      for(var j=0;j<pool.length;j++){
        var s=pool[j];
        if(taken(s,t))continue;
        if(ratio(tj,s.sj)<0.7)continue;
        var x=score(t,s);
        if(!best||x.sc>best.sc){if(best)second=best.sc;best={s:s,sc:x.sc}}else if(x.sc>second)second=x.sc;
      }
      if(best&&best.sc>=AUTO_MIN&&best.sc-second>=0.02)pairs.push({t:t,s:best.s,sc:best.sc});
    }
    if(i<open.length){box.textContent="Looking for matches... "+i+" / "+open.length;setTimeout(chunk,0)}
    else{
      var claims={};pairs.forEach(function(p){claims[p.s.enroll]=(claims[p.s.enroll]||0)+1});
      AUTOPAIRS=pairs.filter(function(p){return claims[p.s.enroll]===1});        // a student wanted by two rows -> by hand
      updateAuto();
    }
  }
  chunk();
}
function updateAuto(){
  var box=$("auto"),list=pending();box.textContent="";box.style.display="block";
  var m=document.createElement("div");
  if(!list.length){m.textContent="No automatic matches pending ("+Math.round(AUTO_MIN*100)+"% or more). Choose the rest by hand.";box.appendChild(m);return}
  m.textContent=list.length+" rows have a student (2nd Year or Left Students) of "+Math.round(AUTO_MIN*100)+"% or more. Their Enroll#, SName, SFname, Program, Class, Section and PCS will be filled.";
  var b=document.createElement("button");b.className="green";b.textContent="Auto-apply to "+list.length+" rows";b.onclick=autoApply;
  box.appendChild(m);box.appendChild(b);
}
function autoApply(){
  var list=pending();if(!list.length)return;
  var items=list.map(function(p){return {row:p.t.row,expected:{name:p.t.name,father:p.t.father},enroll:p.s.enroll,by:"Auto 96%+ (window)"}});
  var byRow={};list.forEach(function(p){byRow[p.t.row]=p});
  $("auto").textContent="Applying to "+list.length+" rows... please wait.";
  google.script.run.withSuccessHandler(function(r){
    var log=$("autolog");log.textContent="";log.style.display="block";
    r.done.forEach(function(x){
      var p=byRow[x.row];markFilled(p.t,p.s,"Auto 96%+ (window)");
      var d=document.createElement("div");d.textContent="Row "+x.row+": "+p.t.name+" (D/o "+p.t.father+") \\u2192 "+p.s.enroll+" "+p.s.name+" (D/o "+p.s.father+") "+p.s.pcs+" | "+Math.round(p.sc*100)+"%";
      log.appendChild(d);
    });
    r.skipped.forEach(function(x){var d=document.createElement("div");d.className="tag";d.textContent="Row "+x.row+" skipped: "+x.msg;log.appendChild(d)});
    rebuildUsed();renderList();updateAuto();if(selT)selectT(selT);
    say(r.done.length+" rows filled automatically"+(r.skipped.length?", "+r.skipped.length+" skipped":"")+". Choose the rest by hand.","ok");
  }).withFailureHandler(function(e){fail(e);updateAuto()}).hfApplyMany(items);
}
</script></body></html>`;
}