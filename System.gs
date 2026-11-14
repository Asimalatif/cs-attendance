/* =====================================================================================================
   System.gs (v7.6) — backups of every workbook of the attendance software, restore, and getting the
   system ready for a new session. Page: system.html (College screen → Backups & new session). Main admin only.

   Add ONE line in Code.gs doPost, next to the other module lines (already there in Code.gs v7.6):
     if (/^sys_/.test(String(req.action || ''))) return json_(sysRoute_(user, req));

   Backups:  Drive folder "JCW System Backups" / "JCW session 2026-27 complete backup 14-11-2026 15-30"
             (a copy of every workbook + backup-manifest.json). Nothing is ever deleted by a backup.
   Restore:  needs your password + typing RESTORE and the backup's date. A safety backup is made first.
             Values and formulas of every sheet are put back into the SAME files (so every link keeps working).
   New session: backup (required, same day) → keep a copy of the strength sheet → remove 2nd Year →
             1st Year becomes 2nd Year → empty the attendance / test sheets. Each step asks first.
   ===================================================================================================== */
var SYS_CFG = {
  ROOT: 'JCW System Backups',
  MANIFEST: 'backup-manifest.json',
  // tabs of the Users workbook that hold one session's work (emptied for a new session if ticked)
  USERS_SESSION_SHEETS: ['Attendance Index', 'Guests', 'Substitutions', 'Student Remarks', 'Chief Proctor Reports', 'Deleted Attendance Log',
                         'Bunk Decisions', 'Bunkers Log', 'ID Card Requests', 'ID Cards Printed', 'Change History', 'Left Students', 'Admission Log'],
  USERS_DEFAULT_ON: ['Attendance Index', 'Guests', 'Substitutions', 'Student Remarks', 'Chief Proctor Reports', 'Deleted Attendance Log', 'Bunk Decisions', 'Bunkers Log', 'ID Card Requests']
};

function sysRoute_(user, req) {
  if (!isAdmin_(user)) return { ok: false, error: 'Only the main admin can make backups or prepare a new session.' };
  switch (String(req.action)) {
    case 'sys_info':          return sysInfo_();
    case 'sys_backup_start':  return sysBackupStart_(req);
    case 'sys_backup_copy':   return sysBackupCopy_(req);
    case 'sys_backup_finish': return sysBackupFinish_(req);
    case 'sys_restore_plan':  return sysGuard_(user, req, 'restore') || sysRestorePlan_(req);
    case 'sys_restore_sheet': return sysGuard_(user, req, 'restore') || sysRestoreSheet_(req);
    case 'sys_restore_tidy':  return sysGuard_(user, req, 'restore') || sysRestoreTidy_(req);
    case 'sys_ns_preview':    return sysNsPreview_(req);
    case 'sys_ns_snapshot':   return sysGuard_(user, req, 'session') || sysNsSnapshot_(req);
    case 'sys_ns_promote':    return sysGuard_(user, req, 'session') || sysNsPromote_(req);
    case 'sys_ns_sheets':     return sysNsSheets_(req);
    case 'sys_ns_empty':      return sysGuard_(user, req, 'session') || sysNsEmpty_(req);
  }
  return { ok: false, error: 'UNKNOWN_ACTION' };
}

/* ---------- safety ---------- */
function sysGuard_(user, req, kind) {
  if (String(req.pass || '') !== String(user.password)) return { ok: false, error: 'Your password is not correct.' };
  if (kind === 'restore') {
    if (String(req.phrase || '').trim().toUpperCase() !== 'RESTORE') return { ok: false, error: 'Type RESTORE to confirm.' };
    var lb = sysLastBackup_();
    if (!lb || Date.now() - lb.at > 3 * 3600 * 1000 || lb.folder === req.folder) return { ok: false, error: 'Make a safety backup first (it is made for you on the page): a restore needs a backup of today\'s state from the last 3 hours.' };
  }
  if (kind === 'session') {
    if (String(req.phrase || '').trim().toUpperCase() !== 'NEW SESSION') return { ok: false, error: 'Type NEW SESSION to confirm.' };
    var b = sysLastBackup_();
    if (!b || Date.now() - b.at > 24 * 3600 * 1000) return { ok: false, error: 'Make a complete backup first (today).' };
  }
  return null;
}
function sysLastBackup_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('SYS_LAST_BACKUP') || 'null'); } catch (e) { return null; } }
function sysSession_() { var y = academicStartYear_(); return y + '-' + String(y + 1).slice(2); }

/* ---------- every workbook the software uses ---------- */
function sysBooks_() {
  var out = [], seen = {};
  var add = function (id, role) {
    id = String(id || '').trim(); if (!id || seen[id]) return; seen[id] = 1;
    try { var f = DriveApp.getFileById(id); if (f.isTrashed()) return; out.push({ id: id, name: f.getName(), role: role, url: f.getUrl(), size: f.getSize() }); } catch (e) { out.push({ id: id, name: '(cannot open)', role: role, error: String(e.message || e) }); }
  };
  add(USERS_SS_ID, 'Users / Transform / Timetable / Portal Data');
  Object.keys(BOOK_IDS).forEach(function (t) { add(BOOK_IDS[t], 'Daily attendance ' + t); });
  try { SEASON.forEach(function (t) { var b = mtBook_(t); if (b) add(b.getId(), 'Monthly test ' + t); }); } catch (e) {}
  try { Object.keys(MA_BOOKS).forEach(function (t) { add(MA_BOOKS[t], 'Monthly attendance / leaves ' + t); }); } catch (e) {}
  try { add(GT_CFG.BOOK_ID, 'Grand test'); } catch (e) {}
  try { add(DT_CFG.BOOK_ID, 'Detention exam'); } catch (e) {}
  try { add(BR_CFG.BOOK, 'Board results / provisional certificates'); } catch (e) {}
  try { add(ADM_CFG.SETTINGS_BOOK_ID, 'Admission quota settings'); } catch (e) {}
  try { add(AT_CFG.SS_ID, 'Admin tools'); } catch (e) {}
  try { add(NTT_CFG.SPREADSHEET_ID, 'Timetable'); } catch (e) {}
  try { add(RS_GT_SS, 'Record summary'); } catch (e) {}
  // record workbooks kept in the Record summary folders (monthly tests, attendance records, detention)
  try {
    [[RS_MONTHLY_FOLDER, 'Record folder: monthly tests'], [RS_ATT_FOLDER, 'Record folder: attendance'], [RS_DET_FOLDER, 'Record folder: detention']].forEach(function (x) {
      var it = DriveApp.getFolderById(x[0]).getFilesByType(MimeType.GOOGLE_SHEETS);
      while (it.hasNext()) { var f = it.next(); add(f.getId(), x[1]); }
    });
  } catch (e) {}
  return out;
}
function sysRoot_() { var it = DriveApp.getFoldersByName(SYS_CFG.ROOT); return it.hasNext() ? it.next() : DriveApp.createFolder(SYS_CFG.ROOT); }
function sysManifest_(folder) {
  var it = folder.getFilesByName(SYS_CFG.MANIFEST);
  if (!it.hasNext()) return null;
  var f = it.next(); try { var m = JSON.parse(f.getBlob().getDataAsString()); m._file = f; return m; } catch (e) { return null; }
}
function sysSaveManifest_(folder, m) {
  var f = m._file; delete m._file;
  var txt = JSON.stringify(m, null, 1);
  if (f) f.setContent(txt); else f = folder.createFile(SYS_CFG.MANIFEST, txt, MimeType.PLAIN_TEXT);
  m._file = f;
}
function sysInfo_() {
  var root = sysRoot_(), list = [], it = root.getFolders();
  while (it.hasNext()) {
    var f = it.next(), m = sysManifest_(f);
    list.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), at: f.getDateCreated().getTime(), kind: m ? m.kind : '', complete: !!(m && m.complete),
                files: m ? m.files.length : 0, done: m ? m.files.filter(function (x) { return x.copy; }).length : 0 });
  }
  list.sort(function (a, b) { return b.at - a.at; });
  return { ok: true, session: sysSession_(), books: sysBooks_(), root: { id: root.getId(), url: root.getUrl() }, backups: list, last: sysLastBackup_() };
}

/* ---------- backup ---------- */
function sysBackupStart_(req) {
  var tz = Session.getScriptTimeZone(), kind = String(req.kind || 'Complete');
  var name = 'JCW session ' + sysSession_() + ' ' + (kind === 'Complete' ? 'complete backup' : kind.toLowerCase() + ' backup') + ' ' + Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HH-mm');
  var folder = sysRoot_().createFolder(name), books = sysBooks_().filter(function (b) { return !b.error; });
  var m = { name: name, kind: kind, session: sysSession_(), started: new Date().toISOString(), complete: false,
            files: books.map(function (b) { return { id: b.id, name: b.name, role: b.role, copy: '' }; }) };
  sysSaveManifest_(folder, m);
  return { ok: true, folder: folder.getId(), url: folder.getUrl(), name: name, files: m.files };
}
function sysBackupCopy_(req) {
  var folder = DriveApp.getFolderById(String(req.folder)), m = sysManifest_(folder);
  if (!m) return { ok: false, error: 'This backup folder has no manifest.' };
  var e = m.files.filter(function (x) { return x.id === req.id; })[0];
  if (!e) return { ok: false, error: 'This file is not in the backup list.' };
  if (!e.copy) {
    var c = DriveApp.getFileById(e.id).makeCopy(e.name + ' (backup)', folder);
    var lock = LockService.getScriptLock(); lock.waitLock(30000);
    try { m = sysManifest_(folder); e = m.files.filter(function (x) { return x.id === req.id; })[0]; e.copy = c.getId(); e.at = new Date().toISOString(); sysSaveManifest_(folder, m); }
    finally { lock.releaseLock(); }
  }
  return { ok: true, copy: e.copy };
}
function sysBackupFinish_(req) {
  var folder = DriveApp.getFolderById(String(req.folder)), m = sysManifest_(folder);
  var miss = m.files.filter(function (x) { return !x.copy; });
  if (miss.length) return { ok: false, error: miss.length + ' file(s) are not copied yet: ' + miss.map(function (x) { return x.name; }).join(', ') };
  m.complete = true; m.finished = new Date().toISOString(); sysSaveManifest_(folder, m);
  PropertiesService.getScriptProperties().setProperty('SYS_LAST_BACKUP', JSON.stringify({ at: Date.now(), folder: folder.getId(), name: m.name, kind: m.kind }));
  return { ok: true, name: m.name, url: folder.getUrl() };
}

/* ---------- restore (values + formulas, sheet by sheet, into the same files) ---------- */
function sysRestorePlan_(req) {
  var folder = DriveApp.getFolderById(String(req.folder)), m = sysManifest_(folder);
  if (!m || !m.complete) return { ok: false, error: 'This backup is not complete, so it cannot be restored.' };
  var want = req.only && req.only.length ? req.only : null;
  var files = m.files.filter(function (x) { return x.copy && (!want || want.indexOf(x.id) >= 0); }).map(function (x) {
    var ss = SpreadsheetApp.openById(x.copy);
    return { id: x.id, copy: x.copy, name: x.name, sheets: ss.getSheets().map(function (s) { return { name: s.getName(), rows: s.getLastRow(), cols: s.getLastColumn() }; }) };
  });
  return { ok: true, name: m.name, files: files };
}
function sysRestoreSheet_(req) {
  var src = SpreadsheetApp.openById(String(req.copy)).getSheetByName(String(req.sheet));
  if (!src) return { ok: false, error: 'Sheet ' + req.sheet + ' not in the backup.' };
  var book = SpreadsheetApp.openById(String(req.id)), dst = book.getSheetByName(String(req.sheet));
  var R = src.getLastRow(), C = src.getLastColumn(), start = Math.max(1, Number(req.start) || 1);
  if (start === 1) {
    if (!dst) dst = book.insertSheet(String(req.sheet));
    dst.clearContents();
    if (dst.getMaxRows() < Math.max(R, 1)) dst.insertRowsAfter(dst.getMaxRows(), Math.max(R, 1) - dst.getMaxRows());
    if (dst.getMaxColumns() < Math.max(C, 1)) dst.insertColumnsAfter(dst.getMaxColumns(), Math.max(C, 1) - dst.getMaxColumns());
  }
  if (!R || !C) return { ok: true, next: 0 };
  var n = Math.max(50, Math.min(R - start + 1, Math.floor(200000 / C)));
  if (start > R) return { ok: true, next: 0 };
  n = Math.min(n, R - start + 1);
  var rg = src.getRange(start, 1, n, C), V = rg.getValues(), F = rg.getFormulas();
  var out = V.map(function (r, i) { return r.map(function (v, j) { return F[i][j] ? F[i][j] : v; }); });
  dst.getRange(start, 1, n, C).setValues(out);
  return { ok: true, next: start + n <= R ? start + n : 0 };
}
/** sheets made after the backup are removed, so the file is as it was */
function sysRestoreTidy_(req) {
  var keep = {}; SpreadsheetApp.openById(String(req.copy)).getSheets().forEach(function (s) { keep[s.getName()] = 1; });
  var book = SpreadsheetApp.openById(String(req.id)), gone = [];
  book.getSheets().forEach(function (s) { if (!keep[s.getName()] && book.getSheets().length > 1) { gone.push(s.getName()); book.deleteSheet(s); } });
  return { ok: true, removed: gone };
}

/* ---------- new session ---------- */
function sysT_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(TRANSFORM_SHEET);
  if (!sh) throw new Error('Sheet ' + TRANSFORM_SHEET + ' not found.');
  var v = sh.getDataRange().getValues(), H = v[0].map(function (h) { return String(h).toLowerCase().replace(/[^a-z0-9]/g, ''); });
  var c = function (L) { for (var i = 0; i < L.length; i++) { var j = H.indexOf(L[i]); if (j >= 0) return j; } return -1; };
  return { sh: sh, v: v, cls: c(['class', 'year']), pcs: c(['pcs']), sec: c(['section']), roll: c(['enroll', 'enrollment', 'rollno']) };
}
function sysYr_(s) { s = String(s || ''); return /\b(1st|first)\b|^\s*(i|xi|part-?i)\s*$/i.test(s) ? 1 : /\b(2nd|second)\b|^\s*(ii|xii|part-?ii)\s*$/i.test(s) ? 2 : 0; }
/** 1st Year section -> 2nd Year section: A1 -> A2, FCS-1 D2 -> FCS-2 D2, Pre-Eng-1 D1 -> Pre-Eng-2 D1 */
function sysNextSec_(s) {
  s = String(s || '').trim(); if (!s) return s;
  if (/-\s*1(\b|\s)/.test(s)) return s.replace(/-\s*1(\b|\s)/, function (m) { return m.replace('1', '2'); });
  if (/^[A-Z]+1$/i.test(s)) return s.slice(0, -1) + '2';
  if (/1st/i.test(s)) return s.replace(/1st/i, '2nd');
  return s;
}
function sysNsPreview_(req) {
  var T = sysT_(), y1 = 0, y2 = 0, other = 0, map = {};
  T.v.slice(1).forEach(function (r) {
    if (!String(r[T.roll] || '').trim()) return;
    var y = sysYr_(r[T.cls]);
    if (y === 1) { y1++; var p = String(r[T.pcs] || r[T.sec] || ''); map[p] = sysNextSec_(p); }
    else if (y === 2) y2++; else other++;
  });
  var out = sysOutgoing_(req);
  var users = SpreadsheetApp.openById(USERS_SS_ID), us = SYS_CFG.USERS_SESSION_SHEETS.map(function (n) { var s = users.getSheetByName(n); return { name: n, rows: s ? Math.max(0, s.getLastRow() - 1) : -1, on: SYS_CFG.USERS_DEFAULT_ON.indexOf(n) >= 0 }; }).filter(function (x) { return x.rows >= 0; });
  var groups = [];
  Object.keys(BOOK_IDS).forEach(function (t) { groups.push({ key: 'ATT_' + t, label: 'Daily attendance ' + t + ' (section sheets are deleted; they are made again from the new timetable when teachers mark)' }); });
  try { SEASON.forEach(function (t) { if (mtBook_(t)) groups.push({ key: t, label: 'Monthly test ' + t }); }); } catch (e) {}
  try { Object.keys(MA_BOOKS).forEach(function (t) { if (MA_BOOKS[t]) groups.push({ key: 'MA_' + t, label: 'Monthly attendance ' + t }); }); } catch (e) {}
  groups.push({ key: 'GT', label: 'Grand test' }); groups.push({ key: 'DT', label: 'Detention exam' });
  return { ok: true, firstYear: y1, secondYear: y2, unknownYear: other, sections: map, snapshot: TRANSFORM_SHEET + ' ' + out,
           snapshotExists: !!SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(TRANSFORM_SHEET + ' ' + out), users: us, books: groups, last: sysLastBackup_() };
}
function sysOutgoing_(req) { var y = Number(req.outgoing) || (new Date().getFullYear() - 1); return String(y); }
/** step 1: a copy of the strength sheet as it is ("Transform-ALL-COLLEGE-STRENGTH 2026"): the 2nd Year leaving keep their record */
function sysNsSnapshot_(req) {
  var book = SpreadsheetApp.openById(USERS_SS_ID), name = TRANSFORM_SHEET + ' ' + sysOutgoing_(req);
  if (book.getSheetByName(name)) return { ok: true, name: name, existed: true };
  var c = book.getSheetByName(TRANSFORM_SHEET).copyTo(book).setName(name);
  var rg = c.getDataRange(); rg.setValues(rg.getValues());          // a frozen copy: no formulas that change later
  return { ok: true, name: name };
}
/** step 2: 2nd Year rows removed, 1st Year -> 2nd Year (Class, PCS, Section) */
function sysNsPromote_(req) {
  var book = SpreadsheetApp.openById(USERS_SS_ID);
  if (!book.getSheetByName(TRANSFORM_SHEET + ' ' + sysOutgoing_(req))) return { ok: false, error: 'Make the copy of the strength sheet first (step 1).' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var T = sysT_(), sh = T.sh, R = T.v.length, C = T.v[0].length;
    if (T.cls < 0) return { ok: false, error: 'No Class column in the strength sheet.' };
    var rg = sh.getRange(1, 1, R, C), F = rg.getFormulasR1C1(), V = T.v;
    if (F[0].some(function (f) { return f; }) || (F[1] || []).some(function (f) { return /QUERY|IMPORTRANGE|ARRAYFORMULA|FILTER\(/i.test(f); }))
      return { ok: false, error: 'The strength sheet is made by a formula (QUERY / IMPORTRANGE …): change it where the data comes from.' };
    var keep = [0], removed = 0, promoted = 0, secMap = req.sections || {};
    for (var i = 1; i < R; i++) { if (sysYr_(V[i][T.cls]) === 2) { removed++; continue; } keep.push(i); }
    var nv = keep.map(function (i) {
      var r = V[i].slice();
      if (i > 0 && sysYr_(r[T.cls]) === 1) {
        r[T.cls] = String(r[T.cls]).replace(/1st/i, '2nd').replace(/\bfirst\b/i, 'Second');
        if (T.pcs >= 0 && r[T.pcs]) r[T.pcs] = secMap[r[T.pcs]] || sysNextSec_(r[T.pcs]);
        if (T.sec >= 0 && r[T.sec]) r[T.sec] = secMap[r[T.sec]] || sysNextSec_(r[T.sec]);
        promoted++;
      }
      return r;
    });
    var nf = keep.map(function (i) { return F[i]; });
    sh.getRange(2, 1, Math.max(1, R - 1), C).clearContent();
    sh.getRange(1, 1, nv.length, C).setValues(nv);
    for (var j = 0; j < C; j++) {                                          // formula columns (R1C1, so they fit their new rows)
      if (!nf.some(function (r, k) { return k > 0 && r[j]; })) continue;
      for (var k = 1; k < nf.length; k++) if (nf[k][j]) sh.getRange(k + 1, j + 1).setFormulaR1C1(nf[k][j]);
    }
    try { cacheClear_('master_att3'); } catch (e) {}
    return { ok: true, removed: removed, promoted: promoted, left: nv.length - 1 };
  } finally { lock.releaseLock(); }
}
/** the sheets of one group to empty: { key } -> [names] */
function sysNsSheets_(req) {
  var key = String(req.key || '');
  if (key === 'USERS') return { ok: true, sheets: SYS_CFG.USERS_SESSION_SHEETS };
  if (/^ATT_/.test(key)) {
    var tok = key.slice(4), book = SpreadsheetApp.openById(BOOK_IDS[tok]);
    var idx = {}, ix = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(INDEX_SHEET);
    if (ix && ix.getLastRow() > 1) ix.getRange(2, 1, ix.getLastRow() - 1, 2).getValues().forEach(function (r) { if (String(r[1]).toUpperCase() === tok) idx[r[0]] = 1; });
    return { ok: true, sheets: book.getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return idx[n] || n === PD_SHEET; }) };
  }
  var b = bkBook_(key).book;
  return { ok: true, sheets: b.getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return !BK_SKIP.test(n); }) };
}
function sysNsEmpty_(req) {
  var key = String(req.key || ''), name = String(req.sheet || '');
  if (key === 'USERS') {
    if (SYS_CFG.USERS_SESSION_SHEETS.indexOf(name) < 0) return { ok: false, error: 'Not a session sheet.' };
    var us = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(name);
    if (us && us.getLastRow() > 1) us.getRange(2, 1, us.getLastRow() - 1, Math.max(1, us.getLastColumn())).clearContent();
    return { ok: true, did: 'emptied' };
  }
  if (/^ATT_/.test(key)) {
    var book = SpreadsheetApp.openById(BOOK_IDS[key.slice(4)]), sh = book.getSheetByName(name);
    if (!sh) return { ok: true, did: 'gone' };
    if (name === PD_SHEET) { if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(1, sh.getLastColumn())).clearContent(); return { ok: true, did: 'emptied' }; }
    if (book.getSheets().length <= 1) { sh.clear(); return { ok: true, did: 'cleared' }; }
    book.deleteSheet(sh);
    return { ok: true, did: 'deleted' };
  }
  // monthly test / monthly attendance / grand test / detention: student rows go, the first row stays (empty) as the pattern
  var B = bkBook_(key).book, s = B.getSheetByName(name);
  if (!s) return { ok: true, did: 'gone' };
  var L = bkSheet_(s);
  if (!L) return { ok: true, did: 'skipped (no Enroll# header)' };
  var first = L.hi + 2, last = s.getLastRow();
  if (last < first) return { ok: true, did: 'already empty' };
  if (last > first) s.deleteRows(first + 1, last - first);
  var rg = s.getRange(first, 1, 1, L.lc), fx = rg.getFormulas()[0];
  fx.forEach(function (f, c) { if (!f) s.getRange(first, c + 1).clearContent(); });
  s.getRange(first, 1, 1, L.lc).setFontLine('none');
  return { ok: true, did: 'emptied' };
}
