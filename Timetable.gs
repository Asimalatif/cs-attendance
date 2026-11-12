/**
 * Timetable.gs — NAILA TIMETABLE WEB EDITOR inside the attendance app (page timetable.html).
 * The timetable editor backend unchanged, plus its "additions" (change log, backups, login lock) already applied.
 * The page calls it through doPost with the app login (no second password): the app user name is looked up in
 * the "Users" sheet and the tt_admin column decides the role (ADMIN / ARTS / SCI MORNING / SCI EVENING / HOD / viewer).
 *
 * INSTALL:
 *   a) + > Script > name it "Timetable" > paste this file.
 *      If the old "naila tt gs" file is ALSO in this project, delete it (same function names).
 *   b) doPost, next to the rsm_ line (once):
 *        if (/^ntt_/.test(String(req.action || ''))) return json_(nttRoute_(user, req));
 *   c) Deploy > Manage deployments > Edit > New version > Deploy.
 */
var NTT_API_FNS = ['nttGetBoot', 'nttLogout', 'nttSaveSettings', 'nttUpdateRow', 'nttUpdateMany', 'nttAddTeacher', 'nttDeleteTeacher',
  'nttAddRoom', 'nttDeleteRoom', 'nttAddSectionName', 'nttAddSection', 'nttDeleteSection', 'nttAddSlot', 'nttDeleteSlot', 'nttGetLog', 'nttBackup'];
function nttRoute_(user, req) {
  try {
    if (req.action === 'ntt_login') {                      // sign in with the app login (no second password)
      var r = nttLogin(user.username, user.password);
      if (!r.ok) return { ok: false, error: 'Your app user name was not found in the timetable Users sheet.' };
      return { ok: true, data: r };
    }
    if (req.action === 'ntt_call') {
      var fn = String(req.fn || '');
      if (NTT_API_FNS.indexOf(fn) < 0) return { ok: false, error: 'Not allowed: ' + fn };
      var args = Array.isArray(req.args) ? req.args : [];
      var FN = { nttGetBoot: nttGetBoot, nttLogout: nttLogout, nttSaveSettings: nttSaveSettings, nttUpdateRow: nttUpdateRow,
        nttUpdateMany: nttUpdateMany, nttAddTeacher: nttAddTeacher, nttDeleteTeacher: nttDeleteTeacher, nttAddRoom: nttAddRoom,
        nttDeleteRoom: nttDeleteRoom, nttAddSectionName: nttAddSectionName, nttAddSection: nttAddSection, nttDeleteSection: nttDeleteSection,
        nttAddSlot: nttAddSlot, nttDeleteSlot: nttDeleteSlot, nttGetLog: nttGetLog, nttBackup: nttBackup };
      return { ok: true, data: FN[fn].apply(null, args) };
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/**
 * NAILA TIMETABLE WEB EDITOR - BACKEND  (file: naila tt gs.gs)
 * HTML file: "naila tt"  (naila tt.html)
 * Sheets: "Users", "Naila Timetable", "DropdownData", "NTT Settings" (auto-created)
 * All names are prefixed (ntt / NTT) so they cannot clash with other modules.
 * Keep ONLY ONE doGet in the whole project (if another file already has a doGet, delete the one below).
 *
 * ROLES  (Users sheet, column "tt_admin")
 *   ADMIN                 -> can change everything, including Settings
 *   ARTS TT ADMIN         -> can change only F.A programs
 *   SCI TT MORNING ADMIN  -> can change only MORNING science sections (every program except F.A)
 *   SCI TT EVENING ADMIN  -> can change only EVENING science sections
 *   HOD                   -> can change only his / her own department (teacher + days)
 *   (an old "SCI TT ADMIN" without the word EVENING is treated as MORNING)
 *   Morning / evening of a section is decided by Settings -> Evening sections.
 *
 * "NTT Settings" sheet stores (edited from the web interface)
 *   - EVENING sections  (clash checks never compare morning vs evening)
 *   - CREDIT hour rules per subject: morning min/max + evening min/max, and a default rule
 *   - TIME  : period timings of Morning / Evening (any length per period)
 *   - COUNT : number of periods defined for Morning / Evening (so removed periods stay removed)
 *   - MINP  : minimum number of period columns shown for Morning / Evening
 *   Only the main ADMIN can change timings / credit hours / rooms. ARTS and SCI EVENING admins can change the
 *   evening-section list for their own programs (needed when they add a section). SCI MORNING admin cannot.
 *
 * ROOMS: "R # 06/55" and "R # 06/55 (EVENING)" are the SAME room (the "(EVENING)" suffix is ignored when rooms are compared).
 *
 * SPEED: DropdownData is cached (CacheService, 6 h, cleared automatically when the data is changed from the
 *        interface; the Refresh button forces a reload). Row edits read only the edited rows of the timetable.
 */

const NTT_CFG = {
  SPREADSHEET_ID: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',   // the Users workbook (has Users, Naila Timetable, DropdownData)
  USERS: 'Users',
  TIMETABLE: 'Naila Timetable',
  DROPDOWN: 'DropdownData',
  SETTINGS: 'NTT Settings',      // created automatically on first "Save settings"
  ADMIN_MATCH: [],               // extra full-admin names / user names (empty: roles come only from the tt_admin column)
  HOD_RESTRICT_TO_DEPARTMENT: true,
  SESSION_SECONDS: 6 * 60 * 60,
  DD_CACHE_SECONDS: 6 * 60 * 60
};

// OPTIONAL: if auto-detection of the days columns fails, type the exact header names here,
// e.g.  const NTT_DAYS_COLS = ['Days'];   const NTT_PRACT_COLS = ['Practical Days'];
const NTT_DAYS_COLS = [];     // THEORY days column(s)
const NTT_PRACT_COLS = [];    // PRACTICAL days column(s)
const NTT_DAYS = ['M', 'TU', 'W', 'TH', 'F', 'S', 'SU'];
const NTT_DAY_ALIAS = {
  M: 0, MON: 0, MONDAY: 0, TU: 1, TUE: 1, TUES: 1, TUESDAY: 1, W: 2, WED: 2, WEDNESDAY: 2,
  TH: 3, THU: 3, THUR: 3, THURS: 3, THURSDAY: 3, F: 4, FRI: 4, FRIDAY: 4,
  S: 5, SAT: 5, SATURDAY: 5, SU: 6, SUN: 6, SUNDAY: 6
};
const NTT_SEC_COLS = ['PCS', 'Section', 'Class/Program/Section'];
const NTT_DD_SEC_COLS = ['PCS', 'Section', 'Sections', 'Class/Program/Section'];
const NTT_DD_ROOM_COLS = ['Room', 'Rooms'];
// header of the Science/Arts column in DropdownData (first match is used). Add your exact header here if different.
const NTT_DD_FACULTY_COLS = ['Arts', 'Science/Arts', 'Science / Arts', 'Arts/Science', 'Arts / Science', 'Faculty'];

const NTT_PD_MEMO = {};       // memo of parsed day strings (per execution)

// F.A programs = ARTS timetable. Every other program = SCIENCE timetable.
const NTT_ARTS_PROG_RE = /^\s*F\.?\s*A\b/i;



/* ---------------- helpers ---------------- */

function ntt_ss_() {
  return NTT_CFG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(NTT_CFG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}
function ntt_norm_(s) { return String(s == null ? '' : s).trim().toUpperCase().replace(/\s+/g, ' '); }
function ntt_key_(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }
// room identity: "R # 06/55 (EVENING)" and "R # 06/55" are the same room
function ntt_roomKey_(s) { return ntt_norm_(s).replace(/\s*\(\s*EVENING\s*\)\s*/g, ' ').replace(/\s+/g, ' ').trim(); }
function ntt_isEveRoomName_(s) { return /\(\s*EVENING\s*\)/i.test(String(s == null ? '' : s)); }

function ntt_headerMap_(headers) {
  const map = {};
  headers.forEach((h, i) => { if (h && map[ntt_key_(h)] === undefined) map[ntt_key_(h)] = i; });
  return map;
}
function ntt_table_(sheetName) {
  const sh = ntt_ss_().getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  return { sh: sh, headers: headers, map: ntt_headerMap_(headers), rows: data.slice(1) };
}
/**
 * FAST partial read: header row + only the block of sheet rows that are going to be edited.
 * t.rows is a sparse array indexed like the full table (index = sheetRow - 2); only the needed rows exist.
 * If the rows are far apart (more than 400 rows between the first and the last) the whole sheet is read.
 */
function ntt_tablePart_(sheetName, rowNums) {
  const sh = ntt_ss_().getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);
  const nums = (rowNums || []).map(Number).filter(n => n >= 2);
  const lo = nums.length ? Math.min.apply(null, nums) : 0, hi = nums.length ? Math.max.apply(null, nums) : 0;
  if (!nums.length || hi - lo > 400) return ntt_table_(sheetName);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const t = { sh: sh, headers: headers, map: ntt_headerMap_(headers), rows: [] };
  const block = sh.getRange(lo, 1, hi - lo + 1, lastCol).getValues();
  block.forEach((r, i) => { t.rows[lo - 2 + i] = r; });
  return t;
}
function ntt_get_(row, map, name) {
  const i = map[ntt_key_(name)];
  return i === undefined ? '' : row[i];
}
// index of first existing column among aliases, or -1
function ntt_col_(t, names) {
  for (let k = 0; k < names.length; k++) {
    const i = t.map[ntt_key_(names[k])];
    if (i !== undefined) return i;
  }
  return -1;
}
const NTT_TZ = Session.getScriptTimeZone();
function ntt_str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, NTT_TZ, 'dd-MMM-yyyy HH:mm');
  return v === null || v === undefined ? '' : String(v).trim();
}
function ntt_natSort_(a, b) { return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' }); }

// ---- batched sheet writing: cells are collected in memory, then written in a few range calls ----
function ntt_mark_(t, sheetRow, c, v) {
  if (c === undefined || c < 0) return;
  t._dirty = t._dirty || {};
  t._dirty[sheetRow + '|' + c] = { r: sheetRow, c: c, v: v };
  if (t.rows[sheetRow - 2]) t.rows[sheetRow - 2][c] = v;
}
function ntt_flush_(t) {
  const d = t._dirty; if (!d) return;
  const byCol = {};
  Object.keys(d).forEach(k => { const o = d[k]; (byCol[o.c] = byCol[o.c] || []).push(o); });
  Object.keys(byCol).forEach(c => {
    const list = byCol[c].sort((a, b) => a.r - b.r);
    let i = 0;
    while (i < list.length) {
      let j = i;
      while (j + 1 < list.length && list[j + 1].r === list[j].r + 1) j++;
      const vals = list.slice(i, j + 1).map(o => [o.v]);
      t.sh.getRange(list[i].r, Number(c) + 1, vals.length, 1).setValues(vals);
      i = j + 1;
    }
  });
  t._dirty = null;
}
// number or '' (blank / invalid)
function ntt_num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isNaN(n) || n < 0 ? '' : n;
}

/* ---------------- role scope (who may change what) ---------------- */

function ntt_isArtsProg_(p) { return NTT_ARTS_PROG_RE.test(String(p == null ? '' : p)); }

// evening sections from NTT Settings (read once per execution)
let NTT_EVE_MEMO = null;
function ntt_eveSet_() {
  if (NTT_EVE_MEMO) return NTT_EVE_MEMO;
  const set = {};
  ntt_settings_().evening.forEach(k => { set[String(k).split('|').map(ntt_norm_).join('|')] = true; });
  return (NTT_EVE_MEMO = set);
}
function ntt_isEveRow_(p, c, s) {
  const set = ntt_eveSet_(), pp = ntt_norm_(p), cc = ntt_norm_(c);
  return !!(set[pp + '|' + cc + '|' + ntt_norm_(s)] || set[pp + '|' + cc + '|*']);
}
// user.scope: ALL | ARTS | SCIM (science morning) | SCIE (science evening)
function ntt_scopeOk_(user, p, c, s) {
  if (user.role !== 'ADMIN') return false;
  const sc = user.scope || 'ALL';
  if (sc === 'ALL') return true;
  const arts = ntt_isArtsProg_(p);
  if (sc === 'ARTS') return arts;
  if (arts) return false;
  const eve = ntt_isEveRow_(p, c, s);
  return sc === 'SCIE' ? eve : !eve;
}
// program-level check (used where the shift of the target is not known yet, e.g. a brand-new section)
function ntt_progOk_(user, p) {
  if (user.role !== 'ADMIN') return false;
  const sc = user.scope || 'ALL';
  if (sc === 'ALL') return true;
  return sc === 'ARTS' ? ntt_isArtsProg_(p) : !ntt_isArtsProg_(p);
}
function ntt_needScope_(user, p, c, s) {
  if (!ntt_scopeOk_(user, p, c, s)) throw new Error('You are not allowed to change this part of the timetable.');
}
function ntt_needProg_(user, p) {
  if (!ntt_progOk_(user, p)) throw new Error('You are not allowed to change this timetable.');
}
function ntt_needFull_(user) {
  if (user.role !== 'ADMIN' || (user.scope || 'ALL') !== 'ALL') throw new Error('Only the main admin can do this.');
}
// [program, class, section] of a raw timetable row
function ntt_rowScope_(t, row) {
  const iS = ntt_col_(t, NTT_SEC_COLS);
  return [ntt_get_(row, t.map, 'Program'), ntt_get_(row, t.map, 'Class'), iS >= 0 ? row[iS] : ''];
}
// which evening-list keys a scoped admin may change in Settings (the morning science admin cannot change any)
function ntt_evKeyOk_(user, key) {
  const arts = ntt_isArtsProg_(String(key).split('|')[0]);
  return user.scope === 'ARTS' ? arts : user.scope === 'SCIE' ? !arts : false;
}

/* ---------------- cache (DropdownData) ---------------- */

function ntt_cachePut_(key, obj) {
  try {
    const s = JSON.stringify(obj), N = 30000, parts = Math.ceil(s.length / N), all = {};
    if (parts < 1 || parts > 90) return;                 // too big for the cache: just do not cache
    for (let i = 0; i < parts; i++) all[key + '_' + i] = s.substr(i * N, N);
    all[key + '_n'] = String(parts);
    CacheService.getScriptCache().putAll(all, NTT_CFG.DD_CACHE_SECONDS);
  } catch (e) {}
}
function ntt_cacheGet_(key) {
  try {
    const c = CacheService.getScriptCache(), n = Number(c.get(key + '_n'));
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const got = c.getAll(keys); let s = '';
    for (let i = 0; i < n; i++) { const p = got[key + '_' + i]; if (p == null) return null; s += p; }
    return JSON.parse(s);
  } catch (e) { return null; }
}
function ntt_cacheDrop_(key) { try { CacheService.getScriptCache().remove(key + '_n'); } catch (e) {} }
function ntt_ddDrop_() { ntt_cacheDrop_('nttdd'); }          // call after ANY change of the DropdownData sheet
function ntt_ddCached_(force) {
  if (!force) { const c = ntt_cacheGet_('nttdd'); if (c) return c; }
  const dd = ntt_dropdown_();
  ntt_cachePut_('nttdd', dd);
  return dd;
}

/* ---------------- days helpers ---------------- */

// "M-TH,F(P),S(P)" -> { th:[7 flags], pr:[7 flags] }   (memoised; the result must be treated as read-only)
function ntt_parseDays_(str) {
  const memoKey = String(str == null ? '' : str);
  if (NTT_PD_MEMO[memoKey]) return NTT_PD_MEMO[memoKey];
  const th = [0, 0, 0, 0, 0, 0, 0], pr = [0, 0, 0, 0, 0, 0, 0];
  memoKey.toUpperCase()
    .replace(/\s*\(\s*P\s*\)/g, '(P)')
    .replace(/\s*([-–—])\s*/g, '-')
    .split(/[,;\/&]+/).forEach(part => {
      part = part.trim(); if (!part) return;
      const isP = part.indexOf('(P)') >= 0;
      part = part.replace(/\(P\)/g, '').trim();
      part.split(/\s+/).forEach(tok => {
        if (!tok) return;
        const e = tok.split('-');
        const a = NTT_DAY_ALIAS[e[0]], b = e.length > 1 ? NTT_DAY_ALIAS[e[e.length - 1]] : a;
        if (a === undefined || b === undefined) return;
        const arr = isP ? pr : th;
        if (a <= b) { for (let d = a; d <= b; d++) arr[d] = 1; } else { arr[a] = 1; arr[b] = 1; }
      });
    });
  return (NTT_PD_MEMO[memoKey] = { th: th, pr: pr });
}
// theory flags -> "M-TH" / "M,TU,F,S" (runs of 3+ become ranges)
function ntt_fmtTheory_(th) {
  const out = []; let d = 0;
  while (d < 7) {
    if (!th[d]) { d++; continue; }
    let e = d; while (e + 1 < 7 && th[e + 1]) e++;
    if (e - d >= 2) out.push(NTT_DAYS[d] + '-' + NTT_DAYS[e]);
    else for (let k = d; k <= e; k++) out.push(NTT_DAYS[k]);
    d = e + 1;
  }
  return out.join(',');
}
function ntt_fmtPract_(pr) { return NTT_DAYS.filter((x, i) => pr[i]).join(','); }          // "F,S"
function ntt_fmtAll_(th, pr) {
  const a = ntt_fmtTheory_(th);
  const p = NTT_DAYS.filter((x, i) => pr[i]).map(x => x + '(P)').join(',');
  return a && p ? a + ',' + p : a + p;                                                      // "M-TH,F(P),S(P)"
}
// find the day-related columns of the timetable sheet
function ntt_dc_(t) {
  if (t._dc) return t._dc;
  const dc = { theory: [], pract: [], perDay: [] };
  const ov = list => list.map(n => t.map[ntt_key_(n)]).filter(i => i !== undefined);
  dc.theory = ov(NTT_DAYS_COLS); dc.pract = ov(NTT_PRACT_COLS);
  t.headers.forEach((h, i) => {
    const k = ntt_key_(h);
    if (!k || dc.theory.indexOf(i) >= 0 || dc.pract.indexOf(i) >= 0) return;
    if (/pract|\blab\b/.test(k)) { if (!NTT_PRACT_COLS.length) dc.pract.push(i); }
    else if (/\bdays?\b|theory/.test(k)) { if (!NTT_DAYS_COLS.length) dc.theory.push(i); }
    else if (NTT_DAY_ALIAS[k.toUpperCase()] !== undefined) dc.perDay.push({ i: i, d: NTT_DAY_ALIAS[k.toUpperCase()] });
  });
  t._dc = dc;
  return dc;
}
function ntt_rowDays_(t, row) {
  const dc = ntt_dc_(t), th = [0, 0, 0, 0, 0, 0, 0], pr = [0, 0, 0, 0, 0, 0, 0];
  dc.theory.forEach(i => {
    const p = ntt_parseDays_(ntt_str_(row[i]));
    for (let d = 0; d < 7; d++) { if (p.th[d]) th[d] = 1; if (p.pr[d]) pr[d] = 1; }
  });
  dc.pract.forEach(i => {
    const p = ntt_parseDays_(ntt_str_(row[i]));
    for (let d = 0; d < 7; d++) if (p.th[d] || p.pr[d]) pr[d] = 1;
  });
  dc.perDay.forEach(o => {
    const v = ntt_str_(row[o.i]);
    if (!v || /^(0|no|false|n|x|-)$/i.test(v)) return;
    if (/p/i.test(v)) pr[o.d] = 1; else th[o.d] = 1;
  });
  for (let d = 0; d < 7; d++) if (pr[d]) th[d] = 0;     // a practical day is not also a theory day
  return ntt_fmtAll_(th, pr);
}

/* ---------------- auth ---------------- */

function nttLogin(username, password) {
  const t = ntt_table_(NTT_CFG.USERS);
  const u = ntt_key_(username), p = String(password).trim();
  const failKey = 'nttfail_' + u, fc = CacheService.getScriptCache();
  if (!u) return { ok: false, error: 'Invalid user name or password.' };
  if (Number(fc.get(failKey) || 0) >= 5) return { ok: false, error: 'Too many wrong attempts. Try again in 5 minutes.' };
  for (let r = 0; r < t.rows.length; r++) {
    const row = t.rows[r];
    const un = ntt_key_(ntt_get_(row, t.map, 'User Name'));
    const pw = String(ntt_get_(row, t.map, 'Password')).trim();
    if (un && un === u && pw === p) {
      const name = String(ntt_get_(row, t.map, 'teacher')).trim();
      let roleRaw = t.map['tt_admin'] !== undefined ? row[t.map['tt_admin']] : row[8];
      roleRaw = ntt_norm_(roleRaw);
      const hasAdmin = roleRaw.indexOf('ADMIN') >= 0;
      let role, scope = '', label = '';
      if (hasAdmin && /\bARTS\b/.test(roleRaw)) { role = 'ADMIN'; scope = 'ARTS'; label = 'ARTS TIMETABLE ADMIN'; }
      else if (hasAdmin && /\bSCI(ENCE)?\b/.test(roleRaw)) {
        role = 'ADMIN';
        if (/\bEVE(NING)?\b/.test(roleRaw)) { scope = 'SCIE'; label = 'SCIENCE EVENING ADMIN'; }
        else { scope = 'SCIM'; label = 'SCIENCE MORNING ADMIN'; }      // "SCI TT MORNING ADMIN" (or an old "SCI TT ADMIN")
      }
      else if (hasAdmin || NTT_CFG.ADMIN_MATCH.some(m => ntt_key_(name).indexOf(m) >= 0 || un.indexOf(m) >= 0)) { role = 'ADMIN'; scope = 'ALL'; label = 'ADMIN'; }
      else { role = roleRaw.indexOf('HOD') >= 0 ? 'HOD' : 'VIEWER'; label = role; }
      const user = {
        name: name, username: un, role: role, scope: scope, label: label,
        department: String(ntt_get_(row, t.map, 'Department')).trim()
      };
      fc.remove(failKey);
      const token = Utilities.getUuid();
      CacheService.getScriptCache().put('ntt_' + token, JSON.stringify(user), NTT_CFG.SESSION_SECONDS);
      return { ok: true, token: token, user: user };
    }
  }
  fc.put(failKey, String(Number(fc.get(failKey) || 0) + 1), 300);
  return { ok: false, error: 'Invalid user name or password.' };
}

function ntt_session_(token) {
  const v = token && CacheService.getScriptCache().get('ntt_' + token);
  if (!v) throw new Error('SESSION_EXPIRED');
  return JSON.parse(v);
}

function nttLogout(token) { CacheService.getScriptCache().remove('ntt_' + token); }

/* ---------------- settings (evening sections + credit hours + timings) ---------------- */

/**
 * Sheet layout (columns): Type | Key | Min | Max | Eve Min | Eve Max
 *   META            saved
 *   EVENING         Program|Class|Section   (Section "*" = whole class)
 *   CREDIT          SUBJECT (upper case)    morning min/max   evening min/max
 *   CREDIT_DEFAULT  *                       morning min/max   evening min/max
 *   TIME            M|n|HH:MM-HH:MM  or  E|n|text   (period n timing of Morning / Evening)
 *   COUNT           M or E                  number of periods defined (value in Min column)
 *   MINP            M or E                  minimum number of period columns shown
 */

let NTT_SET_MEMO = null;
function ntt_settings_() { return NTT_SET_MEMO || (NTT_SET_MEMO = ntt_settingsRead_()); }
function ntt_settingsRead_() {
  const out = { evening: [], credits: {}, creditDefault: { min: '', max: '', emin: '', emax: '' }, times: { M: [], E: [] }, counts: { M: '', E: '' }, minP: { M: '', E: '' }, saved: false };
  const sh = ntt_ss_().getSheetByName(NTT_CFG.SETTINGS);
  if (!sh) return out;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const type = String(r[0] == null ? '' : r[0]).trim().toUpperCase();
    const key = String(r[1] == null ? '' : r[1]).trim();
    if (type === 'META') { out.saved = true; continue; }
    if (!key) continue;
    if (type === 'EVENING') out.evening.push(key);
    else if (type === 'CREDIT') out.credits[ntt_norm_(key)] = { min: ntt_num_(r[2]), max: ntt_num_(r[3]), emin: ntt_num_(r[4]), emax: ntt_num_(r[5]) };
    else if (type === 'TIME') {                       // key  M|3|09:40-10:20   (shift | period no | text)
      const m = key.match(/^([ME])\|(\d+)\|(.*)$/);
      if (m && Number(m[2]) >= 1 && Number(m[2]) <= 12) out.times[m[1]][Number(m[2]) - 1] = m[3];
    }
    else if (type === 'COUNT') { if (key === 'M' || key === 'E') out.counts[key] = ntt_num_(r[2]); }
    else if (type === 'MINP') { if (key === 'M' || key === 'E') out.minP[key] = ntt_num_(r[2]); }
    else if (type === 'CREDIT_DEFAULT') out.creditDefault = { min: ntt_num_(r[2]), max: ntt_num_(r[3]), emin: ntt_num_(r[4]), emax: ntt_num_(r[5]) };
  }
  // fill holes so the browser always receives plain strings
  ['M', 'E'].forEach(k => { for (let i = 0; i < out.times[k].length; i++) if (out.times[k][i] == null) out.times[k][i] = ''; });
  return out;
}

function nttSaveSettings(token, s) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN') throw new Error('Only admin can change settings.');
  s = s || {};

  // scoped admins: only the evening-section list of their own programs can change; everything else stays as saved
  if ((user.scope || 'ALL') !== 'ALL') {
    const cur = ntt_settings_();
    const keepOther = cur.evening.filter(k => !ntt_evKeyOk_(user, k));
    const mine = (s.evening || []).filter(k => ntt_evKeyOk_(user, k));
    s = { evening: keepOther.concat(mine), credits: cur.credits, creditDefault: cur.creditDefault,
          times: cur.times, counts: cur.counts, minP: cur.minP };
  }

  // sanitise
  const seen = {}, evening = [];
  (s.evening || []).forEach(k => {
    k = String(k == null ? '' : k).trim();
    if (k && !seen[k]) { seen[k] = true; evening.push(k); }
  });
  const credits = {};
  Object.keys(s.credits || {}).forEach(k => {
    const nk = ntt_norm_(k); if (!nk) return;
    const c0 = s.credits[k] || {};
    const mn = ntt_num_(c0.min), mx = ntt_num_(c0.max), en = ntt_num_(c0.emin), ex = ntt_num_(c0.emax);
    if (mn === '' && mx === '' && en === '' && ex === '') return;
    if (mn !== '' && mx !== '' && mn > mx) throw new Error('Credit hours for ' + nk + ' (morning): minimum is greater than maximum.');
    if (en !== '' && ex !== '' && en > ex) throw new Error('Credit hours for ' + nk + ' (evening): minimum is greater than maximum.');
    credits[nk] = { min: mn, max: mx, emin: en, emax: ex };
  });
  const d0 = s.creditDefault || {};
  const dmn = ntt_num_(d0.min), dmx = ntt_num_(d0.max), den = ntt_num_(d0.emin), dex = ntt_num_(d0.emax);
  if (dmn !== '' && dmx !== '' && dmn > dmx) throw new Error('Default credit hours (morning): minimum is greater than maximum.');
  if (den !== '' && dex !== '' && den > dex) throw new Error('Default credit hours (evening): minimum is greater than maximum.');

  // number of periods per shift (removed periods must stay removed)
  const counts = {};
  ['M', 'E'].forEach(k => {
    const v = ntt_num_((s.counts || {})[k]);
    counts[k] = (v === '' || v < 1 || v > 12) ? '' : Math.round(v);
  });
  const times = { M: [], E: [] };
  ['M', 'E'].forEach(k => {
    let arr = ((s.times || {})[k] || []).slice(0, 12).map(x => String(x == null ? '' : x).replace(/\|/g, '/').trim());
    if (counts[k] !== '') arr = arr.slice(0, counts[k]);          // periods beyond the count are dropped
    while (arr.length && !arr[arr.length - 1]) arr.pop();
    times[k] = arr;
  });
  const minP = {};
  ['M', 'E'].forEach(k => {
    const v = ntt_num_((s.minP || {})[k]);
    minP[k] = (v === '' || v < 1 || v > 12) ? '' : Math.round(v);
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ntt_ss_();
    let sh = ss.getSheetByName(NTT_CFG.SETTINGS);
    if (!sh) sh = ss.insertSheet(NTT_CFG.SETTINGS);
    const rows = [['Type', 'Key', 'Min', 'Max', 'Eve Min', 'Eve Max'], ['META', 'saved', '', '', '', '']];
    evening.forEach(k => rows.push(['EVENING', k, '', '', '', '']));
    Object.keys(credits).sort().forEach(k => rows.push(['CREDIT', k, credits[k].min, credits[k].max, credits[k].emin, credits[k].emax]));
    if (dmn !== '' || dmx !== '' || den !== '' || dex !== '') rows.push(['CREDIT_DEFAULT', '*', dmn, dmx, den, dex]);
    ['M', 'E'].forEach(k => times[k].forEach((txt, i) => rows.push(['TIME', k + '|' + (i + 1) + '|' + txt, '', '', '', ''])));
    ['M', 'E'].forEach(k => { if (counts[k] !== '') rows.push(['COUNT', k, counts[k], '', '', '']); });
    ['M', 'E'].forEach(k => { if (minP[k] !== '') rows.push(['MINP', k, minP[k], '', '', '']); });
    sh.clearContents();
    sh.getRange(1, 1, rows.length, 2).setNumberFormat('@');
    sh.getRange(1, 1, rows.length, 6).setValues(rows);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  NTT_EVE_MEMO = null; NTT_SET_MEMO = null;   // settings may have changed      
  ntt_log_(user, [['SETTINGS', 'Settings saved (evening sections, credit hours, timings)']]);
  return { ok: true, settings: ntt_settings_() };
}

/* ---------------- dropdown data ---------------- */

function ntt_dropdown_() {
  const t = ntt_table_(NTT_CFG.DROPDOWN);
  const iProg = ntt_col_(t, ['Program']);
  const iSec = ntt_col_(t, NTT_DD_SEC_COLS);
  const iRoom = ntt_col_(t, NTT_DD_ROOM_COLS);
  const iFac = ntt_col_(t, NTT_DD_FACULTY_COLS);

  const subjectInfo = {}, deptList = {}, facList = {}, desList = {}, statList = {};
  const teachers = {};
  const subjectTeachers = {};
  const subjectDepts = {};
  const subjects = {};
  const courses = [];
  const rooms = {};
  const sections = {};
  const byProgram = {};         // NORM program -> {subjects:{}, sections:{}}

  t.rows.forEach(row => {
    // rooms / sections lists (may live in their own columns, so read before the subject check)
    if (iRoom >= 0) { const v = ntt_str_(row[iRoom]); if (v) rooms[v] = true; }
    const secV = iSec >= 0 ? ntt_str_(row[iSec]) : '';
    if (secV) sections[secV] = true;

    const subject = String(ntt_get_(row, t.map, 'Subject')).trim();
    const teacher = String(ntt_get_(row, t.map, 'Teacher')).trim();
    const progK = iProg >= 0 ? ntt_norm_(row[iProg]) : '';
    if (progK) {
      byProgram[progK] = byProgram[progK] || { subjects: {}, sections: {} };
      if (subject) byProgram[progK].subjects[subject.toUpperCase()] = true;
      if (secV) byProgram[progK].sections[secV] = true;
    }

    if (!subject && !teacher) return;
    const sk = ntt_norm_(subject), tk = ntt_norm_(teacher);
    const dept = ntt_norm_(ntt_get_(row, t.map, 'Department'));
    const deptRaw = String(ntt_get_(row, t.map, 'Department')).trim();
    const facRaw = iFac >= 0 ? ntt_str_(row[iFac]) : '';
    const desRaw = String(ntt_get_(row, t.map, 'Designation')).trim();
    const statRaw = String(ntt_get_(row, t.map, 'Status')).trim();
    const contactRaw = String(ntt_get_(row, t.map, 'Contact')).trim();
    if (deptRaw) deptList[dept] = deptList[dept] || deptRaw;
    if (facRaw) facList[ntt_norm_(facRaw)] = facList[ntt_norm_(facRaw)] || facRaw;
    if (desRaw) desList[ntt_norm_(desRaw)] = desList[ntt_norm_(desRaw)] || desRaw;
    if (statRaw) statList[ntt_norm_(statRaw)] = statList[ntt_norm_(statRaw)] || statRaw;
    if (sk) {
      const si = subjectInfo[sk] = subjectInfo[sk] || { department: '', faculty: '' };
      if (!si.department && deptRaw) si.department = deptRaw;
      if (!si.faculty && facRaw) si.faculty = facRaw;
      subjects[sk] = subjects[sk] || subject.toUpperCase();
      subjectDepts[sk] = subjectDepts[sk] || {};
      if (dept) subjectDepts[sk][dept] = true;
      const course = String(ntt_get_(row, t.map, 'Course Name')).trim();
      if (course) courses.push({ subject: sk, cls: ntt_norm_(ntt_get_(row, t.map, 'Class')), course: course });
    }
    if (tk) {
      if (!teachers[tk]) teachers[tk] = {
        name: teacher.toUpperCase(),
        designation: String(ntt_get_(row, t.map, 'Designation')).trim(),
        status: String(ntt_get_(row, t.map, 'Status')).trim(),
        department: dept,
        contact: contactRaw
      };
      else if (!teachers[tk].contact && contactRaw) teachers[tk].contact = contactRaw;
      if (sk) {
        subjectTeachers[sk] = subjectTeachers[sk] || [];
        if (subjectTeachers[sk].indexOf(tk) < 0) subjectTeachers[sk].push(tk);
      }
    }
  });
  return {
    teachers: teachers, subjectTeachers: subjectTeachers, subjectDepts: subjectDepts,
    subjects: subjects, courses: courses, rooms: rooms, sections: sections, byProgram: byProgram,
    hasProg: iProg >= 0,
    subjectInfo: subjectInfo, deptList: deptList, facList: facList, desList: desList, statList: statList
  };
}

function ntt_usersContacts_() {
  const out = {};
  try {
    const t = ntt_table_(NTT_CFG.USERS);
    t.rows.forEach(row => {
      const n = ntt_norm_(ntt_get_(row, t.map, 'teacher'));
      if (n) out[n] = String(ntt_get_(row, t.map, 'Contact')).trim();
    });
  } catch (e) {}
  return out;
}

/* ---------------- permissions ---------------- */

function ntt_canEditRow_(user, subject, dd, program, cls, sec) {
  if (user.role === 'ADMIN') return ntt_scopeOk_(user, program, cls, sec);
  if (user.role !== 'HOD') return false;
  if (!NTT_CFG.HOD_RESTRICT_TO_DEPARTMENT) return true;
  const sk = ntt_norm_(subject), dept = ntt_norm_(user.department);
  if (!dept) return false;
  if (sk === dept) return true;
  return !!(dd.subjectDepts[sk] && dd.subjectDepts[sk][dept]);
}

function ntt_rowObj_(t, row, idx, user, dd) {
  const o = { _row: idx + 2 };
  const hl = t._hl || (t._hl = t.headers.map((h, i) => [h, i]).filter(x => x[0]));
  for (let k = 0; k < hl.length; k++) o[hl[k][0]] = ntt_str_(row[hl[k][1]]);
  if (t._iSec === undefined) t._iSec = ntt_col_(t, NTT_SEC_COLS);
  o._sec = t._iSec >= 0 ? ntt_str_(row[t._iSec]) : '';
  o._days = ntt_rowDays_(t, row);        // e.g.  "M-TH,F(P),S(P)"
  o._canEdit = ntt_canEditRow_(user, o['Subject'], dd, o['Program'], o['Class'], o._sec);
  return o;
}

/* ---------------- API ---------------- */

function nttGetBoot(token, force) {
  const user = ntt_session_(token);
  const dd = ntt_ddCached_(!!force);
  const t = ntt_table_(NTT_CFG.TIMETABLE);
  const rows = [];
  const ps = {}, pse = {}, roomSet = {};
  Object.keys(dd.rooms).forEach(r => roomSet[r] = true);

  t.rows.forEach((row, i) => {
    let empty = true;
    for (let c = 0; c < row.length; c++) { if (row[c] !== '' && row[c] !== null) { empty = false; break; } }
    if (empty) return;
    const o = ntt_rowObj_(t, row, i, user, dd);
    rows.push(o);
    const pk = ntt_norm_(o['Program']);
    if (pk) {
      ps[pk] = ps[pk] || {}; pse[pk] = pse[pk] || {};
      if (o['Subject']) ps[pk][o['Subject'].toUpperCase()] = true;
      if (o._sec) pse[pk][o._sec] = true;
    }
    if (o['Room']) roomSet[o['Room']] = true;
  });

  // rooms: "R # 06/55" and "R # 06/55 (EVENING)" are one room -> keep one entry (the name without "(EVENING)" is preferred)
  const roomMap = {};
  Object.keys(roomSet).forEach(n => {
    const k = ntt_roomKey_(n); if (!k) return;
    const ev = ntt_isEveRoomName_(n);
    if (!roomMap[k] || (roomMap[k].eve && !ev)) roomMap[k] = { name: n, eve: ev };
  });
  const roomList = Object.keys(roomMap).map(k => roomMap[k].name).sort(ntt_natSort_);

  // subjects / sections per program = DropdownData (if it has a Program column) + what is already in the timetable
  const progSubjects = {}, progSections = {};
  const pkeys = {};
  Object.keys(ps).forEach(k => pkeys[k] = true);
  Object.keys(dd.byProgram).forEach(k => pkeys[k] = true);
  Object.keys(pkeys).forEach(k => {
    const s = Object.assign({}, ps[k] || {}, (dd.byProgram[k] || {}).subjects || {});
    const c = Object.assign({}, pse[k] || {}, (dd.byProgram[k] || {}).sections || {});
    progSubjects[k] = Object.keys(s).sort();
    progSections[k] = Object.keys(c).sort(ntt_natSort_);
  });

  const teacherList = Object.keys(dd.teachers).map(k => dd.teachers[k].name).sort();
  const subjectTeachers = {};
  Object.keys(dd.subjectTeachers).forEach(sk => {
    subjectTeachers[sk] = dd.subjectTeachers[sk].map(k => dd.teachers[k].name).sort();
  });
  return {
    user: user,
    rows: rows,
    teachers: teacherList,
    subjectTeachers: subjectTeachers,
    subjects: Object.keys(dd.subjects).map(k => dd.subjects[k]).sort(),
    progSubjects: progSubjects,
    progSections: progSections,
    // DropdownData has no Program column: section names saved there are offered to F.A blocks through this pool
    sectionPool: dd.hasProg ? [] : Object.keys(dd.sections).sort(ntt_natSort_),
    rooms: roomList,
    adminFields: ['Subject', 'Teacher', 'Room', 'Period', 'Days', 'Section'],
    hodFields: ['Teacher', 'Days'],
    teacherDepts: (function () { const o = {}; Object.keys(dd.teachers).forEach(k => o[k] = dd.teachers[k].department); return o; })(),
    hodRestrict: NTT_CFG.HOD_RESTRICT_TO_DEPARTMENT,
    subjectInfo: dd.subjectInfo,                                   // SUBJECT -> {department, faculty} (auto-select in "add teacher")
    departments: Object.keys(dd.deptList).map(k => dd.deptList[k]).sort(),
    faculties: Object.keys(dd.facList).map(k => dd.facList[k]).sort(),
    designations: Object.keys(dd.desList).map(k => dd.desList[k]).sort(),
    statuses: Object.keys(dd.statList).map(k => dd.statList[k]).sort(),
    settings: ntt_settings_()
  };
}

/**
 * Apply one change-set to one row (table `t` is read once by the caller).
 */
function ntt_applyOne_(t, dd, user, sheetRow, sNo, changes) {
  const idx = Number(sheetRow) - 2;
  const row = t.rows[idx];
  if (!row) throw new Error('Row not found. Please refresh.');
  if (String(ntt_get_(row, t.map, 'S.No')) !== String(sNo))
    throw new Error('Sheet changed since you loaded it. Please refresh.');

  const curSubject = String(ntt_get_(row, t.map, 'Subject'));
  const rsc = ntt_rowScope_(t, row);
  if (!ntt_canEditRow_(user, curSubject, dd, rsc[0], rsc[1], rsc[2])) throw new Error('You are not allowed to edit this row.');

  const allowed = user.role === 'ADMIN'
    ? ['Subject', 'Teacher', 'Room', 'Period', 'Days', 'Section']
    : ['Teacher', 'Days'];

  const setIdx = (c, value) => ntt_mark_(t, Number(sheetRow), c, value);
  const set = (name, value) => setIdx(t.map[ntt_key_(name)], value);

  Object.keys(changes).forEach(field => {
    if (allowed.indexOf(field) < 0) throw new Error('You cannot change "' + field + '".');
    const val = String(changes[field] == null ? '' : changes[field]).trim();

    if (field === 'Subject' && val === '') {
      // remove subject and everything attached to it (teacher, course name, days)
      set('Subject', ''); set('Course Name', '');
      set('Teacher', ''); set('Designation', ''); set('Status', ''); set('Contact', '');
      const dcx = ntt_dc_(t);
      dcx.theory.concat(dcx.pract).forEach(i => setIdx(i, ''));
      dcx.perDay.forEach(o => setIdx(o.i, ''));
    } else if (field === 'Subject') {
      const sk = ntt_norm_(val);
      if (!dd.subjects[sk]) throw new Error('Unknown subject: ' + val);
      set('Subject', dd.subjects[sk]);
      const cls = ntt_norm_(ntt_get_(row, t.map, 'Class'));
      const hit = dd.courses.filter(c => c.subject === sk && c.cls === cls)[0] ||
                  dd.courses.filter(c => c.subject === sk)[0];
      if (hit) set('Course Name', hit.course);
    } else if (field === 'Teacher') {
      if (val === '') {
        set('Teacher', ''); set('Designation', ''); set('Status', ''); set('Contact', '');
      } else {
        const tk = ntt_norm_(val), info = dd.teachers[tk];
        if (!info) throw new Error('Teacher not in DropdownData: ' + val);
        set('Teacher', info.name);
        set('Designation', info.designation);
        set('Status', info.status);
        set('Contact', (dd._contacts = dd._contacts || ntt_usersContacts_())[tk] || info.contact || '');
      }
    } else if (field === 'Days') {
      const dc = ntt_dc_(t);
      if (!dc.theory.length) throw new Error('Days column not found. Set NTT_DAYS_COLS at the top of the script.');
      const p = ntt_parseDays_(val);
      if (val && !p.th.concat(p.pr).some(x => x)) throw new Error('Could not read days. Use e.g.  M-TH,F(P),S(P)');
      const th = p.th.slice();
      for (let d = 0; d < 7; d++) if (p.pr[d]) th[d] = 0;
      if (dc.pract.length) {
        setIdx(dc.theory[0], ntt_fmtTheory_(th));
        setIdx(dc.pract[0], ntt_fmtPract_(p.pr));
        dc.pract.slice(1).forEach(i => setIdx(i, ''));
      } else {
        setIdx(dc.theory[0], ntt_fmtAll_(th, p.pr));
      }
      dc.theory.slice(1).forEach(i => setIdx(i, ''));
    } else if (field === 'Section') {
      const iS = ntt_col_(t, NTT_SEC_COLS);
      if (iS < 0) throw new Error('Section column not found.');
      setIdx(iS, val);
    } else {
      set(field, val);
    }
  });

  set('Updated By', user.name);
  set('Updated At', new Date());
  return ntt_rowObj_(t, row, idx, user, dd);
}

/**
 * items: [{row, sno, changes:{Subject?,Teacher?,Room?,Period?,Days?,Section?}}]
 * Server enforces permissions regardless of what the browser sends.
 * FAST: only the edited rows are read from the sheet (header + one block of rows); DropdownData comes from the cache.
 */


function nttUpdateRow(token, sheetRow, sNo, changes) {
  const r = nttUpdateMany(token, [{ row: sheetRow, sno: sNo, changes: changes }]);
  return { ok: true, row: r.rows[0] };
}


/**
 * Add a new teacher to DropdownData (listed under the given subject) so he/she appears in the dropdowns.
 * Writes: Subject, Teacher, Designation, Status, Department, Science/Arts and Contact (also saved in Users sheet).
 */
function nttAddTeacher(token, name, designation, status, subject, department, faculty, contact) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN' && user.role !== 'HOD') throw new Error('You are not allowed to add teachers.');
  const nm = String(name || '').trim();
  if (!nm) throw new Error('Teacher name is required.');
  const sj = String(subject || '').trim();
  if (!sj) throw new Error('Subject is required.');
  const ct = String(contact || '').trim();
  if (!ct) throw new Error('Contact is required.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = ntt_table_(NTT_CFG.DROPDOWN);
    const iT = t.map['teacher'], iS = t.map['subject'];
    if (iT === undefined) throw new Error('Teacher column not found in DropdownData.');
    const iFac = ntt_col_(t, NTT_DD_FACULTY_COLS);
    const tk = ntt_norm_(nm), sk = ntt_norm_(sj);
    const same = t.rows.filter(r => ntt_norm_(r[iT]) === tk);
    const already = same.some(r => iS === undefined || ntt_norm_(r[iS]) === sk);
    if (!already) {
      const arr = new Array(t.headers.length).fill('');
      const put = (col, v) => { const i = t.map[ntt_key_(col)]; if (i !== undefined) arr[i] = v; };
      const old = same[0];
      put('Teacher', nm.toUpperCase());
      put('Subject', sj.toUpperCase());
      put('Designation', String(designation || '').trim() || (old && t.map['designation'] !== undefined ? ntt_str_(old[t.map['designation']]) : ''));
      put('Status', String(status || '').trim() || (old && t.map['status'] !== undefined ? ntt_str_(old[t.map['status']]) : ''));
      put('Department', String(department || '').trim() ||
        (user.role === 'HOD' ? user.department : (old && t.map['department'] !== undefined ? ntt_str_(old[t.map['department']]) : '')));
      if (iFac >= 0) arr[iFac] = String(faculty || '').trim().toUpperCase() || (old ? ntt_str_(old[iFac]) : '');
      const iCt = t.map['contact'];
      if (iCt !== undefined) arr[iCt] = ct;
      t.sh.appendRow(arr);
      if (iCt !== undefined) t.sh.getRange(t.sh.getLastRow(), iCt + 1).setNumberFormat('@').setValue(ct);   // keep leading zero
    }
    ntt_saveContact_(nm, ct, String(department || '').trim());
    ntt_ddDrop_();
    ntt_log_(user, [['ADD TEACHER', nm.toUpperCase() + ' (' + sj + ')']]);
    return { ok: true, name: nm.toUpperCase() };
  } finally {
    lock.releaseLock();
  }
}

// store the contact of a teacher in the Users sheet (row without user name / password cannot log in)
function ntt_saveContact_(nm, contact, department) {
  try {
    const t = ntt_table_(NTT_CFG.USERS);
    const iT = t.map['teacher'], iC = t.map['contact'];
    if (iT === undefined || iC === undefined) return;
    const tk = ntt_norm_(nm);
    for (let i = 0; i < t.rows.length; i++) {
      if (ntt_norm_(t.rows[i][iT]) === tk) {
        if (!ntt_str_(t.rows[i][iC])) t.sh.getRange(i + 2, iC + 1).setNumberFormat('@').setValue(contact);
        return;
      }
    }
    const arr = new Array(t.headers.length).fill('');
    arr[iT] = nm.toUpperCase(); arr[iC] = contact;
    const iD = t.map['department']; if (iD !== undefined) arr[iD] = department || '';
    t.sh.appendRow(arr);
    t.sh.getRange(t.sh.getLastRow(), iC + 1).setNumberFormat('@').setValue(contact);
  } catch (e) {}
}

/**
 * Delete a teacher: the teacher is removed from the dropdown list (Teacher / Designation / Status / Contact cells of
 * the teacher's rows in DropdownData are cleared, subjects stay) and is unassigned from every timetable row.
 * Main admin: any teacher.  Scoped admin: only a teacher who has no periods outside his / her own part of the timetable.
 * HOD: only teachers of the HOD's own department (when HOD_RESTRICT_TO_DEPARTMENT is true).
 * The Users sheet (logins) is not touched.
 */
function nttDeleteTeacher(token, name) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN' && user.role !== 'HOD') throw new Error('You are not allowed to delete teachers.');
  const nm = ntt_norm_(name);
  if (!nm) throw new Error('Choose a teacher.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const dt = ntt_table_(NTT_CFG.DROPDOWN);
    const iT = dt.map['teacher'];
    if (iT === undefined) throw new Error('Teacher column not found in DropdownData.');
    const iDept = dt.map['department'];
    const mine = [];
    dt.rows.forEach((r, i) => { if (ntt_norm_(r[iT]) === nm) mine.push(i); });
    if (!mine.length) throw new Error('Teacher not found in DropdownData.');
    if (user.role === 'ADMIN' && (user.scope || 'ALL') !== 'ALL') {
      const tt0 = ntt_table_(NTT_CFG.TIMETABLE), iT0 = tt0.map['teacher'];
      if (iT0 !== undefined && tt0.rows.some(r => {
        if (ntt_norm_(r[iT0]) !== nm) return false;
        const sc = ntt_rowScope_(tt0, r);
        return !ntt_scopeOk_(user, sc[0], sc[1], sc[2]);
      })) throw new Error('This teacher also teaches in a part of the timetable that is not yours. Only the main admin can delete her.');
    }
    if (user.role === 'HOD' && NTT_CFG.HOD_RESTRICT_TO_DEPARTMENT) {
      const dept = ntt_norm_(user.department);
      const ok = dept && iDept !== undefined && mine.every(i => ntt_norm_(dt.rows[i][iDept]) === dept);
      if (!ok) throw new Error('You can delete only teachers of your own department.');
    }
    const dCols = ['teacher', 'designation', 'status', 'contact'].map(k => dt.map[k]).filter(c => c !== undefined);
    mine.forEach(i => dCols.forEach(c => ntt_mark_(dt, i + 2, c, '')));
    ntt_flush_(dt);
    ntt_ddDrop_();

    const tt = ntt_table_(NTT_CFG.TIMETABLE);
    const iTT = tt.map['teacher'];
    let cleared = 0;
    if (iTT !== undefined) {
      const tCols = ['teacher', 'designation', 'status', 'contact'].map(k => tt.map[k]).filter(c => c !== undefined);
      const iBy = tt.map['updated by'], iAt = tt.map['updated at'];
      const now = new Date();
      tt.rows.forEach((r, i) => {
        if (ntt_norm_(r[iTT]) !== nm) return;
        tCols.forEach(c => ntt_mark_(tt, i + 2, c, ''));
        ntt_mark_(tt, i + 2, iBy, user.name);
        ntt_mark_(tt, i + 2, iAt, now);
        cleared++;
      });
      ntt_flush_(tt);
    }
    ntt_log_(user, [['DELETE TEACHER', nm + ' · ' + cleared + ' slot(s) cleared']]);
    return { ok: true, cleared: cleared };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- rooms / section names (DropdownData columns) ---------------- */

// returns the index of the column (creates it with `header` when missing)
function ntt_ensureCol_(t, aliases, header) {
  let i = ntt_col_(t, aliases);
  if (i < 0) {
    i = t.sh.getLastColumn();
    t.sh.getRange(1, i + 1).setValue(header);
  }
  return i;
}

// writes `value` under the last value of a DropdownData column (the column is created if missing). No lock / session here.
// keyFn decides when two values are "the same" (default: case/space-insensitive; rooms ignore "(EVENING)").
function ntt_colAppend_(aliases, header, value, keyFn) {
  const kf = keyFn || ntt_norm_;
  const t = ntt_table_(NTT_CFG.DROPDOWN);
  const iR = ntt_ensureCol_(t, aliases, header);
  let last = -1, found = '';
  t.rows.forEach((r, i) => {
    const v = ntt_str_(r[iR]);
    if (v) { last = i; if (!found && kf(v) === kf(value)) found = v; }
  });
  if (found) return found;
  const sheetRow = last + 3;                 // data row i is sheet row i+2; the next free cell is one below the last value
  const maxR = t.sh.getMaxRows();
  if (sheetRow > maxR) t.sh.insertRowsAfter(maxR, sheetRow - maxR);
  t.sh.getRange(sheetRow, iR + 1).setValue(value);
  ntt_ddDrop_();
  return value;
}
function ntt_addRoomInternal_(nm) { return ntt_colAppend_(NTT_DD_ROOM_COLS, 'Room', nm, ntt_roomKey_); }

/**
 * Add a new room to the Room column of DropdownData (written under the last existing room).
 * "R # 06/55 (EVENING)" is the same room as "R # 06/55": the existing name is returned instead of adding a duplicate.
 * Every admin type may add a room.
 */
function nttAddRoom(token, name) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN') throw new Error('Only admin can add rooms.');
  const nm = String(name || '').trim();
  if (!nm) throw new Error('Room name is required.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const res = ntt_addRoomInternal_(nm);
    SpreadsheetApp.flush();
    ntt_log_(user, [['ADD ROOM', res]]);
    return { ok: true, name: res };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Delete a room from the room list (every cell of the Room column of DropdownData that holds this room,
 * with or without "(EVENING)", is cleared). A room that is still used in the timetable cannot be deleted:
 * change those timetable rows to another room first. Main admin only.
 */
function nttDeleteRoom(token, name) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN') throw new Error('Only admin can delete rooms.');
  ntt_needFull_(user);
  const key = ntt_roomKey_(name);
  if (!key) throw new Error('Choose a room.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // refuse when the room is still used in the timetable
    const tt = ntt_table_(NTT_CFG.TIMETABLE);
    const iRm = tt.map['room'];
    if (iRm !== undefined) {
      let used = 0;
      tt.rows.forEach(r => { if (ntt_roomKey_(r[iRm]) === key) used++; });
      if (used) throw new Error('Room "' + String(name).trim() + '" is used in ' + used + ' timetable row(s). Change those rows to another room first.');
    }
    const dt = ntt_table_(NTT_CFG.DROPDOWN);
    const iR = ntt_col_(dt, NTT_DD_ROOM_COLS);
    let removed = 0;
    if (iR >= 0) {
      dt.rows.forEach((r, i) => {
        if (ntt_roomKey_(r[iR]) === key) { ntt_mark_(dt, i + 2, iR, ''); removed++; }
      });
      ntt_flush_(dt);
    }
    ntt_ddDrop_();
    SpreadsheetApp.flush();
    ntt_log_(user, [['DELETE ROOM', String(name).trim()]]);
    return { ok: true, removed: removed };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Add a new section NAME to the section dropdown lists (used by the "＋ Add new section name…" option,
 * mainly for F.A blocks). If DropdownData has a Program column, a row (Program + Section only) is appended so the name
 * belongs to that program; otherwise it is written under the last value of the Section column.
 */
function nttAddSectionName(token, program, name) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN') throw new Error('Only admin can add section names.');
  ntt_needProg_(user, program);
  const nm = String(name || '').trim();
  if (!nm) throw new Error('Section name is required.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = ntt_table_(NTT_CFG.DROPDOWN);
    const iProg = ntt_col_(t, ['Program']);
    if (iProg >= 0) {
      const iS = ntt_ensureCol_(t, NTT_DD_SEC_COLS, 'Section');
      const pk = ntt_norm_(program), sk = ntt_norm_(nm);
      const exists = t.rows.some(r => ntt_norm_(r[iProg]) === pk && ntt_norm_(r[iS]) === sk);
      if (!exists) {
        const arr = new Array(Math.max(t.sh.getLastColumn(), iS + 1, iProg + 1)).fill('');
        arr[iProg] = String(program || '').trim();
        arr[iS] = nm;
        t.sh.appendRow(arr);
      }
      ntt_ddDrop_();
    } else {
      ntt_colAppend_(NTT_DD_SEC_COLS, 'Section', nm);
    }
    SpreadsheetApp.flush();
    return { ok: true, name: nm };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Add a new section to a class. New timetable rows are created for the section:
 *  - copyFrom given : the subjects / periods / course names of that section are copied (teachers and days left empty)
 *  - copyFrom empty : 5 empty rows (periods 1-5) are created
 * room (optional) is set on all the new rows.
 * Scoped admins: allowed for the programs of their own timetable (the shift of the new section is set afterwards
 * through Settings -> Evening sections).
 */
function nttAddSection(token, program, cls, section, copyFrom, room) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN') throw new Error('Only admin can add sections.');
  ntt_needProg_(user, program);
  const sec = String(section || '').trim();
  if (!sec) throw new Error('Section name is required.');
  const rm = String(room || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = ntt_table_(NTT_CFG.TIMETABLE);
    const iP = ntt_col_(t, ['Program']), iC = ntt_col_(t, ['Class']), iS = ntt_col_(t, NTT_SEC_COLS);
    if (iP < 0 || iC < 0 || iS < 0) throw new Error('Program / Class / Section column not found in the timetable sheet.');
    const pk = ntt_norm_(program), ck = ntt_norm_(cls), sk = ntt_norm_(sec);
    const cur = t.rows.filter(r => ntt_norm_(r[iP]) === pk && ntt_norm_(r[iC]) === ck);
    if (!cur.length) throw new Error('Class not found. Please refresh.');
    if (cur.some(r => ntt_norm_(r[iS]) === sk)) throw new Error('This section already exists in this class.');
    const cf = ntt_norm_(copyFrom);
    const src = cf ? cur.filter(r => ntt_norm_(r[iS]) === cf) : [];
    if (cf && !src.length) throw new Error('Section to copy from was not found.');

    const iNo = t.map['s.no'], iPer = t.map['period'], iRoom = t.map['room'];
    const iSub = t.map['subject'], iCourse = t.map['course name'];
    const iBy = t.map['updated by'], iAt = t.map['updated at'];
    let max = 0;
    if (iNo !== undefined) t.rows.forEach(r => { const v = Number(r[iNo]); if (!isNaN(v) && v > max) max = v; });
    const dc = ntt_dc_(t);
    const clearCols = ['teacher', 'designation', 'status', 'contact'].map(k => t.map[k]).filter(c => c !== undefined);
    const now = new Date();
    const make = base => {
      const r = base.slice();
      if (iNo !== undefined) r[iNo] = ++max;
      r[iS] = sec;
      clearCols.forEach(c => r[c] = '');
      dc.theory.concat(dc.pract).forEach(c => r[c] = '');
      dc.perDay.forEach(o => r[o.i] = '');
      if (iRoom !== undefined) r[iRoom] = rm;
      if (iBy !== undefined) r[iBy] = user.name;
      if (iAt !== undefined) r[iAt] = now;
      return r;
    };
    const fresh = [];
    if (src.length) {
      src.forEach(b => fresh.push(make(b)));
    } else {
      const base = cur[0];
      const orig = iPer !== undefined ? String(base[iPer]) : '';
      for (let p = 1; p <= 5; p++) {
        const r = make(base);
        if (iSub !== undefined) r[iSub] = '';
        if (iCourse !== undefined) r[iCourse] = '';
        if (iPer !== undefined) r[iPer] = /\d+/.test(orig) ? orig.replace(/\d+/, String(p)) : String(p);
        fresh.push(r);
      }
    }
    if (rm) { try { ntt_addRoomInternal_(rm); } catch (e) {} }
    const first = t.rows.length + 2, n = fresh.length;
    const maxR = t.sh.getMaxRows();
    if (first + n - 1 > maxR) t.sh.insertRowsAfter(maxR, first + n - 1 - maxR);
    t.sh.getRange(first, 1, n, t.headers.length).setValues(fresh);
    SpreadsheetApp.flush();
    const dd = { subjectDepts: {}, subjects: {}, teachers: {}, courses: [] };
    const out = fresh.map((r, k) => { t.rows.push(r); return ntt_rowObj_(t, r, t.rows.length - 1, user, dd); });
    ntt_log_(user, [['ADD SECTION', program + ' ' + cls + ' / ' + sec + (cf ? ' (copied from ' + copyFrom + ')' : ' (blank)')]]);
    return { ok: true, rows: out };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Remove a section of a class: every timetable row of that section is deleted from the sheet.
 * The only section of a class cannot be removed (the class would disappear and sections could not be added back).
 * Scoped admins can remove only sections of their own part (morning / evening / F.A).
 * Returns the sheet row numbers that were deleted.
 */
function nttDeleteSection(token, program, cls, section) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN') throw new Error('Only admin can remove sections.');
  ntt_needScope_(user, program, cls, section);
  const sec = String(section || '').trim();
  if (!sec) throw new Error('Choose a section.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = ntt_table_(NTT_CFG.TIMETABLE);
    const iP = ntt_col_(t, ['Program']), iC = ntt_col_(t, ['Class']), iS = ntt_col_(t, NTT_SEC_COLS);
    if (iP < 0 || iC < 0 || iS < 0) throw new Error('Program / Class / Section column not found in the timetable sheet.');
    const pk = ntt_norm_(program), ck = ntt_norm_(cls), sk = ntt_norm_(sec);
    const cur = [];
    t.rows.forEach((r, i) => { if (ntt_norm_(r[iP]) === pk && ntt_norm_(r[iC]) === ck) cur.push(i); });
    if (!cur.length) throw new Error('Class not found. Please refresh.');
    const secs = {};
    cur.forEach(i => { const v = ntt_norm_(t.rows[i][iS]); if (v) secs[v] = true; });
    if (!secs[sk]) throw new Error('Section not found. Please refresh.');
    if (Object.keys(secs).length < 2) throw new Error('This is the only section of the class and cannot be removed.');
    const sheetRows = cur.filter(i => ntt_norm_(t.rows[i][iS]) === sk).map(i => i + 2).sort((a, b) => a - b);
    // delete from the bottom up, contiguous blocks at once
    let end = sheetRows.length - 1;
    while (end >= 0) {
      let start = end;
      while (start > 0 && sheetRows[start - 1] === sheetRows[start] - 1) start--;
      t.sh.deleteRows(sheetRows[start], end - start + 1);
      end = start - 1;
    }
    SpreadsheetApp.flush();
    ntt_log_(user, [['DELETE SECTION', program + ' ' + cls + ' / ' + sec + ' · ' + sheetRows.length + ' row(s)']]);
    return { ok: true, rows: sheetRows };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Add another slot (row) of the same subject in the same section, in a different period with its own days.
 * e.g. PHYSICS F,S in period 1  +  PHYSICS W,TH in period 2.
 */
function nttAddSlot(token, sheetRow, sNo, period, days) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN') throw new Error('Only admin can add period slots.');
  const n = Number(period);
  if (!(n >= 1 && n <= 12)) throw new Error('Choose a period.');
  if (!String(days || '').trim()) throw new Error('Enter the days for the new slot.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = ntt_table_(NTT_CFG.TIMETABLE);
    const idx = Number(sheetRow) - 2, src = t.rows[idx];
    if (!src) throw new Error('Row not found. Please refresh.');
    if (String(ntt_get_(src, t.map, 'S.No')) !== String(sNo)) throw new Error('Sheet changed since you loaded it. Please refresh.');
    const sc0 = ntt_rowScope_(t, src);
    ntt_needScope_(user, sc0[0], sc0[1], sc0[2]);            // checked BEFORE anything is written
    const iNo = t.map['s.no'];
    let max = 0;
    if (iNo !== undefined) t.rows.forEach(r => { const v = Number(r[iNo]); if (!isNaN(v) && v > max) max = v; });
    const copy = src.slice();
    if (iNo !== undefined) copy[iNo] = max + 1;
    const dc = ntt_dc_(t);
    dc.theory.concat(dc.pract).forEach(i => copy[i] = '');
    dc.perDay.forEach(o => copy[o.i] = '');
    const orig = String(ntt_get_(src, t.map, 'Period'));
    const per = /\d+/.test(orig) ? orig.replace(/\d+/, String(n)) : String(n);
    t.sh.appendRow(copy);
    t.rows.push(copy);                       // the new row is the last row of the sheet
    const ni = t.rows.length - 1;
    const newNo = iNo !== undefined ? String(max + 1) : '';
    const dd = { subjectDepts: {}, subjects: {}, teachers: {}, courses: [] };   // admin only: no lookups needed
    const row = ntt_applyOne_(t, dd, user, ni + 2, newNo, { Period: per, Days: days });
    ntt_flush_(t);
    ntt_log_(user, [['ADD SLOT', ntt_ctx_(t, ni + 2) + ' | Period ' + n + ' · ' + days]]);
    return { ok: true, row: row };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Delete a slot (row) - allowed only when the same subject still has another slot in that section.
 */
function nttDeleteSlot(token, sheetRow, sNo) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN') throw new Error('Only admin can delete slots.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = ntt_table_(NTT_CFG.TIMETABLE);
    const idx = Number(sheetRow) - 2, row = t.rows[idx];
    if (!row) throw new Error('Row not found. Please refresh.');
    if (String(ntt_get_(row, t.map, 'S.No')) !== String(sNo)) throw new Error('Sheet changed since you loaded it. Please refresh.');
    const sc0 = ntt_rowScope_(t, row);
    ntt_needScope_(user, sc0[0], sc0[1], sc0[2]);
    const subj = ntt_norm_(ntt_get_(row, t.map, 'Subject'));
    if (!subj) throw new Error('Nothing to delete here.');
    const iP = ntt_col_(t, ['Program']), iC = ntt_col_(t, ['Class']), iS = ntt_col_(t, NTT_SEC_COLS);
    const key = r => [iP, iC, iS].map(i => i >= 0 ? ntt_norm_(r[i]) : '').join('|');
    const others = t.rows.filter((r, i) => i !== idx && ntt_norm_(ntt_get_(r, t.map, 'Subject')) === subj && key(r) === key(row));
    if (!others.length)
      throw new Error('This is the only slot of this subject in the section. To remove the subject, choose "none" in its subject dropdown.');
    ntt_log_(user, [['DELETE SLOT', ntt_ctx_(t, Number(sheetRow))]]);
    t.sh.deleteRow(Number(sheetRow));
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}



/* =====================================================================
   NTT GS ADDITIONS  -  file: "naila tt gs.gs"
   PART 1: paste this whole block at the very END of the file.
   PART 2: small edits inside existing functions (listed at the bottom).
===================================================================== */

/* ---------------- PART 1: new code (paste at the end) ---------------- */

const NTT_LOG_SHEET = 'NTT Log';
const NTT_BK_PREFIX = 'TT Backup ';
const NTT_BK_KEEP = 10;                       // newest backups kept; older ones are deleted automatically

/* ---- change log ---- */
function ntt_log_(user, entries) {
  try {
    if (!entries || !entries.length) return;
    const ss = ntt_ss_();
    let sh = ss.getSheetByName(NTT_LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(NTT_LOG_SHEET);
      sh.getRange(1, 1, 1, 5).setValues([['Time', 'User', 'Role', 'Action', 'Detail']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const now = new Date(), role = user.label || user.role;
    const rows = entries.map(e => [now, user.name, role, e[0], String(e[1]).slice(0, 500)]);
    const first = sh.getLastRow() + 1, need = first + rows.length - 1, maxR = sh.getMaxRows();
    if (need > maxR) sh.insertRowsAfter(maxR, need - maxR);
    sh.getRange(first, 1, rows.length, 5).setValues(rows);
    if (sh.getLastRow() > 20000) sh.deleteRows(2, 5000);       // keep the log from growing forever
  } catch (e) {}                                                // logging must never break an edit
}
function ntt_ctx_(t, sheetRow) {
  const row = t.rows[sheetRow - 2];
  if (!row) return 'row ' + sheetRow;
  const sc = ntt_rowScope_(t, row);
  return sc.filter(Boolean).join(' ') + ' · ' + ntt_str_(ntt_get_(row, t.map, 'Subject'));
}
function ntt_fieldVal_(t, row, f) {
  if (f === 'Days') return ntt_rowDays_(t, row);
  if (f === 'Section') { const i = ntt_col_(t, NTT_SEC_COLS); return i >= 0 ? ntt_str_(row[i]) : ''; }
  return ntt_str_(ntt_get_(row, t.map, f));
}
function ntt_snap_(t, sheetRow, changes) {
  const row = t.rows[sheetRow - 2], o = { ctx: ntt_ctx_(t, sheetRow), vals: {} };
  if (row) Object.keys(changes).forEach(f => { o.vals[f] = ntt_fieldVal_(t, row, f); });
  return o;
}
function ntt_diffText_(snap, changes) {
  const parts = [];
  Object.keys(changes).forEach(f => {
    const o = snap.vals[f] || '', n = String(changes[f] == null ? '' : changes[f]).trim();
    if (o !== n) parts.push(f + ': ' + (o || '—') + ' → ' + (n || '—'));
  });
  return parts.length ? snap.ctx + ' | ' + parts.join('; ') : '';
}

/* main admin: last `limit` log lines, newest first */
function nttGetLog(token, limit) {
  const user = ntt_session_(token);
  ntt_needFull_(user);
  const sh = ntt_ss_().getSheetByName(NTT_LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  const n = Math.min(Math.max(Number(limit) || 300, 1), 1000), last = sh.getLastRow();
  const first = Math.max(2, last - n + 1);
  const vals = sh.getRange(first, 1, last - first + 1, 5).getValues();
  return { ok: true, rows: vals.reverse().map(r => [ntt_str_(r[0]), ntt_str_(r[1]), ntt_str_(r[2]), ntt_str_(r[3]), ntt_str_(r[4])]) };
}

/* ---- backup of the timetable sheet (hidden copy "TT Backup yyyy-MM-dd HHmmss"; the newest 10 are kept) ---- */
function nttBackup(token, reason) {
  const user = ntt_session_(token);
  if (user.role !== 'ADMIN') throw new Error('Only admin can make backups.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ntt_ss_(), src = ss.getSheetByName(NTT_CFG.TIMETABLE);
    if (!src) throw new Error('Sheet not found: ' + NTT_CFG.TIMETABLE);
    const name = NTT_BK_PREFIX + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmmss');
    const cp = src.copyTo(ss);
    cp.setName(name);
    try { cp.hideSheet(); } catch (e) {}
    const old = ss.getSheets().filter(s => s.getName().indexOf(NTT_BK_PREFIX) === 0)
      .sort((a, b) => a.getName() < b.getName() ? 1 : -1);          // newest first
    old.slice(NTT_BK_KEEP).forEach(s => { try { ss.deleteSheet(s); } catch (e) {} });
    SpreadsheetApp.flush();
    ntt_log_(user, [['BACKUP', name + (reason ? ' — ' + String(reason).slice(0, 120) : '')]]);
    return { ok: true, name: name };
  } finally {
    lock.releaseLock();
  }
}


/* ---- change-log version of nttUpdateMany (from the additions) ---- */
/* ---- B. REPLACE the whole function nttUpdateMany with this one (adds the change log) ---- */
function nttUpdateMany(token, items) {
  const user = ntt_session_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const needDd = user.role !== 'ADMIN' ||
      items.some(it => it.changes && (('Subject' in it.changes) || ('Teacher' in it.changes)));
    const dd = needDd ? ntt_ddCached_(false) : { subjectDepts: {}, subjects: {}, teachers: {}, courses: [] };
    const t = ntt_tablePart_(NTT_CFG.TIMETABLE, items.map(it => it.row));
    const snaps = items.map(it => ntt_snap_(t, it.row, it.changes || {}));
    const out = items.map(it => ntt_applyOne_(t, dd, user, it.row, it.sno, it.changes || {}));
    ntt_flush_(t);
    const entries = [];
    items.forEach((it, i) => { const d = ntt_diffText_(snaps[i], it.changes || {}); if (d) entries.push(['EDIT', d]); });
    ntt_log_(user, entries);
    return { ok: true, rows: out };
  } finally {
    lock.releaseLock();
  }
}

