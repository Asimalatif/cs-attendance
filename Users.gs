/* =====================================================================================================
   Users.gs (v8.19) — logins of the attendance software from the app: add a new login (e.g. a computer
   assistant), change one, set a new password, switch a login off, and switch SERVICES on / off per login.
   Page: users.html (College screen → Logins & services). Main admin only.

   The Users sheet (Users workbook) gets one more column by itself: "Services" — a comma list such as
   "admission, idcards". A login with a service may use that module like the admin (Code.gs sets GRANT_ for
   the actions of that module only). "blocked" in the list = the login cannot sign in.
   ===================================================================================================== */
var SVC_LIST = [
  { key: 'admission',      prefix: 'adm_', page: 'admission.html',     label: 'Admissions',               about: 'Add a new student or change an admission; merit lists, seats, From portal.' },
  { key: 'idcards',        prefix: 'idc_', page: 'idcards.html',       label: 'Student ID (chest) cards', about: 'Make and print ID cards, fix photos, approve students\' corrections.' },
  { key: 'certificates',   prefix: 'pc_',  page: 'certificates.html',  label: 'Provisional certificates', about: 'Search, correct and print provisional certificates.' },
  { key: 'board',          prefix: 'br_',  page: 'board.html',         label: 'Board results',            about: 'Read the board gazette PDF into a result sheet.' },
  { key: 'duties',         prefix: 'dut_', page: 'duties.html',        label: 'Exam duty lists',          about: 'Teacher-, room-, date-, class- and subject-wise duty PDFs.' },
  { key: 'collegereports', prefix: 'rpt_', page: 'reports.html',       label: 'College (strength) reports', about: 'Strength, quota, section-wise, subjects and other college reports.' },
  { key: 'shortage',       prefix: 'shr_', page: 'shortage.html',      label: 'Attendance shortage',      about: 'Shortage cases of any month or the whole session.' },
  { key: 'leaves',         prefix: 'lv_',  page: 'leaves.html',        label: 'Approved leaves',          about: 'Enter students\' approved leaves (works without internet).' },
  { key: 'recordsummary',  prefix: 'rsm_', page: 'recordsummary.html', label: 'Marked-Missed entries',    about: 'Which section and subject has attendance / tests marked.' },
  { key: 'admintools',     prefix: 'at_',  page: 'admintools.html',    label: 'Admin tools',              about: 'Admission slips, orientation lists, room tags, workload, stopping merit.' }
];
var USR_HEAD = 'Services';
/* v8.36 SUB-SERVICES: parts of a service that can be given one by one ("admission.new", "admission.edit" … in the Services column).
   A service with none of its parts listed = the whole service (as before). acts = the server actions only that part may use
   ("adm_save:new" = adm_save without ctx, "at_run:merit" = at_run with tool merit); every other action of the module stays open
   (reading lists, settings the page needs). tab = the page tab hidden when the part is not given. */
var SVC_SUBS = {
  admission: [
    { k: 'new',   label: 'New admission (New tab)',                 tab: 'form',  acts: ['adm_save:new', 'adm_check:new', 'adm_next_roll', 'adm_fee_receipt'] },
    { k: 'edit',  label: 'Find & edit an admission, record PDF',    tab: 'find',  acts: ['adm_save:edit', 'adm_check:edit', 'adm_full'] },
    { k: 'portal',label: 'From portal: add new portal students, update from portal', tab: 'pnew', acts: ['adm_new', 'adm_add_new', 'adm_refresh', 'adm_keep_list', 'adm_keep_save', 'adm_portal'] },
    { k: 'seats', label: 'Seats (view)',                            tab: 'seats', acts: [] },
    { k: 'merit', label: 'Merit lists (confirm, last date, alerts)',tab: 'merit', acts: ['adm_oml_get', 'adm_oml_save', 'adm_oml_confirm', 'adm_portal_quota', 'adm_alerts_save', 'adm_last_date'] },
    { k: 'fix',   label: 'Wrong on portal list',                    tab: 'fix',   acts: ['adm_fixlist'] },
    { k: 'sset',  label: 'Seat settings (change seats, migrated)',  tab: 'sset',  acts: ['adm_seats_save', 'adm_migrated', 'adm_mig_import'] },
    { k: 'wait',  label: 'Waiting (saves kept on the phone)',       tab: 'wait',  acts: [] },
    { k: 'slips', label: 'Slips & quick edit (change quota, section, Enroll#, program, subjects)', tab: 'slips', acts: ['at_meta', 'at_strength', 'at_savestu'] }
  ],
  idcards: [
    { k: 'print',    label: 'Make & print ID cards (PDF), mark printed', acts: ['idc_mark', 'idc_white_put'] },
    { k: 'photos',   label: 'Fix / add photos',                          acts: ['idc_photo_save'] },
    { k: 'requests', label: 'Students\' correction requests (approve / reject)', acts: ['idc_reqs', 'idc_req_do', 'idc_check_set'] },
    { k: 'settings', label: 'Card settings & principal signature',       acts: ['idc_settings', 'idc_sign'] }
  ],
  certificates: [
    { k: 'print',  label: 'Print certificates (PDF / Drive)',             acts: ['pc_pdf'] },
    { k: 'edit',   label: 'Correct a student',                            acts: ['pc_save'] },
    { k: 'strike', label: 'Strike off left students',                     acts: ['pc_strike', 'pc_strikes'] },
    { k: 'link',   label: 'Fill Enroll#, DOB and grade (link)',           acts: ['pc_link'] }
  ],
  board: [
    { k: 'save',    label: 'Read a gazette PDF and save the result sheet', acts: ['br_begin', 'br_rows', 'br_finish'] },
    { k: 'prepare', label: 'Prepare provisional certificates from a result', acts: ['pc_prepare'] }
  ],
  duties: [
    { k: 'make',   label: 'Make duty lists, PDFs, Google Sheet',          acts: ['dut_sheet', 'dut_students', 'dut_stu_cols', 'dut_stu_subj'] },
    { k: 'save',   label: 'Save lists and name fixes on the server',      acts: ['dut_save', 'dut_names_save', 'dut_fix_save'] },
    { k: 'delete', label: 'Delete saved lists',                           acts: ['dut_delete'] }
  ],
  collegereports: [
    { k: 'run',      label: 'Make reports, Excel downloads',              acts: ['rpt_run', 'rpt_xlsx', 'rpt_xlsx_all'] },
    { k: 'sessions', label: 'Change the sessions list',                   acts: ['rpt_sessions'] }
  ],
  shortage: [
    { k: 'view',     label: 'See and print shortage cases',               acts: ['shr_month'] },
    { k: 'save',     label: 'Save the list to the month workbook',        acts: ['shr_write'] },
    { k: 'settings', label: 'Session and default threshold',              acts: ['shr_cfg_set'] }
  ],
  leaves: [
    { k: 'view',  label: 'See leaves, records and reports',               acts: ['lv_attendance'] },
    { k: 'enter', label: 'Enter / change approved leaves',                acts: ['lv_sync'] }
  ],
  recordsummary: [],
  admintools: [
    { k: 'auto',    label: 'Auto timetable',                 tab: 'auto',    acts: ['at_run:auto', 'at_sheet'] },
    { k: 'work',    label: 'Teacher workload, designations', tab: 'work',    acts: ['at_run:workload', 'at_run:workloadxlsx', 'at_desig', 'at_desig_save'] },
    { k: 'merit',   label: 'Stopping merit',                 tab: 'merit',   acts: ['at_run:merit'] },
    { k: 'slips',   label: 'Admission slips (and editing them)', tab: 'slips', acts: ['at_savestu'] },
    { k: 'orient',  label: 'Orientation lists',              tab: 'orient',  acts: [] },
    { k: 'tags',    label: 'Room tags',                      tab: 'tags',    acts: ['at_roomtags'] },
    { k: 'morning', label: 'Morning classes report and fees',tab: 'morning', acts: ['at_morning', 'at_mfee'] }
  ]
};
/** the action key used in SVC_SUBS for this request */
function svcActKey_(req) {
  var a = String(req.action || '');
  if (a === 'adm_save' || a === 'adm_check') return a + (req.ctx ? ':edit' : ':new');
  if (a === 'at_run') return 'at_run:' + String(req.tool || '');
  return a;
}
/** null = allowed; else the label of the part this login was not given */
function svcSubDeny_(user, key, req) {
  var subs = SVC_SUBS[key] || [], mine = (user.svc || []).filter(function (t) { return t.indexOf(key + '.') === 0; });
  if (!mine.length || !subs.length) return null;                                  // no parts listed = the whole service
  var ak = svcActKey_(req), owners = subs.filter(function (x) { return x.acts.indexOf(ak) >= 0; });
  if (!owners.length) return null;                                                // an action every part needs
  return owners.some(function (x) { return mine.indexOf(key + '.' + x.k) >= 0; }) ? null : owners.map(function (x) { return x.label; }).join(' / ');
}

/** "admission, ID cards ; blocked" -> ['admission', 'idcards', 'blocked'] (only known keys) */
function svcParse_(txt) {
  var ok = { blocked: true }; SVC_LIST.forEach(function (s) { ok[s.key] = true; });
  USR_SCREENS.forEach(function (x) { ok['hide' + x.key] = true; });   // v8.35 screens switched off for this login
  Object.keys(SVC_SUBS).forEach(function (k) { SVC_SUBS[k].forEach(function (x) { ok[k + '.' + x.k] = true; }); });   // v8.36 parts of a service
  var out = [];
  String(txt || '').toLowerCase().split(/[,;|\n]+/).forEach(function (w) {
    w = w.replace(/[^a-z.]/g, '');
    if (w === 'chestcards' || w === 'idcard') w = 'idcards';
    if (ok[w] && out.indexOf(w) < 0) out.push(w);
  });
  return out;
}
var SVC_ALSO = { idc_white_get: 'admission', pc_prepare: 'board',
  at_meta: 'admission', at_strength: 'admission', at_savestu: 'admission' };   // v8.38 Admissions → Slips & quick edit (admintools.html?only=slips)   // a page that also calls one action of another module
/** the services that open a request: 'idc_list' -> ['idcards'] ([] = no service, normal rules) */
function svcForAction_(a) {
  a = String(a || ''); var out = [];
  for (var i = 0; i < SVC_LIST.length; i++) if (a.indexOf(SVC_LIST[i].prefix) === 0) out.push(SVC_LIST[i].key);
  if (SVC_ALSO[a]) out.push(SVC_ALSO[a]);
  return out;
}
/** v8.35 screens of the app that can be switched off per login or per role */
var USR_SCREENS = [
  { key: 'att', label: 'Attendance tab' }, { key: 'mt', label: 'Monthly test' },
  { key: 'gt', label: 'Grand test' }, { key: 'dt', label: 'Detention' }
];
var USR_ROLE_PROP = 'USR_ROLE_HIDE';
function usrRoleHide_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(USR_ROLE_PROP) || '{}') || {}; } catch (e) { return {}; } }
/** the services this login can open + 'hideatt' … for the screens switched off (her own + her role's); the admin has everything */
function svcMine_(user) {
  if (String(user.username || '').trim().toLowerCase() === ADMIN_USER) return [];
  var out = (user.svc || []).filter(function (k) { return k !== 'blocked'; });
  var rh = usrRoleHide_()[String(user.role || '').trim().toUpperCase()] || [];
  rh.forEach(function (k) { if (out.indexOf('hide' + k) < 0) out.push('hide' + k); });
  return out;
}

function usrRoute_(user, req) {
  try {
    if (String(user.username || '').trim().toLowerCase() !== ADMIN_USER) return { ok: false, error: 'Only the main admin can add or change logins.' };
    switch (String(req.action || '')) {
      case 'usr_list': return usrList_();
      case 'usr_save': return usrSave_(req);
      case 'usr_set':  return usrSet_(req);
      case 'usr_role_set': {                                            // v8.35 { ROLE: ['att', 'mt', 'gt', 'dt'] }
        var keys = USR_SCREENS.map(function (x) { return x.key; }), clean = {};
        Object.keys(req.rules || {}).forEach(function (r) { var L = (req.rules[r] || []).filter(function (k) { return keys.indexOf(k) >= 0; }); var R = String(r).trim().toUpperCase(); if (R && L.length) clean[R] = L; });
        PropertiesService.getScriptProperties().setProperty(USR_ROLE_PROP, JSON.stringify(clean));
        return { ok: true, rules: clean };
      }
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) { return { ok: false, error: String(e && e.message || e) }; }
}

/* ---------- the sheet ---------- */
function usrSheet_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users');
  if (!sh) throw new Error('Users sheet not found.');
  return sh;
}
/** column numbers (0-based) of the Users sheet; adds the "Services" column when it is missing */
function usrCols_(sh, add) {
  var lc = Math.max(sh.getLastColumn(), 1), hdr = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
  var c = {
    teacher: usersCol_(hdr, 'teacher', 0), desig: usersCol_(hdr, 'designation', 2), dept: usersCol_(hdr, 'department', 3),
    user: usersCol_(hdr, 'username', 4), pass: usersCol_(hdr, 'password', 5), contact: usersCol_(hdr, 'contact', 6),
    role: usersCol_(hdr, 'role', 8), status: usersCol_(hdr, 'status', -1), svc: usersCol_(hdr, 'services', -1), n: lc, hdr: hdr,
    cp: usersCol_(hdr, 'chiefproctor', lc >= 10 ? 9 : -1),                 // v8.20 CP / CP1 / CP2 … (isCP_, getProctors_)
    tt: usersCol_(hdr, 'ttadmin', -1)                                      // v8.20 timetable editor role (Timetable.gs, at_mornOk_)
  };
  c.user = 4; c.pass = 5;                                                  // login reads E and F (authenticate_): always these two
  if (c.svc < 0 && add) {
    sh.getRange(1, lc + 1).setValue(USR_HEAD).setFontWeight('bold');
    c.svc = lc; c.n = lc + 1; c.hdr = hdr.concat([USR_HEAD]);
  }
  return c;
}
function usrClearCaches_(username) {
  ['users_att2', 'users_att3', 'users_full'].forEach(function (k) { try { cacheClear_(k); } catch (e) {} });
  var u = String(username || '').trim().toLowerCase();
  try { if (u) CacheService.getScriptCache().removeAll(['lvop_' + u, 'dutds_' + u]); } catch (e) {}   // leave in-charge, DS ADMIN
}
/* v8.20 facilities that the code gives by the Role, Chief Proctor and tt_admin columns (shown and set on users.html) */
var USR_ROLES = [
  { v: 'TEACHER',            about: 'Her own classes from the timetable (attendance, tests, her reports).' },
  { v: 'HOD',                about: 'Her classes + the department overview, department reports, shortage, marked-missed entries (needs the Department).' },
  { v: 'PRINCIPAL',          about: 'Whole-college overview, all reports, shortage, marked-missed entries, admissions, timetable morning report.' },
  { v: 'DS ADMIN',           about: 'Exam duty lists (and the whole-college view, as before).' },
  { v: 'LEAVE IN-CHARGE',    about: 'Approved leaves page.' },
  { v: 'COMPUTER ASSISTANT', about: 'Only the services ticked below.' }
];
var USR_TT = [
  { v: '',                     about: 'Can only view the timetable.' },
  { v: 'HOD',                  about: 'Can change teacher and days of her own department.' },
  { v: 'ARTS TT ADMIN',        about: 'Can change the F.A timetable.' },
  { v: 'SCI TT MORNING ADMIN', about: 'Can change morning science sections; also morning classes report, room tags and workload in Admin tools.' },
  { v: 'SCI TT EVENING ADMIN', about: 'Can change evening science sections.' },
  { v: 'ADMIN',                about: 'Can change the whole timetable and its settings.' }
];

/* ---------- list ---------- */
function usrList_() {
  var sh = usrSheet_(), c = usrCols_(sh, false), last = sh.getLastRow();
  var v = last > 1 ? sh.getRange(2, 1, last - 1, c.n).getDisplayValues() : [];
  var g = function (r, i) { return i >= 0 && i < r.length ? String(r[i] || '').trim() : ''; };
  var rows = [], roles = {}, depts = {}, desigs = {};
  v.forEach(function (r, i) {
    var u = g(r, c.user); if (!u && !g(r, c.teacher)) return;
    var svc = svcParse_(g(r, c.svc));
    var o = { row: i + 2, teacher: g(r, c.teacher), username: u, desig: g(r, c.desig), dept: g(r, c.dept), contact: g(r, c.contact),
              role: g(r, c.role), status: g(r, c.status), svc: svc.filter(function (k) { return k !== 'blocked'; }),
              blocked: svc.indexOf('blocked') >= 0, admin: u.toLowerCase() === ADMIN_USER, hasPass: !!g(r, c.pass),
              cp: g(r, c.cp), tt: g(r, c.tt) };
    if (o.role) roles[o.role] = 1; if (o.dept) depts[o.dept] = 1; if (o.desig) desigs[o.desig] = 1;
    rows.push(o);
  });
  return { ok: true, rows: rows, services: SVC_LIST.map(function (s) { return { key: s.key, label: s.label, about: s.about, page: s.page, subs: (SVC_SUBS[s.key] || []).map(function (x) { return { k: x.k, label: x.label }; }) }; }),
           roles: Object.keys(roles).sort(), depts: Object.keys(depts).sort(), desigs: Object.keys(desigs).sort(), roleList: USR_ROLES, ttList: USR_TT, screens: USR_SCREENS, roleHide: usrRoleHide_(),
           has: { desig: c.desig >= 0, dept: c.dept >= 0, contact: c.contact >= 0, role: c.role >= 0, status: c.status >= 0, svc: c.svc >= 0, cp: c.cp >= 0, tt: c.tt >= 0 } };
}

/* ---------- add / change one login ---------- */
function usrSave_(req) {
  var rec = req.rec || {}, row = Number(req.row) || 0, np = String(req.newPass || '');
  var t = String(rec.teacher || '').trim(), u = String(rec.username || '').trim();
  if (!t) return { ok: false, error: 'Type the name.' };
  if (!u || /\s/.test(u)) return { ok: false, error: 'Type a username without spaces.' };
  if (!row && np.length < 4) return { ok: false, error: 'A new login needs a password of at least 4 characters.' };
  if (np && np.length < 4) return { ok: false, error: 'The password must have at least 4 characters.' };
  var role0 = String(rec.role || '').trim().toUpperCase(), tt0 = String(rec.tt || '').trim().toUpperCase(), cp0 = String(rec.cp || '').trim().toUpperCase();
  if ((/\bHOD\b/.test(role0) || tt0 === 'HOD') && !String(rec.dept || '').trim()) return { ok: false, error: 'An HOD needs her Department (the department she heads).' };
  if (cp0 && !/^CP\d*$/.test(cp0)) return { ok: false, error: 'Chief Proctor must be CP, CP1, CP2 …' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = usrSheet_(), c = usrCols_(sh, true), last = sh.getLastRow();
    var users = last > 1 ? sh.getRange(2, c.user + 1, last - 1, 1).getDisplayValues().map(function (r) { return String(r[0]).trim().toLowerCase(); }) : [];
    var dup = users.indexOf(u.toLowerCase());
    if (dup >= 0 && dup + 2 !== row) return { ok: false, error: 'The username "' + u + '" is already used (row ' + (dup + 2) + ').' };
    if (row) {
      if (row < 2 || row > last) return { ok: false, error: 'That row is not in the Users sheet any more. Reload the page.' };
      var was = users[row - 2];
      if (was !== String(req.oldUser || '').trim().toLowerCase()) return { ok: false, error: 'The Users sheet was changed meanwhile (row ' + row + ' is now "' + was + '"). Reload the page.' };
      if (was === ADMIN_USER && u.toLowerCase() !== ADMIN_USER) return { ok: false, error: 'The main admin username cannot be changed here.' };
    }
    var r = row || last + 1, put = {};                                    // column -> new value (only these cells are written)
    var set = function (i, val) { if (i >= 0 && val !== undefined) put[i] = val; };
    set(c.teacher, t); set(c.user, u);
    if (np) set(c.pass, np);
    set(c.desig, rec.desig == null ? undefined : String(rec.desig).trim());
    set(c.dept, rec.dept == null ? undefined : String(rec.dept).trim().toUpperCase());
    set(c.contact, rec.contact == null ? undefined : String(rec.contact).trim());
    set(c.role, rec.role == null ? undefined : String(rec.role).trim().toUpperCase());
    if (c.status >= 0) set(c.status, rec.status == null ? undefined : String(rec.status).trim());
    if (c.cp >= 0 && rec.cp != null) set(c.cp, cp0);                     // v8.20 Chief Proctor
    if (c.tt >= 0 && rec.tt != null) set(c.tt, tt0);                     // v8.20 timetable editor role
    var svc = svcParse_((rec.svc || []).join(','));
    if (rec.blocked && u.toLowerCase() !== ADMIN_USER) svc.push('blocked');
    set(c.svc, svc.join(', '));
    Object.keys(put).forEach(function (k) {
      var i = Number(k), cell = sh.getRange(r, i + 1);
      if (i === c.pass || i === c.contact || i === c.user) cell.setNumberFormat('@');   // 0315… and 1234 stay text
      cell.setValue(put[k]);
    });
    SpreadsheetApp.flush();
    usrClearCaches_(u);
    return { ok: true, row: r, added: !row };
  } finally { lock.releaseLock(); }
}

/* ---------- switch services / the login on or off (the switches on the list) ---------- */
function usrSet_(req) {
  var row = Number(req.row) || 0;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = usrSheet_(), c = usrCols_(sh, true);
    if (row < 2 || row > sh.getLastRow()) return { ok: false, error: 'That row is not in the Users sheet any more. Reload the page.' };
    var u = String(sh.getRange(row, c.user + 1).getDisplayValue()).trim();
    if (u.toLowerCase() !== String(req.oldUser || '').trim().toLowerCase()) return { ok: false, error: 'The Users sheet was changed meanwhile. Reload the page.' };
    if (u.toLowerCase() === ADMIN_USER) return { ok: false, error: 'The main admin has every service already and cannot be switched off.' };
    var svc = svcParse_((req.svc || []).join(','));
    if (req.blocked === true) svc.push('blocked');
    sh.getRange(row, c.svc + 1).setValue(svc.join(', '));
    SpreadsheetApp.flush();
    usrClearCaches_(u);
    return { ok: true, svc: svc.filter(function (k) { return k !== 'blocked'; }), blocked: svc.indexOf('blocked') >= 0 };
  } finally { lock.releaseLock(); }
}
