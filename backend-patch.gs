/**
 * BACKEND PATCH v5 for the CS Attendance Apps Script
 *
 * Do these 3 things in your existing script:
 *   1) REPLACE the whole doPost(e) function with the one below.
 *   2) REPLACE the whole changePassword_(user, oldPass, newPass) function with the one below.
 *   3) PASTE everything from "TOKENS" to the end at the bottom of the script.
 * (Optional) change doGet to say 'CS Attendance backend v5'.
 *
 * Then: Deploy > Manage deployments > Edit > Version: New version > Deploy.
 *
 * Login tokens last 120 days and stop working the moment that user's password is changed
 * in the Users sheet. So if a phone is lost, change that teacher's password and the phone is locked out.
 * Old app versions that still send the password keep working.
 */

/* ===== 1) REPLACE doPost ===== */
function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) {
    return json_({ ok: false, error: 'BAD_REQUEST' });
  }

  var user = null;
  if (req.action === 'login')  user = authenticate_(req.username, req.password);
  else if (req.token)          user = verifyToken_(req.username, req.token);
  else                         user = authenticate_(req.username, req.password);
  if (!user) return json_({ ok: false, error: 'BAD_LOGIN' });

  try {
    if (req.action === 'login') {
      return json_({ ok: true, token: makeToken_(user), teacher: user.teacher, admin: isAdmin_(user), sheets: listSheets_(user) });
    }
    if (req.action === 'sheets') {
      return json_({ ok: true, teacher: user.teacher, admin: isAdmin_(user), sheets: listSheets_(user) });
    }
    if (req.action === 'sheet')    return json_(getSheetData_(user, req.sheet, req.create === true));
    if (req.action === 'save')     return json_(saveRecords_(user, req.records || []));
    if (req.action === 'pdf')      return json_(exportPdf_(user, req.sheet));
    if (req.action === 'password') return json_(changePassword_(user, req.oldPass, req.newPass));
    if (req.action === 'overview') return json_(getOverview_(user, req.day));
    return json_({ ok: false, error: 'UNKNOWN_ACTION' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ===== 2) REPLACE changePassword_ ===== */
function changePassword_(user, oldPass, newPass) {
  if (String(oldPass) !== user.password) return { ok: false, error: 'WRONG_OLD' };
  if (!newPass || String(newPass).length < 4) return { ok: false, error: 'TOO_SHORT' };
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users');
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][4]).trim().toLowerCase() === user.username.toLowerCase()) {
      sh.getRange(i + 1, 6).setValue(String(newPass));
      cacheClear_('users_att');
      // all older tokens of this user stop working; hand back a fresh one for this phone
      return { ok: true, token: makeToken_({ username: user.username, password: String(newPass) }) };
    }
  }
  return { ok: false, error: 'User not found' };
}

/* ===== 3) PASTE AT THE BOTTOM: TOKENS ===== */

var TOKEN_DAYS = 120;

function tokenSecret_() {
  var p = PropertiesService.getScriptProperties();
  var s = p.getProperty('TOKEN_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('TOKEN_SECRET', s); }
  return s;
}

function tokenSig_(msg) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(msg, tokenSecret_()));
}

/** token = <expiry ms>.<signature of username|expiry|password> */
function makeToken_(user) {
  var exp = Date.now() + TOKEN_DAYS * 86400000;
  var u = String(user.username).trim().toLowerCase();
  return exp + '.' + tokenSig_(u + '|' + exp + '|' + user.password);
}

function verifyToken_(username, token) {
  var u = String(username || '').trim().toLowerCase();
  var parts = String(token || '').split('.');
  if (!u || parts.length !== 2) return null;
  var exp = Number(parts[0]);
  if (!exp || exp < Date.now()) return null;
  var rows = getUsers_();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][1].toLowerCase() === u) {
      if (tokenSig_(u + '|' + exp + '|' + rows[i][2]) === parts[1]) {
        return { teacher: rows[i][0], username: rows[i][1], password: rows[i][2] };
      }
      return null;
    }
  }
  return null;
}

/* ===== ADMIN: which sections have attendance for a day (default today) ===== */
function getOverview_(user, dayIn) {
  if (!isAdmin_(user)) return { ok: false, error: 'ADMIN_ONLY' };
  var tok = currentSeasonToken_();
  if (!tok) return { ok: true, day: 0, month: '', items: [] };
  var day = Number(dayIn) || Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd'));
  var book = SpreadsheetApp.openById(BOOK_ID);
  var byName = {};
  book.getSheets().forEach(function (s) { byName[s.getName()] = s; });

  var items = [];
  getAssignments_(user).forEach(function (a) {
    var name = makeName_(a.cls, tok, a.subject), sh = byName[name];
    var it = { cls: a.cls, sub: a.subject, teacher: a.teacher, status: 'nosheet', marked: 0, total: 0 };
    if (sh) {
      try {
        var L = readLayout_(sh), col = L.dayCols[day], n = sh.getLastRow() - FIRST_ROW + 1;
        if (col && n > 0) {
          var v = sh.getRange(FIRST_ROW, col, n, 1).getValues();
          var rl = sh.getRange(FIRST_ROW, L.rollCol, n, 1).getValues(), m = 0, t = 0;
          for (var i = 0; i < n; i++) {
            if (!String(rl[i][0]).trim()) continue;
            t++;
            if (VALID[String(v[i][0]).trim().toUpperCase()]) m++;
          }
          it.total = t; it.marked = m;
          it.status = m === 0 ? 'none' : (m >= t ? 'done' : 'partial');
        } else it.status = 'none';
      } catch (err) { it.status = 'error'; }
    }
    items.push(it);
  });
  return { ok: true, day: day, month: tok, items: items };
}
