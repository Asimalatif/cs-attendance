/* =====================================================================================================
   LEAVE LINK  (add this file to the TEACHERS' ATTENDANCE project, next to Code.gs)
   Reads the Leave In-Charge's "ApprovedLeavesStore" sheets so the attendance app can show
   students on approved leave as L.

   Only leaves entered WITH DATES can be matched to a day (column K, e.g.
   "10-Oct-2026 to 12-Oct-2026, 15-Oct-2026"). A leave entered only as a number has no days.
   Leaves are cached for 5 minutes, so a new leave reaches the teachers within 5 minutes.
   ===================================================================================================== */

// The workbooks that hold "ApprovedLeavesStore" (the shortage workbooks, one per month)
var LL_STORES = {
  SEPT: '1Gxjgcb82GZqFtNSMbK2QpVmaK2uppFFKtel6T93L-gk',
  OCT:  '1rs_jrVRislaz9LUMIRwGLdwDv24FTvFXS0L5GKi2IzM',
  NOV:  '1KtN4MBqfapaxqT1WafZJRLOzovFfMrEE00Jq7hlSD8A',
  DEC:  '1Y1bk_kmh4--jqOPNcTkdDd8X5a4iLAFi-EcQgbgis-4',
  JAN:  '16nZoP2pGAvtFhADo2sNqqvrS-0OEY4O-9CgBrrqULpg'
};
var LL_SHEET = 'ApprovedLeavesStore';
var LL_MON = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];

/** One date in any common form -> Date (or null):
 *  a real date cell, "14-Oct-2026", "14 Oct 2026", "14/10/2026", "2026-10-14", "14 OCT" (year taken from the month). */
function llParse_(s, my) {
  if (s instanceof Date && !isNaN(s)) return new Date(s.getFullYear(), s.getMonth(), s.getDate());
  s = String(s || '').trim();
  var m;
  if ((m = /^(\d{1,2})[-\s\/.]+([A-Za-z]{3})[A-Za-z]*[-\s\/.,]*(\d{4})?$/.exec(s))) {
    var mi = LL_MON.indexOf(m[2].toLowerCase()); if (mi < 0) return null;
    var y = m[3] ? Number(m[3]) : (my ? (mi === my.m ? my.y : (mi >= 7 ? (my.m >= 7 ? my.y : my.y - 1) : (my.m >= 7 ? my.y + 1 : my.y))) : new Date().getFullYear());
    return new Date(y, mi, Number(m[1]));
  }
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s))) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if ((m = /^(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})$/.exec(s))) {          // day/month/year (Pakistan)
    var yy = Number(m[3]); if (yy < 100) yy += 2000;
    return new Date(yy, Number(m[2]) - 1, Number(m[1]));
  }
  return null;
}

/** "10-Oct-2026 to 12-Oct-2026, 15-Oct-2026" (or a date cell) -> [Date, ...] */
function llDays_(cell, my) {
  var out = [];
  if (cell instanceof Date) { var d0 = llParse_(cell); if (d0) out.push(d0); return out; }
  String(cell || '').split(/[,;\n]+/).forEach(function (part) {
    part = part.trim(); if (!part) return;
    var bits = part.split(/\s+(?:to|till|until)\s+|\s+[-\u2013\u2014]\s+/i);
    var a = llParse_(bits[0], my), b = bits[1] ? llParse_(bits[1], my) : a;
    if (!a || !b || b < a || (b - a) > 62 * 86400000) return;
    for (var d = new Date(a); d <= b; d.setDate(d.getDate() + 1)) out.push(new Date(d));
  });
  return out;
}

/**
 * Approved leaves of one month token (SEPT ... JAN):  { day: { enroll: 'Leave type' } }
 * Only days that belong to that month are kept.
 */
function llMonth_(tok) {
  var key = 'll2_' + tok, hit = cacheGetBig_(key);
  if (hit) return hit;
  var out = {}, id = LL_STORES[tok];
  if (!id) return out;
  var my = monthYear_(tok);
  var sh = SpreadsheetApp.openById(id).getSheetByName(LL_SHEET);
  if (!sh) throw new Error('Sheet "' + LL_SHEET + '" not found in the ' + tok + ' leave workbook.');
  if (sh.getLastRow() > 1) {
    var n = sh.getLastRow() - 1, width = Math.min(11, sh.getMaxColumns());
    var disp = sh.getRange(2, 1, n, width).getDisplayValues();
    var raw = width >= 11 ? sh.getRange(2, 11, n, 1).getValues() : [];
    disp.forEach(function (r, i) {
      var enroll = String(r[1] || '').trim().toUpperCase();
      var k = raw[i] && raw[i][0] instanceof Date ? raw[i][0] : String(r[10] || '').trim();
      if (!enroll) return;
      var lines = String(r[9] || '').split(/\r?\n/), type = lines[0].trim() || 'Leave';
      var days = k ? llDays_(k, my) : [];
      if (!days.length) {                                    // leave entered as a number only: no days known
        (out._u = out._u || {})[enroll] = [String(r[5] || '').trim(), type];
        return;
      }
      days.forEach(function (d) {
        if (d.getMonth() !== my.m || d.getFullYear() !== my.y) return;
        var day = String(d.getDate());
        (out[day] = out[day] || {})[enroll] = type;
      });
    });
  }
  cachePutBig_(key, out);
  return out;
}

/** Approved leaves for the students of one attendance sheet: { day: { roll: 'Leave type' } } */
function leavesForSheet_(meta, students) {
  var all = llMonth_(meta.mon), mine = {}, out = {};
  (students || []).forEach(function (s) { mine[String(s.roll).trim().toUpperCase()] = s.roll; });
  Object.keys(all).forEach(function (day) {
    Object.keys(all[day]).forEach(function (roll) {
      if (mine[roll]) (out[day] = out[day] || {})[mine[roll]] = all[day][roll];   // day '_u' = leaves without dates: [count, type]
    });
  });
  return out;
}

/** Run by hand to check: lists this month's approved leaves found (see Execution log). */
function testLeaveLink() {
  var tok = currentSeasonToken_() || 'OCT', m = llMonth_(tok), days = Object.keys(m).filter(function (d) { return d !== '_u'; });
  Logger.log(tok + ': ' + days.length + ' day(s) with approved leave; ' + Object.keys(m._u || {}).length + ' leave(s) without dates');
  days.slice(0, 5).forEach(function (d) { Logger.log('Day ' + d + ': ' + JSON.stringify(m[d])); });
}

/**
 * Run by hand to find out why one student's leave does not show.
 * Put her enrollment number below, run it, then open the Execution log.
 */
function debugLeaveFor() {
  var ROLL = 'JCW-F26-2322';                 // <-- change this
  clearLeaveCache();
  SEASON.forEach(function (tok) {
    var id = LL_STORES[tok], sh;
    try { sh = SpreadsheetApp.openById(id).getSheetByName(LL_SHEET); }
    catch (e) { Logger.log(tok + ': CANNOT OPEN the leave workbook (' + e.message + '). The owner of this project needs access.'); return; }
    if (!sh) { Logger.log(tok + ': no "' + LL_SHEET + '" sheet'); return; }
    var n = sh.getLastRow() - 1, found = false;
    if (n > 0) sh.getRange(2, 1, n, Math.min(11, sh.getMaxColumns())).getDisplayValues().forEach(function (r) {
      if (String(r[1]).trim().toUpperCase() !== ROLL.toUpperCase()) return;
      found = true;
      var days = llDays_(String(r[10] || '').trim(), monthYear_(tok)).map(function (d) { return d.getDate() + '/' + (d.getMonth() + 1); });
      Logger.log(tok + ': type="' + String(r[9]).split('\n')[0] + '", count=' + r[5] + ', dates (column K)="' + r[10] + '" -> days ' + (days.length ? days.join(' ') : 'NONE: no dates, so it is shown only as a tag "' + r[5] + ' days" on her row, not as L'));
    });
    if (!found) Logger.log(tok + ': ' + ROLL + ' has no leave in this month (check the enrollment number, e.g. F25 vs F26)');
    var m = llMonth_(tok), cnt = 0; Object.keys(m).forEach(function (d) { if (d !== '_u' && m[d][ROLL.toUpperCase()]) cnt++; });
    Logger.log(tok + ': days the app will show as L for ' + ROLL + ': ' + cnt);
  });
}

/** Run by hand after the Leave In-Charge fixes old entries, to see them at once (else within 5 minutes). */
function clearLeaveCache() {
  SEASON.forEach(function (t) { cacheClear_('ll2_' + t); });
}
