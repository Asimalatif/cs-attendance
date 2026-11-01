/* boardparse.js — reads a Peshawar Board result gazette PDF (pdf.js) into rows, like the Colab notebook
   "HSSC Part I-II Gazette to excel":
     1. every table row of every page (Roll No, Name, Father Name, Registration #, Marks, Remarks … as printed)
     2. the college name printed above each block of students goes into "College Name"
     3. Marks as a number; Position = dense rank of Marks (highest = 1), PRIVATE CANDIDATES left out
     4. sorted by Marks, highest first; and the Jinnah College top 20
   Column edges come from the table's ruling lines; without lines, from the header and the data.
   Works in the browser (window.BoardParse) and in Node (module.exports) for testing. */
(function (root) {
'use strict';
const JUNK = /GAZETTE|BOARD OF INTERMEDIATE|RESULT|\bPAGE\b|03491055220|EXAMINATION|ANNUAL|SUPPLEMENTARY|CONTROLLER|SECRETARY|NOTIFICATION|PRINTED|DOWNLOAD|WWW\.|HTTP/i;
const COLLEGEISH = /COLLEGE|SCHOOL|CANDIDATES|INSTITUTE|ACADEMY|SYSTEM|CAMPUS|CENTRE|CENTER|FOUNDATION|MADRAS|SECONDARY|CADET|UNIVERSITY|EDUCATION|COMPLEX|H\.?\s?S\.?\s?S|PUBLIC|MODEL|GRAMMAR|DEGREE|SCIENCE|ISLAMIA|TRUST|GOVT|GOVERNMENT|\bGGHSS\b|\bGHSS\b/i;
const ROLL = /^\d{4,8}$/;
const norm = s => String(s || '').replace(/\s+/g, ' ').trim();
const isHeader = t => { const u = t.toUpperCase().replace(/\s+/g, ''); return /ROLL/.test(u) && /NAME/.test(u) && (/MARK/.test(u) || /REGIS/.test(u)); };

/* ---------- text items -> lines -> cells ---------- */
function linesOf(tc) {
  const it = tc.items.filter(i => i.str && i.str.trim()).map(i => {
    const h = Math.abs(i.transform[3]) || Math.abs(i.transform[0]) || i.height || 8;
    return { s: i.str, x: i.transform[4], y: i.transform[5], w: Math.abs(i.width) || i.str.length * h * 0.5, h, f: i.fontName || '' };
  });
  it.sort((a, b) => b.y - a.y || a.x - b.x);
  const L = [];
  it.forEach(i => {
    const l = L.find(z => Math.abs(z.y - i.y) <= Math.max(1.6, Math.min(z.h, i.h) * 0.35));
    if (l) { l.items.push(i); l.h = Math.max(l.h, i.h); } else L.push({ y: i.y, h: i.h, items: [i] });
  });
  L.forEach(l => { l.items.sort((a, b) => a.x - b.x); l.cells = cellsFrom(l.items, []); l.text = l.cells.map(c => c.s).join(' '); });
  L.sort((a, b) => b.y - a.y);
  return L;
}
/** text pieces of one line -> cells; never across a column edge (split = x of the edges).
 *  pdf.js joins text that sits close together ("AREEBA KHAN SANA NAWAZ" = name + father): such a piece is cut
 *  at the space nearest to the edge inside it */
function cellsFrom(items, split) {
  let it = items;
  if (split.length) it = [].concat(...items.map(function cut(i) {
    const e = split.find(e => e > i.x + 3 && e < i.x + i.w - 3); if (e == null || !/\s/.test(i.s.trim())) return [i];
    const per = i.w / i.s.length, want = (e - i.x) / per; let best = -1;
    for (let k = 1; k < i.s.length; k++) if (i.s[k] === ' ' && (best < 0 || Math.abs(k - want) < Math.abs(best - want))) best = k;
    if (best < 0 || Math.abs(best - want) > Math.max(6, i.s.length * 0.25)) return [i];
    const a = { s: i.s.slice(0, best), x: i.x, y: i.y, w: best * per, h: i.h, f: i.f };
    const b = { s: i.s.slice(best + 1), x: Math.max(i.x + (best + 1) * per, e + 0.5), y: i.y, w: (i.s.length - best - 1) * per, h: i.h, f: i.f };
    return [a].concat(cut(b));
  }));
  const cells = [];
  it.forEach(i => {
    const c = cells[cells.length - 1], gap = c ? i.x - c.x1 : 1e9;
    if (c && gap < Math.max(2.2, i.h * 0.45) && !split.some(e => e > c.x0 + 1 && e <= i.x + 1)) { c.s += (gap > i.h * 0.12 && !/\s$/.test(c.s) && !/^\s/.test(i.s) ? ' ' : '') + i.s; c.x1 = Math.max(c.x1, i.x + i.w); }
    else cells.push({ s: i.s, x0: i.x, x1: i.x + i.w });
  });
  cells.forEach(c => { c.s = norm(c.s); });
  return cells.filter(c => c.s);
}

/* ---------- the parser ---------- */
async function parse(pdf, opt) {
  opt = opt || {};
  const prog = opt.onProgress || (() => {}), n = pdf.numPages;
  /* Like the gazette itself: every college block is "COLLEGE NAME" / header row (RollNo, Name, Father Name,
     Registration #, Marks, Grade, Remarks) / student rows. The columns of each block are taken from ITS header
     (they shift a little from college to college). Pages before the first such header (cover, messages,
     result at a glance, position holders) are skipped, like START_PAGE in the Colab notebook. */
  const isTable = t => isHeader(t) && /REGIS/i.test(t.replace(/\s+/g, ''));
  let head = null, edges = null, college = '', rows = [], warn = [];
  const colOf = x => { for (let k = edges.length - 1; k >= 1; k--) if (x >= edges[k]) return k; return 0; };
  const cellsOf = l => cellsFrom(l.items, edges ? edges.slice(1) : []);
  for (let p = Math.max(1, opt.from || 1); p <= n; p++) {
    if (opt.cancel && opt.cancel()) throw new Error('Stopped.');
    const page = await pdf.getPage(p), L = linesOf(await page.getTextContent());
    const recs = [], loose = [];
    L.forEach((l, li) => {
      if (isTable(l.text)) {
        const hc = cellsFrom(l.items, []);
        if (!head) head = hc.map(c => norm(c.s).replace(/^ROLL\s*NO\.?$/i, 'Roll No'));
        if (hc.length === head.length) edges = hc.map((c, k) => k ? c.x0 - 3 : -1e9);
        else warn.push('Page ' + p + ': a header with ' + hc.length + ' columns instead of ' + head.length + '.');
        /* the line just above a header is the college name */
        const prev = L[li - 1];
        if (prev && !isTable(prev.text) && !/GAZETTE|BOARD OF INTERMEDIATE/i.test(prev.text) && /[A-Z]{3}/i.test(prev.text) && !ROLL.test(prev.text.split(' ')[0])) { college = norm(prev.text); recs.push({ y: prev.y, mark: true }); }
        recs.push({ y: l.y, mark: true });
        return;
      }
      if (!edges) return;
      if (/^\d{1,4}$/.test(l.text)) return;
      const cells = cellsOf(l), first = cells[0], tok = first ? first.s.split(' ')[0] : '';
      if (first && colOf(first.x0 + 1) === 0 && ROLL.test(tok)) {
        const cs = head.map(() => []);
        if (first.s.length > tok.length) cs[1].push(first.s.slice(tok.length).trim());
        cs[0].push(tok);
        cells.slice(1).forEach(c => cs[colOf(c.x0 + 1)].push(c.s));
        recs.push({ y: l.y, h: l.h, cells: cs, college });
        return;
      }
      if (JUNK.test(l.text)) return;                                                      // page titles
      if (first && colOf(first.x0 + 1) === 0 && /[A-Z]{3}/i.test(l.text)) {          // a college name in the first column
        if (!recs.some(r => r.mark && Math.abs(r.y - l.y) < 1)) { college = norm(l.text); recs.push({ y: l.y, mark: true }); }
        return;
      }
      loose.push({ y: l.y, h: l.h, cells });
    });
    /* a long name / father name on two lines: to the nearest student row, never across a header or college line */
    loose.forEach(l => {
      let best = null, bd = 1e9;
      recs.forEach(r => { if (r.mark) return; const d = Math.abs(r.y - l.y); if (d < bd) { bd = d; best = r; } });
      if (!best || bd > Math.max(best.h, l.h) * 3.2) return;
      if (recs.some(r => r.mark && (r.y - l.y) * (r.y - best.y) < 0)) return;
      l.cells.forEach(c => { const k = colOf(c.x0 + 1); (l.y > best.y ? best.cells[k].unshift(c.s) : best.cells[k].push(c.s)); });
    });
    recs.forEach(r => { if (!r.mark) rows.push(r.cells.map(a => norm(a.join(' '))).concat([r.college])); });
    prog(p, n, rows.length);
    if (opt.limit && rows.length >= opt.limit) break;
  }
  if (!head) throw new Error('No student table (RollNo, Name, Father Name, Registration #, Marks …) was found in this PDF.');
  return { header: head.slice(), rows, warn: [...new Set(warn)].slice(0, 5), lines: false };
}

/* ---------- Colab steps 3 and 4: Marks, Position, sort, Jinnah top 20 ---------- */
function finish(res, o) {
  o = o || {};
  const H = res.header.slice(), mi = H.findIndex(h => /^MARKS?\b/i.test(h.trim())), ci = H.length;
  if (mi < 0) throw new Error('No "Marks" column in the header: ' + H.join(' | '));
  const header = H.concat(['College Name', 'Position']);
  const seen = {}, rows = [];
  res.rows.forEach(r => {
    const k = r.join('|'); if (seen[k]) return; seen[k] = 1;             // the same row twice (page overlap)
    const out = r.slice(0, ci).concat(new Array(Math.max(0, ci - r.length)).fill('')).slice(0, ci);
    out.push(r[r.length - 1] || ''); out.push('');
    const m = String(out[mi]).replace(/,/g, '').trim(); out[mi] = /^\d+(\.\d+)?$/.test(m) ? Number(m) : m;
    rows.push(out);
  });
  const isNum = v => typeof v === 'number';
  const priv = r => /^PRIVATE CANDIDATES?$/i.test(String(r[ci]).trim());
  const marks = [...new Set(rows.filter(r => !priv(r) && isNum(r[mi])).map(r => r[mi]))].sort((a, b) => b - a), pos = {};
  marks.forEach((m, i) => { pos[m] = i + 1; });
  rows.forEach(r => { r[ci + 1] = !priv(r) && isNum(r[mi]) ? pos[r[mi]] : ''; });
  rows.forEach((r, i) => { r._i = i; });
  rows.sort((a, b) => (isNum(b[mi]) ? b[mi] : -1) - (isNum(a[mi]) ? a[mi] : -1) || a._i - b._i);
  rows.forEach(r => { delete r._i; });
  const tc = s => String(s || '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
  const ri = H.findIndex(h => /ROLL/i.test(h)), ni = H.findIndex(h => /^NAME$/i.test(h.trim())), fi = H.findIndex(h => /FATHER/i.test(h));
  const jcw = new RegExp(o.college || 'JINNAH COLLEGE FOR WOMEN', 'i');
  const top = rows.filter(r => jcw.test(r[ci]) && r[ci + 1] !== '' && r[ci + 1] <= (o.top || 20)).sort((a, b) => a[ci + 1] - b[ci + 1])
    .map((r, i) => [i + 1, r[ri], tc(r[ni]), tc(r[fi]), r[mi], r[ci + 1]]);
  return { header, rows, top: { header: ['S.No.', 'Roll No', 'Name', 'Father Name', 'Marks', 'Position'], rows: top },
    stats: { total: rows.length, ranked: rows.filter(r => r[ci + 1] !== '').length, priv: rows.filter(priv).length, noMarks: rows.filter(r => !isNum(r[mi])).length,
      colleges: new Set(rows.map(r => r[ci])).size, noCollege: rows.filter(r => !r[ci]).length } };
}

const api = { parse, finish, linesOf, isHeader };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.BoardParse = api;
})(typeof self !== 'undefined' ? self : this);
