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
const JUNK = /GAZETTE|BOARD OF INTERMEDIATE|B\.?I\.?S\.?E|RESULT|\bPAGE\b|ZAKRIA|0349|EXAMINATION|ANNUAL|SUPPLEMENTARY|CONTROLLER|SECRETARY|NOTIFICATION|PRINTED|DOWNLOAD|WWW\.|HTTP/i;
const COLLEGEISH = /COLLEGE|SCHOOL|CANDIDATES|INSTITUTE|ACADEMY|SYSTEM|CAMPUS|CENTRE|CENTER|FOUNDATION|MADRAS|SECONDARY|CADET|UNIVERSITY|EDUCATION|COMPLEX|H\.?\s?S\.?\s?S|PUBLIC|MODEL|GRAMMAR|DEGREE|SCIENCE|ISLAMIA|TRUST|GOVT|GOVERNMENT|\bGGHSS\b|\bGHSS\b/i;
const ROLL = /^\d{4,8}$/;
const norm = s => String(s || '').replace(/\s+/g, ' ').trim();
const isHeader = t => { const u = t.toUpperCase().replace(/\s+/g, ''); return /ROLL/.test(u) && /NAME/.test(u) && (/MARK/.test(u) || /REGIS/.test(u)); };

/* ---------- text items -> lines -> cells ---------- */
function linesOf(tc, split) {
  split = split || [];
  let it = tc.items.filter(i => i.str && i.str.trim()).map(i => {
    const h = Math.abs(i.transform[3]) || Math.abs(i.transform[0]) || i.height || 8;
    return { s: i.str, x: i.transform[4], y: i.transform[5], w: Math.abs(i.width) || i.str.length * h * 0.5, h, f: i.fontName || '' };
  });
  /* pdf.js joins text that sits close together on a line ("AREEBA KHAN SANA NAWAZ" = name + father):
     cut such a piece at the space nearest to each column edge inside it */
  if (split.length) it = [].concat(...it.map(function cut(i) {
    const e = split.find(e => e > i.x + 3 && e < i.x + i.w - 3); if (e == null || !/\s/.test(i.s.trim())) return [i];
    const per = i.w / i.s.length, want = (e - i.x) / per; let best = -1;
    for (let k = 1; k < i.s.length; k++) if (i.s[k] === ' ' && (best < 0 || Math.abs(k - want) < Math.abs(best - want))) best = k;
    if (best < 0 || Math.abs(best - want) > Math.max(6, i.s.length * 0.25)) return [i];
    const a = { s: i.s.slice(0, best), x: i.x, y: i.y, w: best * per, h: i.h, f: i.f };
    const b = { s: i.s.slice(best + 1), x: Math.max(i.x + (best + 1) * per, e + 0.5), y: i.y, w: (i.s.length - best - 1) * per, h: i.h, f: i.f };
    return [a].concat(cut(b));
  }));
  it.sort((a, b) => b.y - a.y || a.x - b.x);
  const L = [];
  it.forEach(i => {
    const l = L.find(z => Math.abs(z.y - i.y) <= Math.max(1.6, Math.min(z.h, i.h) * 0.35));
    if (l) { l.items.push(i); l.h = Math.max(l.h, i.h); } else L.push({ y: i.y, h: i.h, items: [i] });
  });
  L.forEach(l => {
    l.items.sort((a, b) => a.x - b.x);
    const cells = [];
    l.items.forEach(i => {
      const c = cells[cells.length - 1], gap = c ? i.x - (c.x1) : 1e9;
      if (c && gap < Math.max(2.2, i.h * 0.45) && !split.some(e => e > c.x0 + 1 && e <= i.x + 1)) { c.s += (gap > i.h * 0.12 && !/\s$/.test(c.s) && !/^\s/.test(i.s) ? ' ' : '') + i.s; c.x1 = Math.max(c.x1, i.x + i.w); c.f.push(i.f); }
      else cells.push({ s: i.s, x0: i.x, x1: i.x + i.w, f: [i.f] });
    });
    cells.forEach(c => { c.s = norm(c.s); });
    l.cells = cells.filter(c => c.s);
    l.text = l.cells.map(c => c.s).join(' ');
  });
  L.sort((a, b) => b.y - a.y);
  return L;
}

/* ---------- ruling lines (vertical) from the drawing operators ---------- */
async function verticals(page, OPS) {
  if (!OPS) return [];
  let ol; try { ol = await page.getOperatorList(); } catch (e) { return []; }
  let ctm = [1, 0, 0, 1, 0, 0]; const stack = [], V = [];
  const mul = (m, n) => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
  const tp = (x, y) => [ctm[0] * x + ctm[2] * y + ctm[4], ctm[1] * x + ctm[3] * y + ctm[5]];
  const seg = (a, b) => { if (Math.abs(a[0] - b[0]) < 1.2 && Math.abs(a[1] - b[1]) > 4) V.push({ x: (a[0] + b[0]) / 2, y0: Math.min(a[1], b[1]), y1: Math.max(a[1], b[1]) }); };
  for (let i = 0; i < ol.fnArray.length; i++) {
    const fn = ol.fnArray[i], a = ol.argsArray[i];
    if (fn === OPS.save) stack.push(ctm.slice());
    else if (fn === OPS.restore) ctm = stack.pop() || ctm;
    else if (fn === OPS.transform) ctm = mul(ctm, a);
    else if (fn === OPS.constructPath) {
      const ops = a[0], co = a[1]; let k = 0, cur = null;
      for (const op of ops) {
        if (op === OPS.rectangle) {
          const x = co[k], y = co[k + 1], w = co[k + 2], h = co[k + 3]; k += 4;
          const p = tp(x, y), q = tp(x + w, y + h);
          if (Math.abs(q[0] - p[0]) < 2.5) seg([(p[0] + q[0]) / 2, p[1]], [(p[0] + q[0]) / 2, q[1]]);       // a thin filled bar = a line
          else if (Math.abs(q[1] - p[1]) > 4) { seg([p[0], p[1]], [p[0], q[1]]); seg([q[0], p[1]], [q[0], q[1]]); }   // a cell box: its two sides
        } else if (op === OPS.moveTo) { cur = tp(co[k], co[k + 1]); k += 2; }
        else if (op === OPS.lineTo) { const n = tp(co[k], co[k + 1]); k += 2; if (cur) seg(cur, n); cur = n; }
        else if (op === OPS.curveTo) { cur = tp(co[k + 4], co[k + 5]); k += 6; }
        else if (op === OPS.curveTo2 || op === OPS.curveTo3) { cur = tp(co[k + 2], co[k + 3]); k += 4; }
      }
    }
  }
  return V;
}
function clusterX(V) {
  const xs = V.slice().sort((a, b) => a.x - b.x), out = [];
  xs.forEach(v => { const c = out[out.length - 1]; if (c && v.x - c.x < 2) { c.len += v.y1 - v.y0; c.x = (c.x + v.x) / 2; } else out.push({ x: v.x, len: v.y1 - v.y0 }); });
  const maxLen = Math.max(0, ...out.map(c => c.len));
  return out.filter(c => c.len > maxLen * 0.25).map(c => c.x);
}

/* ---------- the parser ---------- */
async function parse(pdf, opt) {
  opt = opt || {};
  const OPS = opt.OPS, prog = opt.onProgress || (() => {}), n = pdf.numPages;
  let split = [], head = null, edges = null, spans = null, R = null, college = '', pendingCollege = '', rows = [], warn = [], dataFont = '';
  /* which column a piece of text belongs to:
     ruling lines -> the column between two lines; no lines -> the header cell it overlaps, else the column whose
     earlier data (learned from rows that clearly overlapped a header) it falls in, else the nearest header */
  const ov = (c, a, b) => Math.max(0, Math.min(c.x1, b) - Math.max(c.x0, a));
  const lineCol = x => { for (let k = edges.length - 1; k >= 0; k--) if (x >= edges[k] - 1.5) return k; return 0; };
  const hits = c => spans.map((h, k) => ov(c, h[0], h[1]) > 0.5 ? k : -1).filter(k => k >= 0);
  function colOf(c) {
    if (edges) return lineCol(c.x0 + 1);
    const h = hits(c); if (h.length) return h.reduce((a, k) => ov(c, spans[k][0], spans[k][1]) > ov(c, spans[a][0], spans[a][1]) ? k : a, h[0]);
    const m = (c.x0 + c.x1) / 2;
    let k = R.findIndex(r => r && m >= r[0] - 3 && m <= r[1] + 3); if (k >= 0) return k;
    let best = 0, bd = 1e9;
    spans.forEach((sp, j) => { const r = R[j], a = r ? Math.min(r[0], sp[0]) : sp[0], b = r ? Math.max(r[1], sp[1]) : sp[1], d = m < a ? a - m : m > b ? m - b : 0; if (d < bd) { bd = d; best = j; } });
    return best;
  }
  const crosses = c => edges ? lineCol(c.x0 + 1) !== lineCol(c.x1 - 1.5) : hits(c).length > 1 || R.filter(r => r && ov(c, r[0], r[1]) > 2).length > 1;
  function learn(L) {                                   // no lines: remember where each column's data really is
    if (edges) return;
    L.forEach(l => { if (!l.cells.length || !ROLL.test(l.cells[0].s.split(' ')[0])) return;
      l.cells.forEach(c => { const h = hits(c); if (h.length !== 1) return; const k = h[0]; R[k] = R[k] ? [Math.min(R[k][0], c.x0), Math.max(R[k][1], c.x1)] : [c.x0, c.x1]; }); });
  }
  for (let p = Math.max(1, opt.from || 1); p <= n; p++) {
    if (opt.cancel && opt.cancel()) throw new Error('Stopped.');
    const page = await pdf.getPage(p), tc = await page.getTextContent({ disableCombineTextItems: true });
    let L = linesOf(tc, split);
    let hdrIdx = L.findIndex(l => isHeader(l.text));
    if (!head && hdrIdx >= 0) {
      /* the header: its cells (a second header line just above / below is joined to the nearest cell) */
      const hl = L[hdrIdx], cells = hl.cells.map(c => ({ s: c.s, x0: c.x0, x1: c.x1 }));
      [L[hdrIdx - 1], L[hdrIdx + 1]].forEach(o => {
        if (!o || Math.abs(o.y - hl.y) > hl.h * 1.8 || /\d{3}/.test(o.text) || isHeader(o.text) || o.cells.length > cells.length) return;
        o.cells.forEach(c => { const m = (c.x0 + c.x1) / 2; let best = cells[0]; cells.forEach(h => { if (Math.abs((h.x0 + h.x1) / 2 - m) < Math.abs((best.x0 + best.x1) / 2 - m)) best = h; }); best.s = o.y > hl.y ? c.s + ' ' + best.s : best.s + ' ' + c.s; });
      });
      head = cells.map(c => norm(c.s).replace(/^ROLL\s*NO\.?$/i, 'Roll No').replace(/^RollNo$/i, 'Roll No'));
      /* column edges: ruling lines if the table has them, else half-way between the header cells */
      let xs = clusterX(await verticals(page, OPS)), ok = false;
      if (xs.length >= cells.length) {
        const iv = []; for (let k = 0; k + 1 < xs.length; k++) iv.push([xs[k], xs[k + 1]]);
        const owner = cells.map(c => { const m = (c.x0 + c.x1) / 2; return iv.findIndex(v => m >= v[0] && m < v[1]); });
        if (owner.every(o => o >= 0) && new Set(owner).size === owner.length) { edges = owner.map(o => iv[o][0]); ok = true; }
      }
      if (!ok){ edges = null; spans = cells.map(c => [c.x0, c.x1]); R = cells.map(() => null); }
      head.lines = ok;
      if (ok) { split = edges.slice(1); L = linesOf(tc, split); hdrIdx = L.findIndex(l => isHeader(l.text)); }
    }
    if (!head) {                                       // before the first table (cover pages): remember a college line
      L.forEach(l => { if (COLLEGEISH.test(l.text) && !JUNK.test(l.text) && l.cells.length <= 2) pendingCollege = norm(l.text); });
      prog(p, n, rows.length); continue;
    }
    if (!college && pendingCollege) college = pendingCollege;
    /* rows of this page */
    const recs = [], loose = [];
    learn(L);
    if (!edges && !split.length && R.every(Boolean)) {                  // learned every column: never join text across columns again
      split = R.slice(1).map((r, k) => (R[k][1] + r[0]) / 2);
      if (split.every((e, k) => !k || e > split[k - 1])) L = linesOf(tc, split); else split = [];
    }
    const hdrY = L.filter(l => isHeader(l.text)).map(l => l);
    L.forEach((l, li) => {
      if (isHeader(l.text)) return;
      if (hdrY.some(h => Math.abs(h.y - l.y) <= h.h * 1.8) && !/\d{3}/.test(l.text) && !/COLLEGE|SCHOOL|CANDIDATE/i.test(l.text)) return;   // 2nd line of a header
      const first = l.cells[0], c0 = first && colOf(first) === 0 ? first.s : '';
      const rollTok = c0.split(' ')[0];
      if (ROLL.test(rollTok)) {
        const cells = head.map(() => []);
        if (c0.length > rollTok.length && head.length > 1) cells[1].push(c0.slice(rollTok.length).trim());   // roll no and name in one piece of text
        cells[0].push(rollTok);
        l.cells.slice(1).forEach(c => cells[colOf(c)].push(c.s));
        if (!dataFont) dataFont = (first.f || [])[0] || '';
        recs.push({ y: l.y, h: l.h, cells, college, li });
        return;
      }
      if (/^\d{1,4}$/.test(l.text) || (JUNK.test(l.text) && !/COLLEGE|SCHOOL|CANDIDATE|INSTITUTE|ACADEMY/i.test(l.text))) return;          // page numbers, titles
      const bold = l.cells.every(c => (c.f || []).some(f => /bold|black|heavy/i.test(f))) && dataFont && !/bold/i.test(dataFont);
      const isCollege = l.cells.length <= 3 && /[A-Z]{3}/i.test(l.text) && (l.cells.some(crosses) || colOf(l.cells[0]) === 0 || (COLLEGEISH.test(l.text) && l.cells.length === 1 && l.text === l.text.toUpperCase()) || bold);
      if (isCollege) { college = norm(l.text); recs.push({ y: l.y, college: college, mark: true }); return; }
      loose.push(l);
    });
    /* wrapped text (a long name on two lines): to the nearest student row of the same page, never across a college line */
    loose.forEach(l => {
      let best = null, bd = 1e9;
      recs.forEach(r => { if (r.mark) return; const d = Math.abs(r.y - l.y); if (d < bd) { bd = d; best = r; } });
      if (!best || bd > Math.max(best.h, l.h) * 3.2) return;
      const between = recs.some(r => r.mark && (r.y - l.y) * (r.y - best.y) < 0);
      if (between) return;
      l.cells.forEach(c => { const k = colOf(c); (l.y > best.y ? best.cells[k].unshift(c.s) : best.cells[k].push(c.s)); });
    });
    recs.forEach(r => { if (!r.mark) rows.push(r.cells.map(a => norm(a.join(' '))).concat([r.college])); });
    prog(p, n, rows.length);
    if (opt.limit && rows.length >= opt.limit) break;
  }
  if (!head) throw new Error('No table with a "Roll No … Name … Marks" header was found in this PDF.');
  if (!head.lines) warn.push('The table has no ruling lines: columns were taken from the header. Check a few rows.');
  return { header: head.slice(), rows, warn, lines: !!head.lines };
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
