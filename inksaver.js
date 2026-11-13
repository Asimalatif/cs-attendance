/* JCW ink saver: every PDF made in the app (jsPDF) and every printed page come out black & white with almost no toner.
 * - text: black
 * - filled areas: dark fills (bands, table headers) -> very light grey; light tints -> white
 * - lines: grey / black
 * - pictures (logo, photos) are not changed
 * Load this file BEFORE jspdf (plain <script src="inksaver.js"></script> in <head>).
 * ID cards (idcards.html, idcheck.html) keep their colours: they are printed by the card shop.
 * Colour PDFs again (this browser only): localStorage.setItem('jcw_color_pdf', '1')
 */
(function () {
  'use strict';
  if (/idcards|idcheck/i.test(location.pathname)) return;
  var on = function () { try { return localStorage.getItem('jcw_color_pdf') !== '1'; } catch (e) { return true; } };

  /* any jsPDF colour argument -> [r, g, b] (0-255) or null */
  function rgb(a) {
    if (!a.length) return null;
    var x = a[0];
    if (typeof x === 'string') {
      var s = x.trim();
      if (/^#?[0-9a-f]{6}$/i.test(s)) { s = s.replace('#', ''); return [parseInt(s.slice(0, 2), 16), parseInt(s.slice(2, 4), 16), parseInt(s.slice(4, 6), 16)]; }
      if (/^#?[0-9a-f]{3}$/i.test(s)) { s = s.replace('#', ''); return [parseInt(s[0] + s[0], 16), parseInt(s[1] + s[1], 16), parseInt(s[2] + s[2], 16)]; }
      if (/^\d+(\.\d+)?$/.test(s)) x = Number(s); else return null;
    }
    if (typeof x !== 'number') return null;
    if (a.length >= 3 && typeof a[1] === 'number' && typeof a[2] === 'number') {
      var r = x, g = a[1], b = a[2];
      if (a.length >= 4 && typeof a[3] === 'number') {                // CMYK 0-1
        var k = a[3]; return [255 * (1 - r) * (1 - k), 255 * (1 - g) * (1 - k), 255 * (1 - b) * (1 - k)];
      }
      if (r <= 1 && g <= 1 && b <= 1 && (r % 1 || g % 1 || b % 1)) { r *= 255; g *= 255; b *= 255; }
      return [r, g, b];
    }
    var v = x <= 1 && x % 1 ? x * 255 : x; return [v, v, v];            // one grey value
  }
  var lum = function (c) { return (0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]) / 255; };
  /* fills: white for light tints, very light grey (keeps a little difference between colours) for the rest */
  function fill(c) {
    var L = lum(c);
    if (L >= 0.86) return 255;
    return Math.round(Math.max(214, Math.min(242, 214 + L * 30)));
  }
  /* lines: light lines stay light grey, the rest dark grey */
  function line(c) { return lum(c) >= 0.7 ? 165 : 70; }

  function patch(J) {
    if (!J || J.__inkSaver || !J.API) return;
    J.__inkSaver = true;
    (J.API.events = J.API.events || []).push(['initialized', function () {
      if (!on()) return;
      var d = this;
      ['setFillColor', 'setTextColor', 'setDrawColor'].forEach(function (m) {
        var orig = d[m]; if (typeof orig !== 'function') return;
        d[m] = function () {
          var c = rgb([].slice.call(arguments));
          if (!c) return orig.apply(d, arguments);
          var g = m === 'setTextColor' ? 0 : m === 'setFillColor' ? fill(c) : line(c);
          return orig.call(d, g, g, g);
        };
      });
    }]);
  }

  /* jsPDF is loaded later (often on demand): patch it the first time the page uses window.jspdf */
  var cur = window.jspdf;
  try {
    Object.defineProperty(window, 'jspdf', {
      configurable: true,
      get: function () { if (cur && cur.jsPDF) patch(cur.jsPDF); return cur; },
      set: function (v) { cur = v; }
    });
  } catch (e) { if (cur && cur.jsPDF) patch(cur.jsPDF); }
  window.JCW_INK = { patch: patch, on: on };

  /* printed web pages: black text, no backgrounds (the browser's own print) */
  try {
    var st = document.createElement('style');
    st.textContent = '@media print{html.ink *{-webkit-print-color-adjust:economy!important;print-color-adjust:economy!important;' +
      'color:#000!important;background:transparent!important;box-shadow:none!important;text-shadow:none!important}' +
      'html.ink th{border-bottom:1.5px solid #000!important}html.ink img{filter:grayscale(1)}}';
    (document.head || document.documentElement).appendChild(st);
    if (on()) document.documentElement.classList.add('ink');
  } catch (e) {}
})();
