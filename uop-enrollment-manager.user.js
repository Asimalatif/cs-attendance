// ==UserScript==
// @name         UOP VC - Enrollment Manager
// @namespace    UOP-JCW
// @version      3.5
// @description  From the JCW CSV: set Enrollment#, Reg# (= Enrollment#), Program, Class and Section on each student's edit page (matched by PORTAL ID)
// @match        https://portal.uop.edu.pk/Student/Students*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
  "use strict";

  // ==========================================================
  // CONFIGURATION
  // ==========================================================
  const CONFIG = {
    DELAY: 200,             // pause after each page load / dropdown change (ms)
    WORKERS: 3,             // students updated at the same time (hidden frames)
    CHECK_PARALLEL: 8,      // students checked at the same time (fast check, no page opening)
    PAGE_TIMEOUT: 30000,    // max wait for an edit page to load
    SAVE_TIMEOUT: 30000,    // max wait for save to finish
    CASCADE_TIMEOUT: 15000, // max wait for Class / Section dropdown to fill
    // Default edit-page address. {id} = PORTAL ID from CSV.
    // Use the "Detect" button on the panel if your portal uses a different one.
    EDIT_URL: "/Student/NewStudent/{id}",
    // CSV program text  ->  text used in the portal dropdown
    PROGRAM_MAP: { "(JCW)": "(Jinnah College)" }
  };

  // CSV column names (case/space-insensitive)
  const COLS = {
    id: ["PORTAL ID", "PORTALID", "Portal Id"],
    enroll: ["Enroll#", "Enrollment#", "Enrollment No", "Enroll No"],
    program: ["Program"],
    cls: ["Class"],
    section: ["Section"],
    name: ["Student Name", "Name"],
    father: ["Father Name"]
  };

  if (window.__UOP_ENROLL_MANAGER_INSTALLED) return;
  window.__UOP_ENROLL_MANAGER_INSTALLED = true;

  const LS = {
    rows: "uopEM_rows",
    done: "uopEM_done",
    url: "uopEM_editUrl",
    mode: "uopEM_mode"
  };

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const norm = (s) =>
    String(s || "").toLowerCase().replace(/jcw/g, "jinnahcollege").replace(/[^a-z0-9]/g, "");
  const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { log("⚠ Could not save to browser storage: " + e.message, "warn"); } };

  let rows = lsGet(LS.rows, []);
  let done = lsGet(LS.done, {});
  let report = [];
  let stopFlag = false;
  let running = false;

  // ==========================================================
  // CSV
  // ==========================================================
  function parseCSV(text) {
    text = text.replace(/^﻿/, "");
    const out = []; let row = []; let f = ""; let q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; }
        else f += c;
      } else if (c === '"') q = true;
      else if (c === ",") { row.push(f); f = ""; }
      else if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++;
        row.push(f); f = ""; out.push(row); row = [];
      } else f += c;
    }
    if (f !== "" || row.length) { row.push(f); out.push(row); }
    return out.filter((r) => r.some((x) => x.trim() !== ""));
  }

  function buildRows(text) {
    const data = parseCSV(text);
    if (data.length < 2) throw new Error("CSV is empty");
    const head = data[0].map((h) => norm(h));
    const idx = {};
    for (const k in COLS) {
      idx[k] = -1;
      for (const name of COLS[k]) { const i = head.indexOf(norm(name)); if (i >= 0) { idx[k] = i; break; } }
    }
    const need = ["id", "enroll", "program", "cls", "section"];
    const missing = need.filter((k) => idx[k] < 0);
    if (missing.length) throw new Error("CSV column(s) not found: " + missing.map((k) => COLS[k][0]).join(", "));

    const list = [];
    const seen = {};
    data.slice(1).forEach((r, n) => {
      const g = (k) => (idx[k] >= 0 ? String(r[idx[k]] || "").trim() : "");
      const id = g("id").replace(/\D/g, "");
      if (!id) return;
      const item = {
        line: n + 2, id, enroll: g("enroll"), program: g("program"),
        cls: g("cls"), section: g("section"), name: g("name"), father: g("father")
      };
      seen[id] = (seen[id] || 0) + 1;
      list.push(item);
    });
    list.forEach((x) => { x.dup = seen[x.id] > 1; });
    return list;
  }

  function portalProgram(csvProgram) {
    let p = csvProgram;
    for (const k in CONFIG.PROGRAM_MAP) p = p.split(k).join(CONFIG.PROGRAM_MAP[k]);
    return p;
  }

  // ==========================================================
  // PAGE LOADING (hidden frame, or popup window)
  // ==========================================================
  let popup = null;
  const frames = [];

  function getWin(ctx) {
    if (panelMode() === "popup") return popup && !popup.closed ? popup : null;
    return ctx && ctx.frame ? ctx.frame.contentWindow : null;
  }

  function safeDoc(win) {
    try { return win && win.document; } catch (e) { return null; }
  }

  function patchDialogs(win, sink) {
    try {
      win.alert = (m) => { sink.push(String(m)); };
      win.confirm = (m) => { sink.push(String(m)); return true; };
    } catch (e) { /* ignore */ }
  }

  async function loadPage(url, sink, ctx) {
    const oldDoc = safeDoc(getWin(ctx));
    if (panelMode() === "popup") {
      if (!popup || popup.closed) popup = window.open("about:blank", "uopEM_editWin", "width=1100,height=800");
      if (!popup) throw new Error("Popup blocked — allow popups for portal.uop.edu.pk");
      popup.location.href = url;
    } else {
      if (!ctx.frame) {
        ctx.frame = document.createElement("iframe");
        ctx.frame.className = "uopEM_frame";
        ctx.frame.style.cssText = "position:fixed;left:-10000px;top:0;width:1200px;height:900px;border:0;";
        document.body.appendChild(ctx.frame);
        frames.push(ctx.frame);
      }
      ctx.frame.src = url;
    }
    const t0 = Date.now();
    while (Date.now() - t0 < CONFIG.PAGE_TIMEOUT) {
      await sleep(250);
      const win = getWin(ctx);
      let doc;
      try { doc = win && win.document; }
      catch (e) { throw new Error("Portal blocks loading in a frame. Switch panel mode to 'Popup window'."); }
      if (doc && doc !== oldDoc && doc.readyState === "complete" && doc.location.href !== "about:blank") {
        patchDialogs(win, sink);
        if (/login|signin/i.test(doc.location.pathname)) throw new Error("LOGIN: session expired — log in again, then Start (it resumes).");
        await waitIdle(win, 1100); // page loads Program -> Class -> Section lists by AJAX (and ignores changes for 1 s)
        learnPrograms(doc);
        return { win, doc };
      }
    }
    throw new Error("Edit page did not load in time");
  }

  // ==========================================================
  // FIELD FINDING
  // ==========================================================
  function controls(doc) {
    return [...doc.querySelectorAll(
      "input:not([type=hidden]):not([type=button]):not([type=submit]):not([type=checkbox]):not([type=radio]):not([type=file]), select, textarea"
    )];
  }

  function labelOf(el) {
    const doc = el.ownerDocument;
    let t = "";
    if (el.id) {
      try { const l = doc.querySelector('label[for="' + CSS.escape(el.id) + '"]'); if (l) t += l.textContent + " "; } catch (e) {}
    }
    const wrap = el.closest("label");
    if (wrap) t += wrap.textContent + " ";
    if (!t.trim()) {
      const grp = el.closest(".form-group, .mb-3, .mb-2, .input-group, .form-floating");
      if (grp) { const l = grp.querySelector("label, .control-label, .form-label"); if (l) t += l.textContent + " "; }
    }
    if (!t.trim()) {
      const td = el.closest("td");
      if (td && td.previousElementSibling) t += td.previousElementSibling.textContent + " ";
    }
    if (!t.trim()) {
      let p = el.previousElementSibling;
      while (p && !t.trim()) { if (/^(LABEL|SPAN|B|STRONG|P|DIV)$/.test(p.tagName)) t += p.textContent + " "; p = p.previousElementSibling; }
    }
    return t.replace(/\s+/g, " ").trim();
  }

  const FIELD_RULES = {
    enroll:  { tag: "input",  inc: /enrol/i, exc: /date|status|type/i },
    reg:     { tag: "input",  inc: /\breg(istration)?\b|regno|reg\s*#|registration/i, exc: /enrol|date|fee|board|mobile|phone/i },
    program: { tag: "select", inc: /program/i, exc: /type|level/i },
    cls:     { tag: "select", inc: /class|year|semester/i, exc: /program|section|session|batch|passing|birth|date|admission/i },
    section: { tag: "select", inc: /section/i, exc: /^$/ }
  };

  // Exact element IDs on the UOP student page (/Student/NewStudent/{id})
  const FIELD_IDS = { enroll: "txtEnrollmentNo", reg: "txtRegistrationNo", program: "cmbProgramId", cls: "cmbClassId", section: "cmbSectionId" };

  function findField(doc, key) {
    const byId = doc.getElementById(FIELD_IDS[key]);
    if (byId) return byId;
    const r = FIELD_RULES[key];
    const els = controls(doc).filter((e) => (r.tag === "select" ? e.tagName === "SELECT" : e.tagName !== "SELECT"));
    // 1) by visible label
    let hit = els.find((e) => { const l = labelOf(e); return l && r.inc.test(l) && !r.exc.test(l); });
    // 2) by name / id attribute
    if (!hit) hit = els.find((e) => { const a = (e.name || "") + " " + (e.id || ""); return r.inc.test(a) && !r.exc.test(a); });
    return hit || null;
  }

  function describeFields(doc) {
    return controls(doc).map((e) => "[" + e.tagName.toLowerCase() + "] " + (labelOf(e) || "?") + " (" + (e.name || e.id || "") + ")").join(" | ");
  }

  const selText = (sel) => (sel && sel.selectedIndex >= 0 ? sel.options[sel.selectedIndex].text.trim() : "");

  function isPlaceholder(o) {
    return o.value === "" || o.value === "0" || o.value === "-1" || /^\s*-*\s*(all|select|choose)\b/i.test(o.text);
  }

  function findOption(sel, wanted, kind) {
    if (!sel) return null;
    const opts = [...sel.options].filter((o) => !isPlaceholder(o));
    const w = norm(wanted);
    let o = opts.find((x) => norm(x.text) === w);
    if (o) return o;
    if (kind === "program") {
      const core = norm(wanted.replace(/\(.*?\)/g, ""));
      const same = opts.filter((x) => norm(x.text.replace(/\(.*?\)/g, "")) === core);
      o = same.find((x) => /jinnah|jcw/i.test(x.text)) || (same.length === 1 ? same[0] : null);
    } else if (kind === "cls") {
      const n = (wanted.match(/\d/) || [""])[0];
      const word = { 1: "first", 2: "second" }[n] || "@@";
      o = opts.find((x) => { const t = norm(x.text); return t.startsWith(n + "styear") || t.startsWith(n + "ndyear") || t.startsWith(word + "year") || t === n + "st" || t === n + "nd"; });
    } else if (kind === "section") {
      o = opts.find((x) => norm(x.text).replace(/^section/, "") === w.replace(/^section/, ""));
    }
    return o || null;
  }

  function fire(win, el, type) {
    el.dispatchEvent(new win.Event(type, { bubbles: true }));
  }

  function setSelect(win, sel, opt) {
    sel.value = opt.value;
    opt.selected = true;
    if (win.jQuery) { win.jQuery(sel).val(opt.value).trigger("change"); } // also refreshes select2
    else fire(win, sel, "change");
  }

  // Wait until the page has no pending AJAX (dropdown lists finished loading)
  async function waitIdle(win, extra) {
    await sleep(extra || 450); // portal debounces dropdown changes by 300 ms
    const t0 = Date.now();
    let quiet = 0;
    while (Date.now() - t0 < CONFIG.CASCADE_TIMEOUT) {
      const busy = win.jQuery ? win.jQuery.active > 0 : false;
      quiet = busy ? 0 : quiet + 1;
      if (quiet >= 3) return;
      await sleep(250);
    }
  }

  function setText(win, el, value) {
    const proto = Object.getPrototypeOf(el);
    const desc = Object.getOwnPropertyDescriptor(proto, "value");
    if (desc && desc.set) desc.set.call(el, value); else el.value = value;
    fire(win, el, "input");
    fire(win, el, "change");
  }

  // Wait until dropdown `key` contains the wanted option (cascading AJAX)
  async function waitOption(doc, key, wanted, kind) {
    const t0 = Date.now();
    while (Date.now() - t0 < CONFIG.CASCADE_TIMEOUT) {
      const sel = findField(doc, key);
      const o = findOption(sel, wanted, kind);
      if (o) return { sel, opt: o };
      await sleep(300);
    }
    return { sel: findField(doc, key), opt: null };
  }

  // ==========================================================
  // READ / WRITE ONE STUDENT
  // ==========================================================
  function readCurrent(doc) {
    const f = {
      enroll: findField(doc, "enroll"), reg: findField(doc, "reg"),
      program: findField(doc, "program"), cls: findField(doc, "cls"), section: findField(doc, "section")
    };
    return {
      f,
      enroll: f.enroll ? f.enroll.value.trim() : null,
      reg: f.reg ? f.reg.value.trim() : null,
      program: f.program ? selText(f.program) : null,
      cls: f.cls ? selText(f.cls) : null,
      section: f.section ? selText(f.section) : null
    };
  }

  function diffs(cur, s) {
    const d = [];
    const wantProg = portalProgram(s.program);
    if (cur.f.enroll && cur.enroll !== s.enroll) d.push("Enroll# '" + cur.enroll + "' → '" + s.enroll + "'");
    if (cur.f.reg && cur.reg !== s.enroll) d.push("Reg# '" + cur.reg + "' → '" + s.enroll + "'");
    if (cur.f.program && norm(cur.program) !== norm(wantProg)) d.push("Program '" + cur.program + "' → '" + wantProg + "'");
    if (cur.f.cls && !findOptionMatchesText(cur.cls, s.cls)) d.push("Class '" + cur.cls + "' → '" + s.cls + "'");
    if (cur.f.section && norm(cur.section).replace(/^section/, "") !== norm(s.section)) d.push("Section '" + cur.section + "' → '" + s.section + "'");
    return d;
  }

  function findOptionMatchesText(text, wanted) {
    const fake = { options: [{ text, value: "x" }] };
    return !!findOption(fake, wanted, "cls");
  }

  async function processStudent(s, dryRun, ctx) {
    const sink = [];
    const url = editUrl().replace("{id}", encodeURIComponent(s.id));
    let { win, doc } = await loadPage(url, sink, ctx);

    // Safety: make sure the page opened is the same student as the CSV row
    const pgName = doc.getElementById("txtStudentName"), pgFather = doc.getElementById("txtFatherName");
    if (pgName && !nameOk(pgName.value, pgFather ? pgFather.value : "", s))
      return { status: "ERROR", msg: "Name mismatch — portal has '" + pgName.value + " / " + (pgFather ? pgFather.value : "") + "', not changed" };

    let cur = readCurrent(doc);
    const missing = ["enroll", "reg", "program", "cls", "section"].filter((k) => !cur.f[k]);
    if (missing.length === 5) return { status: "ERROR", msg: "No edit form found at " + url + " (check the Edit URL)" };
    if (missing.length) log("  ⚠ Field(s) not found on page: " + missing.join(", ") + ". Fields seen: " + describeFields(doc), "warn");

    const d = diffs(cur, s);
    let q = null;
    if (s.obt) {                                  // v3.4: matric / O-Level (IBCC) marks = first row of the Qualifications table
      q = await qualRow(doc);
      if (!q) return { status: "ERROR", msg: "Qualifications table did not load — fix marks by hand" };
      if (num(q.obt.value) !== num(s.obt) || num(q.tot.value) !== num(s.tot)) d.push("Marks " + q.obt.value + "/" + q.tot.value + " → " + s.obt + "/" + s.tot); else q = null;
    }
    if (!d.length) return { status: "SAME", msg: "Already correct" };
    if (dryRun) return { status: "WOULD CHANGE", msg: d.join("; ") };

    // --- Program -> Class -> Section (in order, waiting for each cascade) ---
    const wantProg = portalProgram(s.program);
    let r = await waitOption(doc, "program", wantProg, "program");
    if (!r.sel) return { status: "ERROR", msg: "Program dropdown not found" };
    if (!r.opt) return { status: "ERROR", msg: "Program '" + wantProg + "' not in dropdown. Options: " + [...r.sel.options].map((o) => o.text).join(", ") };
    const oldTbl = doc.querySelector("#tblStuQualification"); let progChanged = false;
    if (r.sel.value !== r.opt.value) { setSelect(win, r.sel, r.opt); progChanged = true; await waitIdle(win); }

    if (findField(doc, "cls")) {
      r = await waitOption(doc, "cls", s.cls, "cls");
      if (!r.opt) return { status: "ERROR", msg: "Class '" + s.cls + "' not in dropdown. Options: " + [...(r.sel ? r.sel.options : [])].map((o) => o.text).join(", ") };
      // Always fire change on Class so the Section list loads
      setSelect(win, r.sel, r.opt);
      await waitIdle(win);
    }

    r = await waitOption(doc, "section", s.section, "section");
    if (!r.opt) {
      // retry: re-select Class to reload the Section list
      const c = findField(doc, "cls"); const co = findOption(c, s.cls, "cls");
      if (c && co) { setSelect(win, c, co); await waitIdle(win, 1000); }
      r = await waitOption(doc, "section", s.section, "section");
    }
    if (!r.sel) return { status: "ERROR", msg: "Section dropdown not found" };
    if (!r.opt) return { status: "ERROR", msg: "Section '" + s.section + "' not in dropdown. Options: " + [...r.sel.options].map((o) => o.text).join(", ") };
    setSelect(win, r.sel, r.opt);
    await sleep(300);

    // --- v3.4 marks (set after Program, because a Program change reloads the Qualifications table) ---
    if (q) {
      for (let t = 0; progChanged && t < 40 && doc.querySelector("#tblStuQualification") === oldTbl; t++) await sleep(250);   // the new Program's table replaces the old one
      await sleep(600);
      q = await qualRow(doc);
      if (!q) return { status: "ERROR", msg: "Qualifications table did not load — fix marks by hand" };
      if (q.incomplete) return { status: "ERROR", msg: "A qualification row on the portal is incomplete (" + q.incomplete + ") — saving could drop it, fix by hand" };
      setText(win, q.obt, String(s.obt)); setText(win, q.tot, String(s.tot));
      if (q.pct) setText(win, q.pct, (Math.round(num(s.obt) / num(s.tot) * 10000) / 100).toFixed(2));
      await sleep(200);
    }

    // --- Enrollment# and Reg# ---
    const en = findField(doc, "enroll"); if (en && en.value.trim() !== s.enroll) setText(win, en, s.enroll);
    const rg = findField(doc, "reg");    if (rg && rg.value.trim() !== s.enroll) setText(win, rg, s.enroll);
    await sleep(300);

    // --- Save (portal's own "Save" = SaveRecord(3): saves and stays on the page) ---
    const oldDoc = doc;
    if (typeof win.SaveRecord === "function") {
      win.SaveRecord(3);
    } else {
      const form = (en || r.sel).closest("form") || doc.querySelector("form");
      const scope = form || doc;
      let btn = [...scope.querySelectorAll("button, input[type=submit], input[type=button], a.btn")]
        .find((b) => /save|update|submit/i.test((b.textContent || "") + " " + (b.value || "")) && !/cancel|back|delete/i.test(b.textContent || b.value || ""));
      if (!btn && form) btn = form.querySelector("[type=submit]");
      if (!btn) return { status: "ERROR", msg: "Save button not found" };
      btn.click();
    }

    const t0 = Date.now();
    let navigated = false;
    while (Date.now() - t0 < CONFIG.SAVE_TIMEOUT) {
      await sleep(400);
      const nd = safeDoc(getWin(ctx));
      if (nd && nd !== oldDoc && nd.readyState === "complete") { navigated = true; doc = nd; patchDialogs(getWin(ctx), sink); break; }
      const errNow = pageErrors(oldDoc);
      if (errNow) return { status: "ERROR", msg: "Portal says: " + errNow };
      if (sink.some((m) => /success|saved|updated/i.test(m)) || successText(oldDoc)) break;
      if (sink.some((m) => /error|fail|invalid|already/i.test(m))) return { status: "ERROR", msg: "Portal says: " + sink.join(" | ") };
    }
    await sleep(CONFIG.DELAY);
    const err = pageErrors(doc);
    if (err) return { status: "ERROR", msg: "Portal says: " + err };

    // --- Verify (fast: re-reads the saved page without opening it) ---
    if (verifyOn()) {
      await sleep(300);
      const q = await quickCheck(s);
      const q2 = s.obt ? { diffs: null } : await quickCheck(s);
      if (q2.diffs && q2.diffs.length) return { status: "ERROR", msg: "Saved but still different: " + q2.diffs.join("; ") };
      if (s.obt) { const m = await readMarks(s.id); if (m && (num(m.obt) !== num(s.obt) || num(m.tot) !== num(s.tot))) return { status: "ERROR", msg: "Saved but marks still " + m.obt + "/" + m.tot }; if (m) return { status: "UPDATED", msg: d.join("; ") + " ✔ verified" }; }
      return { status: "UPDATED", msg: d.join("; ") + (q2.diffs ? " ✔ verified" : "") };
    }
    return { status: "UPDATED", msg: d.join("; ") + (navigated ? "" : " (no page reload seen)") };
  }

  // v3.4 Qualifications table helpers (row 0 = Matric / O-Level)
  const num = (x) => { const n = parseFloat(String(x == null ? "" : x).replace(/,/g, "")); return isNaN(n) ? -1 : Math.round(n * 100) / 100; };
  async function qualRow(doc) {
    for (let t = 0; t < 40; t++) {
      const rows = [...doc.querySelectorAll("#tblStuQualification tbody tr")];
      const r = rows[0];
      if (r && r.querySelector(".obt-marks-class")) {
        const need = [".degree-class", ".board-class", ".passing-year-class", ".roll-no-class", ".total-marks-class", ".obt-marks-class", ".division-class", ".percentage-marks-class"];
        let incomplete = "";
        rows.forEach((x, i) => need.forEach((c) => { const e = x.querySelector(c); if (!incomplete && (!e || !String(e.value).trim() || (c === ".degree-class" && e.value == 0))) incomplete = "row " + (i + 1) + " " + c.replace(/^\.|-class$/g, ""); }));
        return { obt: r.querySelector(".obt-marks-class"), tot: r.querySelector(".total-marks-class"), pct: r.querySelector(".percentage-marks-class"), incomplete };
      }
      await sleep(250);
    }
    return null;
  }
  async function readMarks(id) {
    try {
      const st = await readQuick(id); if (!st.programId) return null;
      const r = await fetch("/Student/StudentQualificationPartail", { method: "POST", credentials: "include", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: "Id=" + encodeURIComponent(id) + "&ProgramId=" + encodeURIComponent(st.programId) });
      const d = new DOMParser().parseFromString(await r.text(), "text/html"), row = d.querySelector("#tblStuQualification tbody tr");
      if (!row) return null;
      const v = (c) => { const e = row.querySelector(c); return e ? (e.getAttribute("value") || e.value || "") : ""; };
      return { obt: v(".obt-marks-class"), tot: v(".total-marks-class") };
    } catch (e) { return null; }
  }

  function swalInfo(doc) {
    const m = doc && doc.querySelector(".swal-overlay--show-modal .swal-modal, .swal-modal");
    if (!m || !doc.querySelector(".swal-overlay--show-modal")) return null;
    const title = (m.querySelector(".swal-title") || {}).textContent || "";
    const text = (m.querySelector(".swal-text") || {}).textContent || "";
    const bad = m.querySelector(".swal-icon--error, .swal-icon--warning") || /error|warning/i.test(title);
    return { ok: !bad && /success|saved/i.test(title + " " + text), bad: !!bad, msg: (title + " " + text).trim() };
  }

  function pageErrors(doc) {
    if (!doc) return "";
    const sw = swalInfo(doc);
    if (sw && sw.bad) return sw.msg;
    const els = doc.querySelectorAll(".field-validation-error, .validation-summary-errors li, .alert-danger, .text-danger, .swal2-error ~ .swal2-html-container, .toast-error");
    const t = [...els].map((e) => e.textContent.trim()).filter((x) => x && x !== "*").join(" | ");
    return t;
  }
  function successText(doc) {
    if (!doc) return false;
    const sw = swalInfo(doc);
    if (sw && sw.ok) return true;
    const els = doc.querySelectorAll(".alert-success, .toast-success, .swal2-success, .swal2-html-container");
    return [...els].some((e) => /success|saved|updated/i.test(e.textContent) || e.classList.contains("swal2-success"));
  }

  function nameOk(name, father, s) {
    const first = (x) => norm(String(x || "").trim().split(/\s+/)[0]);
    const okName = norm(name) === norm(s.name) || (first(name) && first(name) === first(s.name));
    const okFather = s.father && father && (norm(father) === norm(s.father) || first(father) === first(s.father));
    return okName || okFather;
  }

  // ==========================================================
  // FAST CHECK — reads the edit page as plain text (no scripts, no images)
  // ==========================================================
  const programOpts = {};               // portal program id -> name (learned once from a real page)
  const classCache = {}, sectionCache = {};

  function learnPrograms(doc) {
    const sel = doc && doc.getElementById("cmbProgramId");
    if (!sel) return;
    [...sel.options].forEach((o) => { if (!isPlaceholder(o)) programOpts[o.value] = o.text.trim(); });
  }

  async function postJSON(url, data) {
    const r = await fetch(url, {
      method: "POST", credentials: "include",
      headers: { "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8", "X-Requested-With": "XMLHttpRequest" },
      body: new URLSearchParams(data).toString()
    });
    if (!r.ok) throw new Error(url + " HTTP " + r.status);
    return r.json();
  }
  function classesOf(pid) {
    if (!classCache[pid]) classCache[pid] = postJSON("/Program/LaodClasses", { ProgramId: pid, Option: "" }).catch((e) => { delete classCache[pid]; throw e; });
    return classCache[pid];
  }
  function sectionsOf(cid) {
    if (!sectionCache[cid]) sectionCache[cid] = postJSON("/Program/LoadSections", { ClassId: cid, Option: "OptionAll" }).catch((e) => { delete sectionCache[cid]; throw e; });
    return sectionCache[cid];
  }
  const fakeSel = (pairs) => ({ options: pairs.map(([value, text]) => ({ value: String(value), text: String(text) })) });

  // The page writes the saved Program/Class/Section ids into its script: var selectedId = '1142';
  function embeddedId(html, fn) {
    const m = html.match(new RegExp("function\\s+" + fn + "\\s*\\([\\s\\S]{0,800}?selectedId\\s*=\\s*['\"]?(\\d*)"));
    return m ? m[1] : null;
  }

  async function readQuick(id) {
    const r = await fetch(editUrl().replace("{id}", encodeURIComponent(id)), { credentials: "include" });
    if (/login|signin/i.test(new URL(r.url).pathname)) throw new Error("LOGIN: session expired — log in again, then Start (it resumes).");
    const html = await r.text();
    const doc = new DOMParser().parseFromString(html, "text/html");
    const v = (x) => { const e = doc.getElementById(x); return e ? (e.getAttribute("value") || "").trim() : null; };
    return {
      name: v("txtStudentName"), father: v("txtFatherName"), enroll: v("txtEnrollmentNo"), reg: v("txtRegistrationNo"),
      programId: embeddedId(html, "GetPrograms"), classId: embeddedId(html, "GetClasses"), sectionId: embeddedId(html, "GetSections")
    };
  }

  // Returns {diffs:[...]} when sure, {diffs:null} when the page must be opened, {error} on a definite problem
  async function quickCheck(s) {
    if (s.obt) return { diffs: null };          // v3.4: marks live in the Qualifications tab: open the page
    let st;
    try { st = await readQuick(s.id); }
    catch (e) { if (/^LOGIN/.test(e.message)) throw e; return { diffs: null }; }
    if ([st.name, st.enroll, st.reg, st.programId, st.classId, st.sectionId].some((x) => x === null)) return { diffs: null };
    if (!nameOk(st.name, st.father, s)) return { error: "Name mismatch — portal has '" + st.name + " / " + st.father + "', not changed" };
    if (!Object.keys(programOpts).length) return { diffs: null };
    try {
      const d = [];
      if (st.enroll !== s.enroll) d.push("Enroll# '" + st.enroll + "' → '" + s.enroll + "'");
      if (st.reg !== s.enroll) d.push("Reg# '" + st.reg + "' → '" + s.enroll + "'");
      const wantProg = portalProgram(s.program);
      const pOpt = findOption(fakeSel(Object.entries(programOpts)), wantProg, "program");
      if (!pOpt) return { error: "Program '" + wantProg + "' not in portal list" };
      if (st.programId !== pOpt.value) d.push("Program '" + (programOpts[st.programId] || "?") + "' → '" + pOpt.text + "'");
      const cls = await classesOf(pOpt.value);
      const cOpt = findOption(fakeSel(cls.map((x) => [x.Id, x.Name])), s.cls, "cls");
      if (!cOpt) return { error: "Class '" + s.cls + "' not found for " + pOpt.text };
      if (st.classId !== cOpt.value) d.push("Class → '" + cOpt.text + "'");
      const secs = await sectionsOf(cOpt.value);
      const sOpt = findOption(fakeSel(secs.map((x) => [x.Id, x.Section])), s.section, "section");
      if (!sOpt) return { error: "Section '" + s.section + "' not found in " + pOpt.text + " " + cOpt.text + ". Sections: " + secs.map((x) => x.Section).join(", ") };
      if (st.sectionId !== sOpt.value) {
        const now = (secs.find((x) => String(x.Id) === st.sectionId) || {}).Section;
        d.push("Section '" + (now || "--All--") + "' → '" + sOpt.text + "'");
      }
      return { diffs: d };
    } catch (e) { return { diffs: null }; }
  }

  // ==========================================================
  // RUNNER
  // ==========================================================
  function selectedRows() {
    const ids = $("#uopEM_ids").value.split(/[\s,;]+/).map((x) => x.replace(/\D/g, "")).filter(Boolean);
    let list = rows;
    if (ids.length) list = rows.filter((r) => ids.includes(r.id));
    const cls = $("#uopEM_cls").value;
    if (cls) list = list.filter((r) => norm(r.cls) === norm(cls));
    const prog = $("#uopEM_prog").value;
    if (prog) list = list.filter((r) => r.program === prog);
    const from = parseInt($("#uopEM_from").value, 10) || 1;
    list = list.filter((r, i) => i + 1 >= from);
    const lim = parseInt($("#uopEM_limit").value, 10) || 0;
    if (lim > 0) list = list.slice(0, lim);
    return list;
  }

  async function run(dryRun) {
    if (running) return;
    if (!rows.length) { log("Load the CSV first.", "err"); return; }
    if (!dryRun && !$("#uopEM_ids").value.trim() && !Object.keys(done).length &&
        !confirm("No test IDs given. Update ALL " + selectedRows().length + " selected students now?\n\nTip: first try 2-3 PORTAL IDs in the 'Only IDs' box.")) return;

    running = true; stopFlag = false;
    try {
      const list = selectedRows();
      const skipDone = $("#uopEM_skip").checked;
      const c = { UPDATED: 0, SAME: 0, "WOULD CHANGE": 0, ERROR: 0, SKIPPED: 0 };
      let finished = 0;
      const t0 = Date.now();
      log((dryRun ? "🔍 CHECK (no saving)" : "▶ UPDATE") + " — " + list.length + " student(s)", "head");

      const finish = (s, res, quiet) => {
        finished++;
        c[res.status] = (c[res.status] || 0) + 1;
        addReport(s, res.status, res.msg);
        if (!quiet) {
          const tag = "#" + s.id + " " + s.enroll + " " + s.name + " (" + s.program + " / " + s.cls + " / " + s.section + ")";
          const icon = { UPDATED: "✅", SAME: "✔", "WOULD CHANGE": "✏", ERROR: "✖", SKIPPED: "⏭" }[res.status] || "•";
          log(icon + " " + tag + " — " + res.status + ": " + res.msg, res.status === "ERROR" ? "err" : res.status === "SAME" ? "dim" : res.status === "SKIPPED" ? "warn" : "ok");
        }
        if (!dryRun && (res.status === "UPDATED" || res.status === "SAME")) {
          done[s.id] = s.enroll + "|" + s.program + "|" + s.cls + "|" + s.section;
          lsSet(LS.done, done);
        }
        setProgress(finished, list.length, c);
      };

      const todo = [];
      for (const s of list) {
        if (s.dup) { finish(s, { status: "SKIPPED", msg: "PORTAL ID repeated in CSV — fix CSV" }); continue; }
        if (!dryRun && skipDone && done[s.id] === s.enroll + "|" + s.program + "|" + s.cls + "|" + s.section) { finish(s, { status: "SKIPPED", msg: "done earlier" }, true); continue; }
        todo.push(s);
      }

      const nWorkers = panelMode() === "popup" ? 1 : Math.min(6, Math.max(1, parseInt($("#uopEM_workers").value, 10) || CONFIG.WORKERS));
      const ctxs = Array.from({ length: nWorkers }, () => ({ frame: null }));
      let fatal = null;

      // Learn the portal's Program list once (needed for the fast check)
      if (todo.length && !Object.keys(programOpts).length) {
        log("Reading the portal's Program list…", "dim");
        try { await loadPage(editUrl().replace("{id}", encodeURIComponent(todo[0].id)), [], ctxs[0]); }
        catch (e) { fatal = e.message; }
      }

      // Phase 1: fast check of everyone (no page opening)
      const needPage = [];
      let qi = 0;
      const checker = async () => {
        while (!stopFlag && !fatal && qi < todo.length) {
          const s = todo[qi++];
          let q;
          try { q = await quickCheck(s); } catch (e) { fatal = e.message; return; }
          if (q.error) finish(s, { status: "ERROR", msg: q.error });
          else if (q.diffs && !q.diffs.length) finish(s, { status: "SAME", msg: "Already correct" }, !dryRun && list.length > 50);
          else if (q.diffs && dryRun) finish(s, { status: "WOULD CHANGE", msg: q.diffs.join("; ") });
          else needPage.push(s);
        }
      };
      if (!fatal) await Promise.all(Array.from({ length: CONFIG.CHECK_PARALLEL }, checker));

      // Phase 2: open only the students that need changing
      if (!fatal && !stopFlag && needPage.length)
        log("Fast check done. " + needPage.length + " student(s) need " + (dryRun ? "a full check" : "updating") + " — " + nWorkers + " at a time…", "head");
      let pi = 0;
      const worker = async (ctx) => {
        while (!stopFlag && !fatal && pi < needPage.length) {
          const s = needPage[pi++];
          let res;
          try { res = await processStudent(s, dryRun, ctx); }
          catch (e) { res = { status: "ERROR", msg: e.message }; if (/^LOGIN|frame/i.test(e.message)) fatal = e.message; }
          finish(s, res);
        }
      };
      if (!fatal) await Promise.all(ctxs.map(worker));

      frames.forEach((f) => f.remove()); frames.length = 0;
      if (fatal) log("✖ " + fatal, "err");
      if (stopFlag) log("⏹ Stopped by user.", "warn");
      const mins = ((Date.now() - t0) / 60000).toFixed(1);
      log("Finished in " + mins + " min. Updated " + c.UPDATED + ", already correct " + c.SAME + ", would change " + c["WOULD CHANGE"] + ", errors " + c.ERROR + ", skipped " + c.SKIPPED + ". Use '⬇ Report'.", "head");
    } finally {
      running = false;
    }
  }

  function addReport(s, status, msg) {
    report.push({ time: new Date().toLocaleString(), id: s.id, enroll: s.enroll, name: s.name, father: s.father, program: s.program, cls: s.cls, section: s.section, status, msg });
  }

  function downloadReport() {
    if (!report.length) { log("Nothing to download yet.", "warn"); return; }
    const head = ["Time", "PORTAL ID", "Enroll#", "Student Name", "Father Name", "Program", "Class", "Section", "Result", "Details"];
    const esc = (v) => '"' + String(v == null ? "" : v).replace(/"/g, '""') + '"';
    const lines = [head.map(esc).join(",")].concat(report.map((r) => [r.time, r.id, r.enroll, r.name, r.father, r.program, r.cls, r.section, r.status, r.msg].map(esc).join(",")));
    const blob = new Blob(["﻿" + lines.join("\r\n")], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "UOP-Enrollment-Report-" + new Date().toISOString().slice(0, 16).replace(/[:T]/g, "-") + ".csv";
    a.click();
  }

  // Find an Edit link on the current Students page and turn it into a template
  function detectEditUrl() {
    const links = [...document.querySelectorAll("a[href]")].map((a) => a.getAttribute("href")).filter((h) => /edit|newstudent/i.test(h) && /\d{3,}/.test(h));
    if (!links.length) { log("No Edit link found on this page. Search a student so the list shows an Edit button, then press Detect again — or type the address yourself using {id}.", "warn"); return; }
    const sample = links[0];
    const tpl = sample.replace(/(\d{3,})(?!.*\d{3,})/, "{id}");
    $("#uopEM_url").value = tpl;
    localStorage.setItem(LS.url, tpl);
    log("Edit URL detected: " + tpl + "  (from " + sample + ")", "ok");
  }

  // ==========================================================
  // PANEL UI
  // ==========================================================
  const $ = (q) => document.querySelector(q);
  const editUrl = () => ($("#uopEM_url").value.trim() || CONFIG.EDIT_URL);
  const panelMode = () => ($("#uopEM_mode") ? $("#uopEM_mode").value : "frame");
  const verifyOn = () => $("#uopEM_verify").checked;

  function log(msg, cls) {
    const box = $("#uopEM_log");
    if (!box) { console.log("[UOP-EM]", msg); return; }
    const d = document.createElement("div");
    d.className = "uopEM_" + (cls || "n");
    d.textContent = msg;
    box.appendChild(d);
    while (box.childNodes.length > 1500) box.removeChild(box.firstChild);
    box.scrollTop = box.scrollHeight;
  }

  function setProgress(i, n, c) {
    const pct = n ? Math.round((i / n) * 100) : 0;
    $("#uopEM_bar").style.width = pct + "%";
    $("#uopEM_stat").textContent = i + " / " + n + "  •  ✅ " + c.UPDATED + "  ✔ " + c.SAME + "  ✏ " + c["WOULD CHANGE"] + "  ✖ " + c.ERROR;
  }

  function refreshFilters() {
    const progs = [...new Set(rows.map((r) => r.program))].sort();
    $("#uopEM_prog").innerHTML = '<option value="">All programs</option>' + progs.map((p) => '<option>' + p.replace(/</g, "&lt;") + "</option>").join("");
    const dups = rows.filter((r) => r.dup).map((r) => r.id);
    $("#uopEM_info").textContent = rows.length ? rows.length + " students loaded" + (dups.length ? " • repeated PORTAL ID: " + [...new Set(dups)].join(", ") : "") : "No CSV loaded";
  }

  function buildPanel() {
    const css = document.createElement("style");
    css.textContent = `
#uopEM{position:fixed;right:12px;bottom:12px;width:440px;z-index:2147483647;background:#fff;border:2px solid #0b5394;border-radius:10px;box-shadow:0 6px 24px rgba(0,0,0,.25);font:13px Segoe UI,Arial,sans-serif;color:#222}
#uopEM header{background:#0b5394;color:#fff;padding:7px 10px;font-weight:600;display:flex;justify-content:space-between;align-items:center;border-radius:7px 7px 0 0;cursor:pointer}
#uopEM .b{padding:8px 10px}
#uopEM .r{display:flex;gap:6px;align-items:center;margin:4px 0;flex-wrap:wrap}
#uopEM input[type=text],#uopEM select{border:1px solid #bbb;border-radius:4px;padding:3px 5px;font-size:12px}
#uopEM button{border:0;border-radius:5px;padding:6px 10px;font-size:12px;font-weight:600;cursor:pointer;color:#fff;background:#0b5394}
#uopEM button.g{background:#38761d}#uopEM button.o{background:#b45f06}#uopEM button.x{background:#990000}#uopEM button.s{background:#666}
#uopEM_log{height:190px;overflow:auto;background:#f6f8fa;border:1px solid #ddd;border-radius:5px;padding:4px;font:11px Consolas,monospace;white-space:pre-wrap}
#uopEM .uopEM_err{color:#b00}#uopEM .uopEM_ok{color:#1b5e20}#uopEM .uopEM_warn{color:#a35a00}#uopEM .uopEM_dim{color:#888}#uopEM .uopEM_head{font-weight:700;color:#0b5394}
#uopEM_track{height:8px;background:#e5e5e5;border-radius:4px;overflow:hidden}#uopEM_bar{height:100%;width:0;background:#38761d}
#uopEM small{color:#666}`;
    document.head.appendChild(css);

    const p = document.createElement("div");
    p.id = "uopEM";
    p.innerHTML = `
<header id="uopEM_hdr"><span>🎓 UOP Enrollment Manager v3.5 (JCW)</span><span id="uopEM_min">–</span></header>
<div class="b" id="uopEM_body">
  <div class="r"><button class="g" id="uopEM_app">⬇ Load wrong values from JCW app</button><small>or CSV:</small><input type="file" id="uopEM_file" accept=".csv"><small id="uopEM_info"></small></div>
  <div class="r"><small>Edit URL</small><input type="text" id="uopEM_url" style="flex:1"><button class="s" id="uopEM_detect">Detect</button></div>
  <div class="r"><small>Only IDs</small><input type="text" id="uopEM_ids" placeholder="PORTAL IDs e.g. 115108, 124026  (blank = all)" style="flex:1"></div>
  <div class="r">
    <select id="uopEM_prog"></select>
    <select id="uopEM_cls"><option value="">All classes</option><option>1st Year</option><option>2nd Year</option></select>
    <small>From #</small><input type="text" id="uopEM_from" value="1" style="width:45px">
    <small>Max</small><input type="text" id="uopEM_limit" value="" style="width:45px">
  </div>
  <div class="r">
    <label><input type="checkbox" id="uopEM_verify" checked> Verify after save</label>
    <label><input type="checkbox" id="uopEM_skip" checked> Skip already done</label>
    <small>Parallel</small><input type="text" id="uopEM_workers" value="3" style="width:28px">
    <select id="uopEM_mode"><option value="frame">Hidden frame</option><option value="popup">Popup window</option></select>
  </div>
  <div class="r">
    <button class="o" id="uopEM_check">🔍 Check only</button>
    <button class="g" id="uopEM_start">▶ Update</button>
    <button class="x" id="uopEM_stop">⏹ Stop</button>
    <button id="uopEM_dl">⬇ Report</button>
    <button class="s" id="uopEM_reset">Reset done</button>
  </div>
  <div id="uopEM_track"><div id="uopEM_bar"></div></div>
  <div class="r"><small id="uopEM_stat"></small></div>
  <div id="uopEM_log"></div>
</div>`;
    document.body.appendChild(p);

    const savedUrl = localStorage.getItem(LS.url);
    $("#uopEM_url").value = savedUrl && !/\/Students\/Edit\//i.test(savedUrl) ? savedUrl : CONFIG.EDIT_URL;
    $("#uopEM_mode").value = localStorage.getItem(LS.mode) || "frame";
    refreshFilters();

    $("#uopEM_hdr").onclick = () => {
      const b = $("#uopEM_body"); const hide = b.style.display !== "none";
      b.style.display = hide ? "none" : ""; $("#uopEM_min").textContent = hide ? "+" : "–";
    };
    $("#uopEM_url").onchange = () => localStorage.setItem(LS.url, $("#uopEM_url").value.trim());
    $("#uopEM_mode").onchange = () => localStorage.setItem(LS.mode, $("#uopEM_mode").value);
    $("#uopEM_detect").onclick = detectEditUrl;
    $("#uopEM_file").onchange = (e) => {
      const f = e.target.files[0]; if (!f) return;
      const rd = new FileReader();
      rd.onload = () => {
        try {
          rows = buildRows(rd.result);
          lsSet(LS.rows, rows);
          refreshFilters();
          log("CSV loaded: " + rows.length + " students (" + f.name + ")", "ok");
        } catch (err) { log("CSV error: " + err.message, "err"); }
      };
      rd.readAsText(f, "UTF-8");
    };
    $("#uopEM_app").onclick = () => loadFromApp(false);
    $("#uopEM_check").onclick = () => run(true);
    $("#uopEM_start").onclick = () => run(false);
    $("#uopEM_stop").onclick = () => { stopFlag = true; log("Stopping after current student…", "warn"); };
    $("#uopEM_dl").onclick = downloadReport;
    $("#uopEM_reset").onclick = () => { if (confirm("Forget which students were already done?")) { done = {}; lsSet(LS.done, done); log("Done-list cleared.", "warn"); } };

    if (/uopem-fix/i.test(location.hash)) { history.replaceState(null, "", location.pathname + location.search); setTimeout(() => loadFromApp(true), 800); return; }
    if (rows.length) log("Previous CSV still loaded (" + rows.length + " students). Choose a new file to replace it.", "dim");
  }

  // v3.3: load the students whose Program / Class / Section / Enroll# are wrong on the portal straight from the JCW app (no CSV)
  const JCW_API = "https://script.google.com/macros/s/AKfycbzl-DwgASIhf-pIGYi9D3M-CGOm7d7-UKg0W-OoWAiuRBHe4mjlRgMQXIHuLm6zi--i/exec";
  async function loadFromApp(auto) {
    let key = localStorage.getItem("uopEM_key") || "";
    if (!key) { key = (prompt("Sync key from the JCW attendance app\n(College > UOP portal sync > Get sync key):") || "").trim(); if (!key) return; }
    log("Reading the wrong values from the JCW app…", "dim");
    try {
      const res = await fetch(JCW_API, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ action: "portal_fixlist", key }) });
      const j = await res.json();
      if (!j.ok) { if (/key/i.test(j.error || "")) localStorage.removeItem("uopEM_key"); log("JCW app: " + (/BAD_LOGIN|UNKNOWN_ACTION/.test(j.error || "") ? "the server does not know 'portal_fixlist' yet. Paste the Code.gs block ABOVE  if (req.action === 'portal_known')  and the Admissions.gs function, then Deploy > Manage deployments > Edit > New version > Deploy." : (j.error || "error")), "err"); return; }
      localStorage.setItem("uopEM_key", key);
      rows = (j.rows || []).map((x, n) => ({ line: n + 1, id: String(x.id || "").replace(/\D/g, ""), enroll: x.enroll, program: x.program, cls: x.cls, section: x.section, name: x.name, father: x.father, obt: x.obt || "", tot: x.tot || "", why: x.why })).filter((x) => x.id);
      rows.forEach((x) => { delete done[x.id]; }); lsSet(LS.done, done); lsSet(LS.rows, rows);
      $("#uopEM_ids").value = ""; refreshFilters();
      if (j.note) log(j.note, "warn");
      log(rows.length ? rows.length + " student(s) with wrong values on the portal:" : "Nothing to correct: Program, Class, Section and Enroll# on the portal match the strength sheet. ✔", "ok");
      rows.slice(0, 200).forEach((x) => log("• " + x.id + " " + x.name + ": " + x.why, "dim"));
      (j.keep || []).forEach((k) => log("✋ By hand (marks etc.): " + k[0] + " " + k[1] + " — " + k[4] + ": app " + k[5] + ", portal " + k[6], "warn"));
      if (rows.length && auto) { log("Checking each student on the portal first (nothing is saved)…", "dim"); await run(true); log("Check done. Press ▶ Update to correct them on the portal.", "head"); }
    } catch (e) { log("Could not reach the JCW app: " + e.message, "err"); }
  }

  // Don't build the panel inside our own hidden frame / popup
  if (window.top !== window.self || window.name === "uopEM_editWin") return;
  buildPanel();
})();