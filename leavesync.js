/* ============================================================
   Approved Leaves inside the attendance app — shared by leaves.html
   and the service worker (sw.js): storage on the phone, server calls
   and sending the queue of saved leave changes.
   Everything lives in self.LV so it never clashes with sw.js's own
   kvGet / kvSet. The page also gets the old global names.
   The login is the attendance app's own login (kv 'token' holds
   {username, password}, copied from the app by leaves.html).
   ============================================================ */
(function (G) {
  const API_URL = 'https://script.google.com/macros/s/AKfycbzl-DwgASIhf-pIGYi9D3M-CGOm7d7-UKg0W-OoWAiuRBHe4mjlRgMQXIHuLm6zi--i/exec';
  const DB_NAME = 'csa-leaves', DB_STORE = 'kv';
  let _db = null;

  function idbOpen() {
    if (_db) return Promise.resolve(_db);
    return new Promise((res, rej) => {
      const r = indexedDB.open(DB_NAME, 1);
      r.onupgradeneeded = () => r.result.createObjectStore(DB_STORE);
      r.onsuccess = () => { _db = r.result; res(_db); };
      r.onerror = () => rej(r.error);
    });
  }
  async function kvGet(key) {
    const db = await idbOpen();
    return new Promise((res, rej) => {
      const q = db.transaction(DB_STORE).objectStore(DB_STORE).get(key);
      q.onsuccess = () => res(q.result);
      q.onerror = () => rej(q.error);
    });
  }
  async function kvSet(key, val) {
    const db = await idbOpen();
    return new Promise((res, rej) => {
      const t = db.transaction(DB_STORE, 'readwrite');
      t.objectStore(DB_STORE).put(val, key);
      t.oncomplete = () => res();
      t.onerror = () => rej(t.error);
    });
  }
  // Queue changes always go through this lock, so the page and the worker never overwrite each other.
  function withQueueLock(fn) {
    if (G.navigator && navigator.locks) return navigator.locks.request('csa-leave-queue', fn);
    return fn();
  }

  // action 'data' -> server action 'lv_data'. payload.token = {username, password}.
  async function apiCall(action, payload, timeoutMs) {
    payload = Object.assign({}, payload || {});
    const auth = payload.token || {};
    delete payload.token;
    if (!auth.username || !auth.password) return { ok: false, code: 'AUTH', error: 'Sign in to the attendance app first.' };
    const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    const timer = ctrl ? setTimeout(() => ctrl.abort(), timeoutMs || 60000) : null;
    try {
      const r = await fetch(API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(Object.assign({ action: 'lv_' + action, username: auth.username, password: auth.password }, payload)),
        signal: ctrl ? ctrl.signal : undefined,
        redirect: 'follow'
      });
      if (!r.ok) throw new Error('Server error ' + r.status);
      const text = await r.text();
      let j;
      try { j = JSON.parse(text); }
      catch (e) { throw new Error('The server sent a web page instead of data. Check the deployment.'); }
      if (j && j.error === 'BAD_LOGIN') return { ok: false, code: 'AUTH', error: 'Your attendance-app login was refused (password changed?). Sign in again.' };
      if (j && j.error === 'UNKNOWN_ACTION') return { ok: false, error: 'The server does not know Approved leaves yet: paste Leaves.gs, add the lv_ line to doPost and deploy a new version.' };
      return j;
    } catch (e) {
      if (e.name === 'AbortError') throw new Error('The server took too long. Will try again.');
      throw e;
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  // Send waiting changes in order. Returns {sent, failed, sentOps, authError} or {busy:true}.
  async function syncQueue() {
    const run = async () => {
      const token = await kvGet('token');
      if (!token) return { sent: 0, failed: 0, sentOps: [], noToken: true };
      let sent = 0, failed = 0;
      const sentOps = [];
      for (let round = 0; round < 25; round++) {
        const q = (await kvGet('queue')) || [];
        const batch = [];   // up to 40 changes, one call under ~3 MB (photos are big)
        let size = 0;
        for (const o of q.filter(o => !o.failed)) {
          const sz = (o.photos || []).reduce((a, p) => a + (p.data ? p.data.length : 0), 0) + 600;
          if (batch.length && (batch.length >= 40 || size + sz > 3000000)) break;
          batch.push(o); size += sz;
        }
        if (!batch.length) break;
        const res = await apiCall('sync', {
          token,
          ops: batch.map(o => ({
            id: o.id, type: o.type, enroll: o.enroll, month: o.month, leaves: o.leaves,
            leaveType: o.leaveType, datesFull: o.datesFull, datesShort: o.datesShort,
            photos: (o.photos || []).map(p => ({ data: p.data }))
          }))
        }, size > 500000 ? 180000 : 60000);
        if (!res.ok) {
          if (res.code === 'AUTH' || res.code === 'DENIED') {
            await kvSet('authError', res.error || 'Please sign in again.');
            return { sent, failed, sentOps, authError: res.error || 'Please sign in again.' };
          }
          throw new Error(res.error || 'Server error');
        }
        const byId = {};
        (res.results || []).forEach(r => { byId[r.id] = r; });
        await withQueueLock(async () => {
          const cur = (await kvGet('queue')) || [];
          const keep = [];
          cur.forEach(o => {
            const r = byId[o.id];
            if (!r) keep.push(o);
            else if (r.ok) { sent++; sentOps.push(o); }
            else { o.failed = true; o.error = r.error; failed++; keep.push(o); }
          });
          await kvSet('queue', keep);
        });
        await kvSet('lastSync', Date.now());
        if (!(res.results || []).length) break;
      }
      return { sent, failed, sentOps };
    };
    if (G.navigator && navigator.locks) {
      return navigator.locks.request('csa-leave-send', { ifAvailable: true },
        lock => lock ? run() : { busy: true, sent: 0, failed: 0, sentOps: [] });
    }
    return run();
  }

  G.LV = { kvGet, kvSet, withQueueLock, apiCall, syncQueue };
  if (typeof window !== 'undefined') Object.assign(window, G.LV);   // the page uses the plain names
})(self);
