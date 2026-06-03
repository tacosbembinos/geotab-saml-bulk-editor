/* global geotab */
/*
 * SAML Bulk Editor — main.js
 *
 * Patterns inherited from the Fuel Transactions Bulk Editor:
 *  - shim window.geotab before assignment (allows file:// preview)
 *  - callback() invoked unconditionally in initialize
 *  - blur() aborts in-flight + clears timers; unload() full teardown
 *  - all bulk writes via api.multiCall, chunked at MULTICALL_CHUNK = 50,
 *    SEQUENTIAL chunks throttled to the documented per-(method,entity) bucket
 *  - escapeHtml() before any innerHTML; textContent preferred otherwise
 *  - whitelist enums (userAuthenticationType) before composing entities
 *  - optimistic concurrency: keep `version` from Get; if Set fails re-Get
 *  - modal a11y: toggle [hidden] + [inert] + aria-hidden; restore focus on close
 *
 * What this add-in mutates:
 *  - User.userAuthenticationType  ("BasicAuthentication" | "SAML")
 *  - User.issuerCertificate       ({ id, isRoot: false })   (SAML only)
 *
 * Per https://support.geotab.com/mygeotab/doc/sso-saml (Additional Options for
 * Enabling SAML Authentication for Users): switching a user to SAML requires
 * setting BOTH fields; reverting clears issuerCertificate.
 */
(function () {
  'use strict';
  if (typeof window !== 'undefined' && typeof window.geotab === 'undefined') {
    window.geotab = { addin: {} };
  }
})();

geotab.addin.samlBulkEditor = function () {
  'use strict';

  // ── Constants ──────────────────────────────────────────────────────────
  // Chunk size below the tightest documented Set/User bucket so a single
  // batch fits with headroom. Larger chunks just trade a round-trip for a
  // 60-second stall when the rate window hits.
  const MULTICALL_CHUNK = 50;
  const RESULTS_LIMIT   = 50000;

  // Documented per-(method, entity, user, db) quotas in the Active service
  // tier. Anything not listed falls back to DEFAULT_LIMIT_PER_MIN. The
  // server's error message ("Maximum admitted N per Ws") narrows the bucket
  // live if our guess is too generous.
  const METHOD_LIMITS = {
    'Get:User':            60,
    'Set:User':            60,
    'Get:Certificate':     60,
    'Get:SystemSettings':  60,
    'ExecuteMultiCall':    1000
  };
  const DEFAULT_LIMIT_PER_MIN     = 60;
  const RATE_WINDOW_MS            = 60 * 1000;
  const RATE_MAX_RETRIES          = 5;
  const RATE_DEFAULT_COOLDOWN_MS  = 60 * 1000;

  const AUTH_TYPES = Object.freeze(['BasicAuthentication', 'SAML']);

  // The User entity returned by Get does NOT include the password (server
  // never serialises it). When switching a user back to BasicAuthentication
  // via Set, the server requires entity.password to be non-null or it
  // rejects with `ArgumentNullException: Value cannot be null. (Parameter
  // 'Password')`. We capture a temporary password in the bulk / row modals
  // (and stash it on the patch as __password) and apply it during commit.
  // It never goes over the wire except as part of the legitimate Set call.
  function generateTempPassword() {
    // 16 chars, mixed-case + digits + symbol. Avoids visually-ambiguous
    // glyphs (0/O, 1/l/I) so an admin who reads it aloud doesn't fumble.
    const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
    const lower = 'abcdefghijkmnopqrstuvwxyz';
    const digit = '23456789';
    const sym   = '!@#$%&*?';
    const all = upper + lower + digit + sym;
    const rnd = (set) => set[Math.floor(Math.random() * set.length)];
    let pw = rnd(upper) + rnd(lower) + rnd(digit) + rnd(sym);
    for (let i = pw.length; i < 16; i++) pw += rnd(all);
    // Fisher–Yates shuffle so the guaranteed-class chars aren't always at the front.
    const arr = pw.split('');
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      const t = arr[i]; arr[i] = arr[j]; arr[j] = t;
    }
    return arr.join('');
  }

  // ── State ──────────────────────────────────────────────────────────────
  let api = null;
  let state = null;
  const ui = {
    initialized: false,
    inflight: new Set(),       // AbortController-like { abort() } shims
    opGen: 0,                  // bumped on blur/unload; stale callbacks bail
    users: [],                 // canonical User entities (raw shape from Get)
    certs: [],                 // SAML certificates available for assignment
    certById: new Map(),
    edited: new Map(),         // id -> patch { userAuthenticationType?, issuerCertificate? }
    selected: new Set(),
    sortKey: 'name',
    sortDir: 'asc',
    activeEdit: null,          // { id, field } currently editing inline, or null
    suppressBlurCommit: false, // set briefly when Esc cancels
    lastFocusEl: null,         // focus restore on modal close
    // Quick filter mode, layered on top of the toolbar's search/auth/cert
    // filters. Lets the user one-click "show me what the pill is counting"
    // (or the action bar) when the regular filters have hidden those rows.
    //   '' = no quick filter
    //   'pending'  = only rows in ui.edited
    //   'selected' = only rows in ui.selected
    quickFilter: ''
  };

  // Cancellation: blur bumps ui.opGen + aborts every handle. apiCall rejects
  // with CANCELLED; apiMultiCall short-circuits between chunks. Handlers
  // capture `const myGen = ui.opGen` at entry and short-circuit if stale.
  const CANCELLED = Object.freeze({ __cancelled: true });
  function isCancelled(e) { return !!(e && e.__cancelled); }
  function isStale(gen) { return gen !== ui.opGen; }

  // ── DOM helpers ────────────────────────────────────────────────────────
  const $ = (id) => document.getElementById(id);
  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }
  function setStatus(msg, kind) {
    const el = $('sbe-status');
    if (!el) return;
    el.textContent = msg || '';
    el.classList.remove('is-error', 'is-success');
    if (kind === 'error')   el.classList.add('is-error');
    if (kind === 'success') el.classList.add('is-success');
  }
  function showToast(opts) {
    const host = $('sbe-toast-host');
    if (!host) return;
    const el = document.createElement('div');
    el.className = 'sbe-toast' + (opts.kind ? ' is-' + opts.kind : '');
    el.textContent = opts.message || '';
    host.appendChild(el);
    setTimeout(() => { el.remove(); }, opts.durationMs || 4000);
  }

  // ── Rate-limit infra (per-bucket sliding window) ───────────────────────
  // Buckets keyed by "Method:TypeName". Lazy-created. Single-flight gate
  // serialises all api.{call,multiCall} so the ledger update is atomic with
  // the actual emission — otherwise two callers can both pass the capacity
  // check in the same tick and double-emit.
  const rateBuckets = new Map();
  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }
  let apiGate = Promise.resolve();
  function withApiGate(fn) {
    const run = apiGate.then(fn, fn);
    apiGate = run.then(() => {}, () => {});
    return run;
  }
  function bucketKey(method, typeName) {
    if (!method) return 'unknown';
    if (method === 'ExecuteMultiCall' || method === 'Authenticate') return method;
    return method + ':' + (typeName || '*');
  }
  function getBucket(key) {
    let b = rateBuckets.get(key);
    if (!b) {
      const perMin = METHOD_LIMITS[key] != null ? METHOD_LIMITS[key] : DEFAULT_LIMIT_PER_MIN;
      b = { events: [], perMin };
      rateBuckets.set(key, b);
    }
    return b;
  }
  function pruneBucket(b, now) {
    const cutoff = now - RATE_WINDOW_MS;
    while (b.events.length && b.events[0].t < cutoff) b.events.shift();
  }
  function usedInBucket(b, now) {
    pruneBucket(b, now);
    let sum = 0;
    for (const e of b.events) sum += e.n;
    return sum;
  }
  function recordSubcalls(key, n) {
    getBucket(key).events.push({ t: Date.now(), n });
  }
  async function awaitCapacity(counts, handle) {
    for (const [key, n] of counts) {
      const b = getBucket(key);
      if (n > b.perMin) {
        throw new Error('Rate-limit misconfig: requested ' + n +
          ' against "' + key + '" (cap ' + b.perMin + '/min). Reduce MULTICALL_CHUNK.');
      }
    }
    while (!(handle && handle.aborted)) {
      const now = Date.now();
      let waitMs = 0, blockingKey = null;
      for (const [key, n] of counts) {
        const b = getBucket(key);
        const used = usedInBucket(b, now);
        if (used + n <= b.perMin) continue;
        const need = used + n - b.perMin;
        let freed = 0, waitUntil = now;
        for (const e of b.events) {
          freed += e.n;
          waitUntil = e.t + RATE_WINDOW_MS;
          if (freed >= need) break;
        }
        const w = Math.max(250, waitUntil - now + 50);
        if (w > waitMs) { waitMs = w; blockingKey = key; }
      }
      if (waitMs === 0) return;
      setStatus('Rate-limit throttle: waiting ' + Math.ceil(waitMs / 1000) +
                's for ' + blockingKey + ' bucket…');
      await sleep(waitMs);
    }
  }
  const OVER_LIMIT_RX = /OverLimitException|quota exceeded/i;
  function isOverLimitError(err) {
    if (err == null) return false;
    const seen = new Set();
    function dig(node, depth) {
      if (node == null || depth > 6) return false;
      if (typeof node === 'string') return OVER_LIMIT_RX.test(node);
      if (typeof node !== 'object') return false;
      if (seen.has(node)) return false;
      seen.add(node);
      if (Array.isArray(node)) {
        for (const item of node) if (dig(item, depth + 1)) return true;
        return false;
      }
      for (const k of ['name', 'message', 'type']) {
        if (node[k] != null && OVER_LIMIT_RX.test(String(node[k]))) return true;
      }
      for (const k of ['error', 'data', 'cause']) {
        if (k in node && dig(node[k], depth + 1)) return true;
      }
      if (Array.isArray(node.errors)) {
        for (const e of node.errors) if (dig(e, depth + 1)) return true;
      }
      return false;
    }
    return dig(err, 0);
  }
  function inlineMultiCallErrors(results) {
    if (!Array.isArray(results)) return null;
    const errs = [];
    for (let i = 0; i < results.length; i++) {
      const r = results[i];
      if (r && typeof r === 'object' && r.error) errs.push({ index: i, err: r.error });
    }
    if (!errs.length) return null;
    return { errs, overLimit: errs.some((e) => isOverLimitError(e.err)), firstErr: errs[0].err };
  }
  function parseQuotaFromMessage(err, key) {
    const msg = (err && err.message) || '';
    const m = msg.match(/Maximum admitted\s+(\d+)\s+per\s+(\d+)\s*([smh])/i);
    if (!m || !key) return;
    const n = parseInt(m[1], 10);
    if (!(n > 0)) return;
    const b = getBucket(key);
    if (n < b.perMin) b.perMin = n;
  }
  function parseRetryAfterMs(err) {
    const ra = err && (err.retryAfter || (err.data && err.data.retryAfter));
    if (typeof ra === 'number' && isFinite(ra) && ra > 0) return Math.max(1000, ra * 1000);
    return RATE_DEFAULT_COOLDOWN_MS;
  }

  // ── api.call / api.multiCall wrappers ──────────────────────────────────
  function apiCall(method, params, opts) {
    opts = opts || {};
    if (!api || typeof api.call !== 'function') {
      return Promise.reject(new Error('API unavailable (standalone preview)'));
    }
    const handle = { aborted: false, abort() { this.aborted = true; } };
    ui.inflight.add(handle);
    const invoke = () => new Promise((resolve) => {
      api.call(method, params,
        (res) => resolve({ ok: true, res }),
        (err) => resolve({ ok: false, err }));
    });
    const key = bucketKey(method, params && params.typeName);
    const counts = new Map([[key, 1]]);
    return (async () => {
      let attempt = 0;
      try {
        while (true) {
          if (handle.aborted || (opts.gen != null && isStale(opts.gen))) throw CANCELLED;
          await awaitCapacity(counts, handle);
          if (handle.aborted) throw CANCELLED;
          const out = await withApiGate(async () => {
            recordSubcalls(key, 1);
            return invoke();
          });
          if (handle.aborted || (opts.gen != null && isStale(opts.gen))) throw CANCELLED;
          if (out.ok) return out.res;
          if (isOverLimitError(out.err) && attempt < RATE_MAX_RETRIES) {
            parseQuotaFromMessage(out.err, key);
            const waitMs = Math.floor(parseRetryAfterMs(out.err) * Math.pow(1.25, attempt));
            setStatus('Rate limit hit. Cooling down ' + Math.ceil(waitMs / 1000) + 's…', 'error');
            await sleep(waitMs);
            attempt++;
            continue;
          }
          throw out.err;
        }
      } finally {
        ui.inflight.delete(handle);
      }
    })();
  }

  function apiMultiCall(calls, opts) {
    opts = opts || {};
    if (!api || typeof api.multiCall !== 'function') {
      return Promise.reject(new Error('multiCall unavailable'));
    }
    const valid = calls.filter((c) => Array.isArray(c) && typeof c[0] === 'string' && c[1] && typeof c[1] === 'object');
    if (valid.length !== calls.length) {
      return Promise.reject(new Error('multiCall: malformed sub-call; refusing to send.'));
    }
    const chunks = [];
    for (let i = 0; i < valid.length; i += MULTICALL_CHUNK) {
      chunks.push(valid.slice(i, i + MULTICALL_CHUNK));
    }
    const handle = { aborted: false, abort() { this.aborted = true; } };
    ui.inflight.add(handle);
    const sendChunk = (chunk) => new Promise((resolve) => {
      api.multiCall(chunk,
        (results) => resolve({ ok: true, results: results || [] }),
        (err)     => resolve({ ok: false, err }));
    });
    const total = valid.length;
    const label = opts.label || null;
    const checkAborted = () => handle.aborted || (opts.gen != null && isStale(opts.gen));
    return (async () => {
      const all = [];
      let done = 0;
      try {
        for (let idx = 0; idx < chunks.length; idx++) {
          const chunk = chunks[idx];
          if (checkAborted()) {
            const rest = chunks.slice(idx).reduce((n, c) => n + c.length, 0);
            for (let k = 0; k < rest; k++) all.push(CANCELLED);
            return all;
          }
          const counts = new Map();
          counts.set('ExecuteMultiCall', 1);
          for (const sub of chunk) {
            const k = bucketKey(sub[0], sub[1] && sub[1].typeName);
            counts.set(k, (counts.get(k) || 0) + 1);
          }
          let attempt = 0;
          while (true) {
            await awaitCapacity(counts, handle);
            if (checkAborted()) {
              const rest = chunks.slice(idx).reduce((n, c) => n + c.length, 0);
              for (let k = 0; k < rest; k++) all.push(CANCELLED);
              return all;
            }
            if (label && chunks.length > 1) {
              setStatus(label + ' — ' + done + ' / ' + total +
                ' (chunk ' + (idx + 1) + ' of ' + chunks.length + ')…');
            }
            const out = await withApiGate(async () => {
              for (const [k, n] of counts) recordSubcalls(k, n);
              return sendChunk(chunk);
            });
            const inline = out.ok ? inlineMultiCallErrors(out.results) : null;
            const overLimit = (!out.ok && isOverLimitError(out.err)) || (inline && inline.overLimit);
            if (overLimit && attempt < RATE_MAX_RETRIES) {
              const errForMsg = out.ok ? inline.firstErr : out.err;
              let dominantKey = null, dominantCount = -1;
              for (const [k, n] of counts) {
                if (k === 'ExecuteMultiCall') continue;
                if (n > dominantCount) { dominantCount = n; dominantKey = k; }
              }
              parseQuotaFromMessage(errForMsg, dominantKey || 'ExecuteMultiCall');
              const waitMs = Math.floor(parseRetryAfterMs(errForMsg) * Math.pow(1.25, attempt));
              setStatus('Rate limit on chunk ' + (idx + 1) + '/' + chunks.length +
                '. Cooling down ' + Math.ceil(waitMs / 1000) + 's…', 'error');
              await sleep(waitMs);
              attempt++;
              continue;
            }
            if (out.ok) {
              if (inline) {
                for (let i = 0; i < out.results.length; i++) {
                  const r = out.results[i];
                  if (r && typeof r === 'object' && r.error) all.push({ __error: r.error });
                  else all.push(r);
                }
              } else {
                all.push.apply(all, out.results);
              }
              done += chunk.length;
              break;
            }
            for (let k = 0; k < chunk.length; k++) all.push({ __error: out.err });
            done += chunk.length;
            break;
          }
        }
        return all;
      } finally {
        ui.inflight.delete(handle);
      }
    })();
  }

  // ── Data loading ───────────────────────────────────────────────────────
  // Loads Users + Certificates in one multiCall. Certificates power the
  // inline-edit dropdown; the User list is the primary table.
  function loadAll() {
    setStatus('Loading users and certificates…');
    const myGen = ui.opGen;
    const activeOnly = !!$('sbe-filter-active').checked;
    const userSearch = activeOnly ? { isDriver: false, activeFrom: new Date().toISOString() } : {};
    // Geotab doesn't expose a stable search clause for "active SAML certs
    // only" — fetch everything and filter client-side. The set is small
    // (typically <20) so the cost is negligible.
    const calls = [
      ['Get', { typeName: 'User',        search: userSearch, resultsLimit: RESULTS_LIMIT }],
      ['Get', { typeName: 'Certificate', resultsLimit: 5000 }]
    ];
    return apiMultiCall(calls, { gen: myGen, label: 'Loading' })
      .then((results) => {
        if (isStale(myGen)) return;
        const usersRes = results && results[0];
        const certsRes = results && results[1];
        if (usersRes && usersRes.__cancelled) return;
        if (usersRes && usersRes.__error) {
          setStatus('Load failed: ' + errMsg(usersRes.__error), 'error');
          return;
        }
        ui.users = Array.isArray(usersRes) ? usersRes : [];
        ui.certs = Array.isArray(certsRes) ? certsRes.filter((c) => c && c.id) : [];
        ui.certById.clear();
        ui.certs.forEach((c) => ui.certById.set(c.id, c));
        ui.edited.clear();
        ui.selected.clear();
        populateCertFilter();
        render();
        setStatus(ui.users.length + ' users · ' + ui.certs.length + ' certificates loaded.', 'success');
        showToast({ kind: 'success', message: ui.users.length + ' users loaded' });
      })
      .catch((err) => {
        if (isStale(myGen) || isCancelled(err)) return;
        setStatus('Load failed: ' + errMsg(err), 'error');
      });
  }
  function errMsg(err) {
    if (!err) return 'unknown error';
    if (typeof err === 'string') return err;
    return err.message || (err.name ? err.name : JSON.stringify(err));
  }

  // ── Derived views ──────────────────────────────────────────────────────
  // User → flat row shape for the table. Applies pending patches on top of
  // server values so the user sees the post-edit state immediately.
  function rowFor(u) {
    const patch = ui.edited.get(u.id) || {};
    const authType = patch.userAuthenticationType != null ? patch.userAuthenticationType : u.userAuthenticationType;
    const issuer   = ('issuerCertificate' in patch) ? patch.issuerCertificate : u.issuerCertificate;
    const certId   = issuer && issuer.id ? issuer.id : null;
    const cert     = certId ? ui.certById.get(certId) : null;
    return {
      raw: u,
      id: u.id,
      name: u.name || '',
      firstName: u.firstName || '',
      lastName:  u.lastName  || '',
      authType:  authType || 'BasicAuthentication',
      certId:    certId,
      certName:  cert ? certLabel(cert) : (certId ? '(unknown ' + certId.slice(0, 8) + '…)' : ''),
      lastAccess: u.lastAccessDate || u.lastLogin || '',
      active:    isActive(u),
      version:   u.version
    };
  }
  // Human-readable label for a Certificate entity. Falls back to the
  // X.509 issuer DN before the Geotab object id — the opaque guid is
  // useless to the admin choosing which cert to assign.
  function certLabel(c) {
    if (!c) return '';
    return c.name || c.subject || c.issuer || c.id;
  }
  function isActive(u) {
    if (!u) return false;
    const now = Date.now();
    const from = u.activeFrom ? Date.parse(u.activeFrom) : null;
    const to   = u.activeTo   ? Date.parse(u.activeTo)   : null;
    if (from && from > now) return false;
    if (to   && to   < now) return false;
    return true;
  }
  function deriveDisplayRows() {
    const q = ($('sbe-search').value || '').trim().toLowerCase();
    const fAuth = $('sbe-filter-auth').value;
    const fCert = $('sbe-filter-cert').value;
    const activeOnly = !!$('sbe-filter-active').checked;
    let arr = ui.users.map(rowFor);
    if (activeOnly) arr = arr.filter((r) => r.active);
    if (fAuth) arr = arr.filter((r) => r.authType === fAuth);
    if (fCert === '__none__') arr = arr.filter((r) => !r.certId);
    else if (fCert)           arr = arr.filter((r) => r.certId === fCert);
    if (q) arr = arr.filter((r) =>
      [r.name, r.firstName, r.lastName].some((v) => String(v).toLowerCase().indexOf(q) !== -1));
    if (ui.quickFilter === 'pending')  arr = arr.filter((r) => ui.edited.has(r.id));
    if (ui.quickFilter === 'selected') arr = arr.filter((r) => ui.selected.has(r.id));
    const dir = ui.sortDir === 'asc' ? 1 : -1;
    const k = ui.sortKey;
    arr.sort((a, b) => {
      const av = a[k], bv = b[k];
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      return String(av).localeCompare(String(bv)) * dir;
    });
    return arr;
  }

  // ── Render ─────────────────────────────────────────────────────────────
  let virtualRows = [];
  function render() {
    virtualRows = deriveDisplayRows();
    renderTable();
    renderTiles();
    renderActionBar();
    renderSavePill();
  }
  function renderTable() {
    const tbody = $('sbe-tbody');
    if (!tbody) return;
    if (!virtualRows.length) {
      tbody.innerHTML = '<tr><td colspan="9" class="addin-empty">No users match the current filters.</td></tr>';
      return;
    }
    const html = virtualRows.map(buildRowHtml).join('');
    tbody.innerHTML = html;
    if (ui.activeEdit) focusActiveEditInput();
  }
  function buildRowHtml(d) {
    const isSel = ui.selected.has(d.id);
    const isEdited = ui.edited.has(d.id);
    const rowCls = ['sbe-row'];
    if (isSel)    rowCls.push('is-selected');
    if (isEdited) rowCls.push('is-pending');
    if (!d.active) rowCls.push('is-inactive');

    const patch = ui.edited.get(d.id) || {};
    const orig = d.raw;
    const origAuth = orig.userAuthenticationType || 'BasicAuthentication';
    const origCertId = orig.issuerCertificate && orig.issuerCertificate.id;

    const authChanged = ('userAuthenticationType' in patch) && patch.userAuthenticationType !== origAuth;
    const certChanged = ('issuerCertificate' in patch) && ((patch.issuerCertificate && patch.issuerCertificate.id) || null) !== (origCertId || null);

    const authPillCls = d.authType === 'SAML' ? 'sbe-pill--saml' : 'sbe-pill--basic';
    const authCellInner = authChanged
      ? '<span class="cell-edit" title="was: ' + escapeHtml(origAuth) + '  →  now: ' + escapeHtml(d.authType) + '">' +
          '<span class="sbe-pill ' + authPillCls + '">' + escapeHtml(d.authType) + '</span>' +
          '<span class="cell-edit__marker" aria-hidden="true">✎</span>' +
        '</span>'
      : '<span class="sbe-pill ' + authPillCls + '">' + escapeHtml(d.authType) + '</span>';

    const certCellInner = (() => {
      const display = d.authType === 'SAML'
        ? (d.certName || '(no certificate)')
        : '—';
      if (!certChanged) return escapeHtml(display);
      const wasName = origCertId ? (ui.certById.get(origCertId) ? certLabel(ui.certById.get(origCertId)) : origCertId) : '(none)';
      return '<span class="cell-edit" title="was: ' + escapeHtml(wasName) + '  →  now: ' + escapeHtml(display) + '">' +
               escapeHtml(display) +
               '<span class="cell-edit__marker" aria-hidden="true">✎</span>' +
             '</span>';
    })();

    return (
      '<tr class="' + rowCls.join(' ') + '" data-id="' + escapeHtml(d.id) + '">' +
        '<td class="addin-col-check">' +
          '<input type="checkbox" class="sbe-row-check"' + (isSel ? ' checked' : '') +
          ' aria-label="Select ' + escapeHtml(d.name) + '">' +
        '</td>' +
        '<td>' + escapeHtml(d.name) + '</td>' +
        '<td>' + escapeHtml(d.firstName) + '</td>' +
        '<td>' + escapeHtml(d.lastName) + '</td>' +
        cellTd(d, 'authType', authCellInner) +
        cellTd(d, 'certId',   certCellInner) +
        '<td>' + escapeHtml(fmtDate(d.lastAccess)) + '</td>' +
        '<td>' + (d.active
          ? '<span class="sbe-pill sbe-pill--basic">Active</span>'
          : '<span class="sbe-pill sbe-pill--inactive">Inactive</span>') + '</td>' +
        '<td class="addin-col-actions">' +
          (isEdited
            ? '<button type="button" class="addin-row-btn" data-action="revert">Revert</button> '
            : '') +
          '<button type="button" class="addin-row-btn" data-action="edit">Edit…</button>' +
        '</td>' +
      '</tr>'
    );
  }
  function fmtDate(iso) {
    if (!iso) return '';
    const t = Date.parse(iso);
    if (!isFinite(t)) return '';
    const d = new Date(t);
    return d.toLocaleString();
  }
  function renderTiles() {
    const samlCount  = ui.users.filter((u) => (u.userAuthenticationType || '') === 'SAML').length;
    const basicCount = ui.users.length - samlCount;
    $('sbe-tile-loaded').textContent   = String(ui.users.length);
    $('sbe-tile-saml').textContent     = String(samlCount);
    $('sbe-tile-basic').textContent    = String(basicCount);
    $('sbe-tile-selected').textContent = String(ui.selected.size);
    $('sbe-tile-pending').textContent  = String(ui.edited.size);
    $('sbe-tile-certs').textContent    = String(ui.certs.length);
  }
  function renderActionBar() {
    const bar = $('sbe-action-bar');
    if (!bar) return;
    if (ui.selected.size === 0) { bar.hidden = true; return; }
    bar.hidden = false;
    $('sbe-sel-count').textContent = String(ui.selected.size);
    // Of the currently-selected IDs, how many are not in the visible
    // (post-filter) table? Reveals state hidden by search / auth / cert /
    // active-only / quick filters so the user isn't surprised that the
    // count is bigger than what they can see.
    const visibleIds = new Set(virtualRows.map((r) => r.id));
    let hidden = 0;
    ui.selected.forEach((id) => { if (!visibleIds.has(id)) hidden++; });
    const note = $('sbe-sel-hidden-note');
    const review = $('sbe-sel-review');
    if (note) {
      note.hidden = hidden === 0;
      if (hidden > 0) note.textContent = ' (' + hidden + ' off-screen)';
    }
    if (review) review.hidden = hidden === 0;
  }
  function renderSavePill() {
    const pill = $('sbe-save-pill');
    if (!pill) return;
    const n = ui.edited.size;
    if (n === 0) { pill.hidden = true; return; }
    pill.hidden = false;
    // Single textContent — see HTML comment on #sbe-save-edits. Splitting
    // into multiple child nodes makes .btn's inline-flex gap visible.
    const label = $('sbe-save-edits-label');
    if (label) label.textContent = 'Commit ' + n + ' edit' + (n === 1 ? '' : 's');
    // Mirror the action-bar's hidden-count: tells the user when staged
    // edits live on rows the current filter is hiding (the v1.0.1 → 1.0.2
    // confusion: "I haven't selected anything, why does it say 8 edits?").
    const visibleIds = new Set(virtualRows.map((r) => r.id));
    let hidden = 0;
    ui.edited.forEach((_, id) => { if (!visibleIds.has(id)) hidden++; });
    const note = $('sbe-save-hidden-note');
    const review = $('sbe-review-pending');
    if (note) {
      note.hidden = hidden === 0;
      if (hidden > 0) note.textContent = '(' + hidden + ' off-screen)';
    }
    if (review) review.hidden = hidden === 0;
  }
  // Quick-filter toggle. mode is 'pending' or 'selected'. Toggling the
  // active mode clears it. Also drives the toggleable summary tiles'
  // aria-pressed so the user can see at a glance which filter is on.
  function setQuickFilter(mode) {
    ui.quickFilter = (ui.quickFilter === mode) ? '' : mode;
    const pendBtn = $('sbe-tile-pending-btn');
    const selBtn  = $('sbe-tile-selected-btn');
    if (pendBtn) pendBtn.setAttribute('aria-pressed', ui.quickFilter === 'pending'  ? 'true' : 'false');
    if (selBtn)  selBtn.setAttribute('aria-pressed', ui.quickFilter === 'selected' ? 'true' : 'false');
    render();
  }
  function populateCertFilter() {
    const sel = $('sbe-filter-cert');
    if (!sel) return;
    while (sel.options.length > 2) sel.remove(2);
    ui.certs
      .slice()
      .sort((a, b) => certLabel(a).localeCompare(certLabel(b)))
      .forEach((c) => {
        const opt = document.createElement('option');
        opt.value = c.id;
        opt.textContent = certLabel(c);
        sel.appendChild(opt);
      });
  }

  // ── Inline cell editor ─────────────────────────────────────────────────
  // Two editable fields: authType (select) and certId (select). Clicking the
  // cell swaps it for a select; Enter/blur commits, Esc reverts. Tab moves
  // between editable cells on the same row.
  const EDIT_FIELD_ORDER = Object.freeze(['authType', 'certId']);
  const EDITABLE_CELL_MAP = {
    authType: { kind: 'select', options: () => AUTH_TYPES },
    certId:   { kind: 'select', options: () => ([{ id: '', label: '(none)' }].concat(
      ui.certs
        .slice()
        .sort((a, b) => certLabel(a).localeCompare(certLabel(b)))
        .map((c) => ({ id: c.id, label: certLabel(c) }))
    )) }
  };
  function cellEditorHtml(field, cur) {
    const spec = EDITABLE_CELL_MAP[field];
    if (!spec) return '';
    const opts = spec.options();
    return '<select class="sbe-cell-input" aria-label="Edit ' + escapeHtml(field) + '">' +
      opts.map((o) => {
        const val = typeof o === 'string' ? o : o.id;
        const lbl = typeof o === 'string' ? o : o.label;
        return '<option value="' + escapeHtml(val) + '"' +
               (String(val) === String(cur == null ? '' : cur) ? ' selected' : '') + '>' +
               escapeHtml(lbl) + '</option>';
      }).join('') +
    '</select>';
  }
  function cellTd(d, field, defaultHtml) {
    if (!EDITABLE_CELL_MAP[field]) return '<td>' + defaultHtml + '</td>';
    const isEdit = ui.activeEdit && ui.activeEdit.id === d.id && ui.activeEdit.field === field;
    if (isEdit) {
      const cur = field === 'authType' ? d.authType : d.certId;
      return '<td class="sbe-cell-editing" data-edit-field="' + field + '">' +
             cellEditorHtml(field, cur) + '</td>';
    }
    return '<td data-edit-field="' + field + '">' + defaultHtml + '</td>';
  }
  function enterCellEdit(id, field) {
    if (!EDITABLE_CELL_MAP[field]) return;
    // Cert column only makes sense when the (current) auth type is SAML.
    if (field === 'certId') {
      const r = virtualRows.find((x) => x.id === id);
      if (!r || r.authType !== 'SAML') return;
    }
    // Already editing this exact cell — bail. Re-rendering would replace
    // the live <select> node mid-click and slam the native dropdown popup
    // shut as soon as the user opened it.
    if (ui.activeEdit && ui.activeEdit.id === id && ui.activeEdit.field === field) return;
    if (ui.activeEdit && (ui.activeEdit.id !== id || ui.activeEdit.field !== field)) {
      const prev = document.querySelector('.sbe-cell-editing .sbe-cell-input');
      if (prev) commitCellEdit(ui.activeEdit.id, ui.activeEdit.field, prev.value, 0);
    }
    ui.activeEdit = { id, field };
    renderTable();
    focusActiveEditInput();
  }
  function focusActiveEditInput() {
    if (!ui.activeEdit) return;
    const tr = document.querySelector('tr[data-id="' + cssEscape(ui.activeEdit.id) + '"]');
    if (!tr) return;
    const input = tr.querySelector('td.sbe-cell-editing .sbe-cell-input');
    if (!input) return;
    input.focus();
    if (input.select) { try { input.select(); } catch (_) {} }
  }
  function cssEscape(s) {
    return String(s).replace(/[^a-zA-Z0-9_-]/g, (c) => '\\' + c);
  }
  function commitCellEdit(id, field, raw, moveDir) {
    const u = ui.users.find((x) => x.id === id);
    if (!u) { ui.activeEdit = null; render(); return; }
    const patch = ui.edited.get(id) || {};
    const origAuth = u.userAuthenticationType || 'BasicAuthentication';
    const origCertId = u.issuerCertificate && u.issuerCertificate.id;
    if (field === 'authType') {
      if (AUTH_TYPES.indexOf(raw) === -1) raw = origAuth;
      // SAML→Basic via the inline editor would need a temp password (the
      // User entity Get returns has no password; Set without one fails with
      // ArgumentNullException). Route this transition through the row Edit
      // modal where we can collect it.
      if (raw === 'BasicAuthentication' && origAuth === 'SAML' && !patch.__password) {
        showToast({
          kind: 'error',
          message: 'Reverting to Basic needs a temp password — use the Edit… button or the bulk Revert to Basic action.'
        });
        ui.activeEdit = null;
        render();
        return;
      }
      if (raw === origAuth) {
        delete patch.userAuthenticationType;
        // Reverting to original auth type also drops a staged cert change
        // and any captured temp password if they became meaningless.
        if (raw === 'BasicAuthentication') {
          delete patch.issuerCertificate;
          delete patch.__password;
        }
      } else {
        patch.userAuthenticationType = raw;
        if (raw === 'BasicAuthentication') {
          // SAML→Basic: clear the certificate to be tidy server-side.
          patch.issuerCertificate = null;
        } else if (raw === 'SAML' && !patch.issuerCertificate && !origCertId) {
          // Basic→SAML with no certificate: stay in edit mode and prompt
          // the user to pick one — empty cert + SAML is invalid.
          showToast({ kind: 'error', message: 'Pick a certificate for this user.' });
        }
        // Basic→SAML drops a previously-captured temp password (no longer relevant).
        if (raw === 'SAML') delete patch.__password;
      }
    } else if (field === 'certId') {
      const newId = raw || null;
      if (newId === origCertId) {
        delete patch.issuerCertificate;
      } else {
        patch.issuerCertificate = newId ? { id: newId, isRoot: false } : null;
      }
    }
    if (Object.keys(patch).length === 0) ui.edited.delete(id);
    else ui.edited.set(id, patch);
    ui.activeEdit = null;
    if (moveDir) {
      const target = nextEditableCell(id, field, moveDir);
      if (target) ui.activeEdit = target;
    }
    render();
    if (ui.activeEdit) focusActiveEditInput();
  }
  function cancelCellEdit() {
    if (!ui.activeEdit) return;
    ui.activeEdit = null;
    ui.suppressBlurCommit = true;
    renderTable();
    setTimeout(() => { ui.suppressBlurCommit = false; }, 0);
  }
  function nextEditableCell(id, field, dir) {
    const idx = EDIT_FIELD_ORDER.indexOf(field);
    if (idx < 0) return null;
    let nextIdx = idx + dir;
    let nextId = id;
    if (nextIdx < 0 || nextIdx >= EDIT_FIELD_ORDER.length) {
      const rowIdx = virtualRows.findIndex((r) => r.id === id);
      if (rowIdx < 0) return null;
      const sib = virtualRows[rowIdx + dir];
      if (!sib) return null;
      nextId = sib.id;
      nextIdx = dir > 0 ? 0 : EDIT_FIELD_ORDER.length - 1;
    }
    return { id: nextId, field: EDIT_FIELD_ORDER[nextIdx] };
  }

  // ── Modal (bulk edit + confirm) ────────────────────────────────────────
  function openModal(title, bodyHtml, onApply) {
    const modal = $('sbe-modal');
    if (!modal) return;
    ui.lastFocusEl = document.activeElement;
    $('sbe-modal-title').textContent = title;
    $('sbe-modal-body').innerHTML = bodyHtml;
    modal.hidden = false;
    modal.removeAttribute('inert');
    modal.setAttribute('aria-hidden', 'false');
    const saveBtn = $('sbe-modal-save');
    const handler = () => {
      // onApply may return `false` to keep the modal open (e.g. validation
      // failed). Any other return value (including undefined) closes it.
      let keepOpen = false;
      try { keepOpen = onApply && onApply() === false; }
      finally {
        if (!keepOpen) {
          closeModal();
          saveBtn.removeEventListener('click', handler);
        }
      }
    };
    saveBtn.addEventListener('click', handler);
    const firstFocus = modal.querySelector('input, select, button');
    if (firstFocus) firstFocus.focus();
  }
  function closeModal() {
    const modal = $('sbe-modal');
    if (!modal) return;
    modal.hidden = true;
    modal.setAttribute('inert', '');
    modal.setAttribute('aria-hidden', 'true');
    if (ui.lastFocusEl && typeof ui.lastFocusEl.focus === 'function') {
      try { ui.lastFocusEl.focus(); } catch (_) {}
    }
  }
  function openBulkSaml() {
    const optionsHtml = ui.certs
      .slice()
      .sort((a, b) => certLabel(a).localeCompare(certLabel(b)))
      .map((c) => '<option value="' + escapeHtml(c.id) + '">' + escapeHtml(certLabel(c)) + '</option>')
      .join('');
    if (!optionsHtml) {
      showToast({ kind: 'error', message: 'No SAML certificates available. Install one in Geotab first.' });
      return;
    }
    const body =
      '<p>Set <strong>' + ui.selected.size + '</strong> selected user' +
      (ui.selected.size === 1 ? '' : 's') + ' to <strong>SAML</strong> authentication and assign:</p>' +
      '<label class="addin-field">' +
        '<span>SAML certificate</span>' +
        '<select id="sbe-bulk-cert-select">' + optionsHtml + '</select>' +
      '</label>' +
      '<p class="sbe-modal-note">Existing certificate assignments will be replaced. Edits are staged — nothing is sent to Geotab until you click <strong>Commit</strong>.</p>';
    openModal('Set SAML for ' + ui.selected.size + ' users', body, () => {
      const certId = $('sbe-bulk-cert-select').value;
      if (!certId) return;
      let n = 0;
      ui.selected.forEach((id) => {
        const u = ui.users.find((x) => x.id === id);
        if (!u) return;
        const patch = ui.edited.get(id) || {};
        const origAuth = u.userAuthenticationType || 'BasicAuthentication';
        const origCertId = u.issuerCertificate && u.issuerCertificate.id;
        if (origAuth !== 'SAML') patch.userAuthenticationType = 'SAML';
        else delete patch.userAuthenticationType;
        if (certId !== origCertId) patch.issuerCertificate = { id: certId, isRoot: false };
        else delete patch.issuerCertificate;
        if (Object.keys(patch).length) { ui.edited.set(id, patch); n++; }
        else ui.edited.delete(id);
      });
      render();
      showToast({ kind: 'success', message: 'Staged SAML change on ' + n + ' user' + (n === 1 ? '' : 's') });
    });
  }
  function openBulkBasic() {
    // Count how many selected users are actually transitioning SAML→Basic.
    // Only those need a temp password; users already on Basic do not.
    let needPw = 0;
    ui.selected.forEach((id) => {
      const u = ui.users.find((x) => x.id === id);
      if (u && (u.userAuthenticationType || 'BasicAuthentication') === 'SAML') needPw++;
    });
    const pwId = 'sbe-bulk-basic-pw';
    const initialPw = needPw > 0 ? generateTempPassword() : '';
    const body =
      '<p>Revert <strong>' + ui.selected.size + '</strong> selected user' +
      (ui.selected.size === 1 ? '' : 's') + ' to <strong>BasicAuthentication</strong> and clear their SAML certificate.</p>' +
      (needPw > 0
        ? '<div class="sbe-pw-block">' +
            '<label class="addin-field"><span>Temporary password (applied to ' + needPw +
              ' SAML→Basic user' + (needPw === 1 ? '' : 's') + ')</span>' +
              '<div class="sbe-pw-row">' +
                '<input type="text" id="' + pwId + '" value="' + escapeHtml(initialPw) + '" autocomplete="off" spellcheck="false">' +
                '<button type="button" class="btn btn--ghost btn--sm" id="' + pwId + '-regen">Regenerate</button>' +
              '</div>' +
            '</label>' +
            '<p class="sbe-modal-note">The Geotab API requires a non-null password when switching a user back to Basic. The same temp password is applied to every SAML→Basic user in this batch — share it securely, or have each user use the standard password-reset flow afterward.</p>' +
          '</div>'
        : '<p class="sbe-modal-note">No selected users are on SAML — nothing requires a password.</p>') +
      '<p class="sbe-modal-note">Edits are staged — nothing is sent to Geotab until you click <strong>Commit</strong>.</p>';
    openModal('Revert ' + ui.selected.size + ' users to Basic', body, () => {
      const pwInput = $(pwId);
      const pw = pwInput ? pwInput.value : '';
      if (needPw > 0 && (!pw || pw.length < 8)) {
        showToast({ kind: 'error', message: 'Temporary password must be at least 8 characters.' });
        return false;
      }
      let n = 0;
      ui.selected.forEach((id) => {
        const u = ui.users.find((x) => x.id === id);
        if (!u) return;
        const patch = ui.edited.get(id) || {};
        const origAuth = u.userAuthenticationType || 'BasicAuthentication';
        const origCertId = u.issuerCertificate && u.issuerCertificate.id;
        if (origAuth !== 'BasicAuthentication') {
          patch.userAuthenticationType = 'BasicAuthentication';
          patch.__password = pw;
        } else {
          delete patch.userAuthenticationType;
          delete patch.__password;
        }
        if (origCertId) patch.issuerCertificate = null;
        else delete patch.issuerCertificate;
        if (Object.keys(patch).length) { ui.edited.set(id, patch); n++; }
        else ui.edited.delete(id);
      });
      render();
      showToast({ kind: 'success', message: 'Staged Basic revert on ' + n + ' user' + (n === 1 ? '' : 's') });
    });
    // Wire the Regenerate button (added to the modal body after openModal()).
    const regen = $(pwId + '-regen');
    if (regen) regen.addEventListener('click', () => {
      const inp = $(pwId);
      if (inp) inp.value = generateTempPassword();
    });
  }
  function openRowEdit(id) {
    const u = ui.users.find((x) => x.id === id);
    if (!u) return;
    const cur = rowFor(u);
    const certOpts = '<option value="">(none)</option>' +
      ui.certs
        .slice()
        .sort((a, b) => certLabel(a).localeCompare(certLabel(b)))
        .map((c) => '<option value="' + escapeHtml(c.id) +
          (c.id === cur.certId ? '" selected>' : '">') + escapeHtml(certLabel(c)) + '</option>')
        .join('');
    const origAuth = u.userAuthenticationType || 'BasicAuthentication';
    const body =
      '<p><strong>' + escapeHtml(u.name) + '</strong></p>' +
      '<label class="addin-field"><span>Authentication type</span>' +
        '<select id="sbe-row-auth">' +
          AUTH_TYPES.map((t) => '<option value="' + t + '"' + (t === cur.authType ? ' selected' : '') + '>' + t + '</option>').join('') +
        '</select>' +
      '</label>' +
      '<label class="addin-field"><span>SAML certificate</span>' +
        '<select id="sbe-row-cert">' + certOpts + '</select>' +
      '</label>' +
      // Password field is only meaningful when transitioning SAML→Basic.
      // Always render but hide unless the auth select transitions there.
      '<div id="sbe-row-pw-block" class="sbe-pw-block"' +
        (origAuth === 'SAML' && cur.authType === 'BasicAuthentication' ? '' : ' hidden') + '>' +
        '<label class="addin-field"><span>Temporary password</span>' +
          '<div class="sbe-pw-row">' +
            '<input type="text" id="sbe-row-pw" value="' + escapeHtml(generateTempPassword()) + '" autocomplete="off" spellcheck="false">' +
            '<button type="button" class="btn btn--ghost btn--sm" id="sbe-row-pw-regen">Regenerate</button>' +
          '</div>' +
        '</label>' +
        '<p class="sbe-modal-note">The Geotab API requires a non-null password when switching a user back to Basic.</p>' +
      '</div>';
    openModal('Edit ' + (u.name || 'user'), body, () => {
      const auth = $('sbe-row-auth').value;
      const certId = $('sbe-row-cert').value || null;
      const patch = {};
      const origCertId = u.issuerCertificate && u.issuerCertificate.id;
      if (auth !== origAuth) patch.userAuthenticationType = auth;
      if (auth === 'BasicAuthentication') {
        if (origCertId) patch.issuerCertificate = null;
        if (origAuth === 'SAML') {
          const pw = ($('sbe-row-pw') || {}).value || '';
          if (!pw || pw.length < 8) {
            showToast({ kind: 'error', message: 'Temporary password must be at least 8 characters.' });
            return false;
          }
          patch.__password = pw;
        }
      } else {
        if (certId !== origCertId) patch.issuerCertificate = certId ? { id: certId, isRoot: false } : null;
      }
      if (Object.keys(patch).length === 0) ui.edited.delete(id);
      else ui.edited.set(id, patch);
      render();
    });
    // Wire post-open handlers: regen button + show/hide password block as
    // the user toggles the auth select.
    const regen = $('sbe-row-pw-regen');
    if (regen) regen.addEventListener('click', () => {
      const inp = $('sbe-row-pw'); if (inp) inp.value = generateTempPassword();
    });
    const sel = $('sbe-row-auth');
    if (sel) sel.addEventListener('change', () => {
      const block = $('sbe-row-pw-block');
      if (!block) return;
      const toBasic = sel.value === 'BasicAuthentication' && origAuth === 'SAML';
      block.hidden = !toBasic;
    });
  }

  // ── Save (commit staged edits) ─────────────────────────────────────────
  // Per https://support.geotab.com/mygeotab/doc/sso-saml — the documented
  // pattern is to Get the user, mutate the auth/cert fields, Set the whole
  // entity back. We follow that exactly: a Get-then-Set pair per edited
  // user, batched via multiCall. Sending only a partial entity to Set can
  // silently null out other fields, so we always round-trip.
  function saveEdits() {
    if (ui.edited.size === 0) return;
    const ids = Array.from(ui.edited.keys());
    // Pre-flight: every SAML→Basic edit must carry a temp password (the
    // Geotab API rejects Set User without entity.password — the password
    // field is never returned by Get). The modal flows enforce this for
    // their input paths; this is the last-line defence for any patch that
    // slipped through (e.g. a future entry point that forgot to capture).
    const missingPw = [];
    ids.forEach((id) => {
      const u = ui.users.find((x) => x.id === id);
      const patch = ui.edited.get(id) || {};
      const origAuth = (u && u.userAuthenticationType) || 'BasicAuthentication';
      const targetAuth = patch.userAuthenticationType != null ? patch.userAuthenticationType : origAuth;
      if (origAuth === 'SAML' && targetAuth === 'BasicAuthentication' && !patch.__password) {
        missingPw.push(u ? (u.name || id) : id);
      }
    });
    if (missingPw.length) {
      setStatus('Save blocked: ' + missingPw.length + ' SAML→Basic edit(s) need a temp password.', 'error');
      showToast({
        kind: 'error',
        message: missingPw.length + ' user(s) need a temp password — open Edit… or use bulk Revert to Basic.'
      });
      console.warn('[samlBulkEditor] missing temp password for:', missingPw);
      return;
    }
    if (!confirm('Commit ' + ids.length + ' edit' + (ids.length === 1 ? '' : 's') + ' to Geotab? This cannot be undone via this add-in.')) return;
    setStatus('Saving…');
    const myGen = ui.opGen;
    // Phase 1: re-Get each User to pick up the freshest `version` and full
    // entity shape — optimistic concurrency. If the user was modified
    // server-side between our Load and Save, this is where we detect it.
    const getCalls = ids.map((id) => ['Get', { typeName: 'User', search: { id } }]);
    apiMultiCall(getCalls, { gen: myGen, label: 'Refreshing users' })
      .then((getResults) => {
        if (isStale(myGen)) return;
        const setCalls = [];
        const setIndexToUserId = [];
        const fetchErrors = []; // [{id, err}] — Get failed or patch was invalid
        ids.forEach((id, i) => {
          const res = getResults[i];
          if (!res || res.__cancelled) return;
          if (res.__error || !Array.isArray(res) || !res.length) {
            fetchErrors.push({ id, err: res && res.__error ? errMsg(res.__error) : 'not found' });
            return;
          }
          const fresh = res[0];
          const patch = ui.edited.get(id) || {};
          // Whitelist the mutable fields we own; do NOT spread the whole
          // patch object — __password is UI-only and must be moved into
          // entity.password, never sent as a top-level patch key.
          if ('userAuthenticationType' in patch) {
            const v = patch.userAuthenticationType;
            if (AUTH_TYPES.indexOf(v) === -1) {
              fetchErrors.push({ id, err: 'invalid authType ' + v });
              return;
            }
            fresh.userAuthenticationType = v;
          }
          if ('issuerCertificate' in patch) {
            const cert = patch.issuerCertificate;
            if (cert && cert.id) fresh.issuerCertificate = { id: cert.id, isRoot: false };
            else fresh.issuerCertificate = null;
          }
          // Validation: SAML without a certificate is a server error and
          // the message is not helpful. Catch it client-side instead.
          if (fresh.userAuthenticationType === 'SAML' &&
              (!fresh.issuerCertificate || !fresh.issuerCertificate.id)) {
            fetchErrors.push({ id, err: 'SAML auth requires a certificate' });
            return;
          }
          // Apply the temp password when the user is switching back to (or
          // staying on) Basic and we captured one. Geotab requires the
          // password to be non-null on every Set targeting a Basic user
          // whose stored password is unset (which is true after a SAML
          // round-trip — the SAML user has no Basic password).
          if (fresh.userAuthenticationType === 'BasicAuthentication' && patch.__password) {
            fresh.password = patch.__password;
          }
          setCalls.push(['Set', { typeName: 'User', entity: fresh }]);
          setIndexToUserId.push(id);
        });
        // Mark patches whose pre-flight Get / validation failed so the user
        // sees them flagged in the pill count rather than silently dropped.
        // We keep them in ui.edited so the user can fix and retry.
        fetchErrors.forEach((e) => console.warn('[samlBulkEditor] pre-flight', e));
        if (!setCalls.length) {
          setStatus('Nothing to save: ' + fetchErrors.length + ' user(s) failed pre-flight. See console.', 'error');
          showToast({ kind: 'error', message: fetchErrors.length + ' pre-flight failure(s). Edits kept.' });
          render(); // refresh pill / tiles
          return;
        }
        // Phase 2: bulk Set. Sequential, chunked, throttled.
        apiMultiCall(setCalls, { gen: myGen, label: 'Saving edits' })
          .then((setResults) => {
            if (isStale(myGen)) return;
            const errors = [];
            const succeededIds = [];
            setResults.forEach((r, i) => {
              const id = setIndexToUserId[i];
              if (!r || r.__cancelled) return;
              if (r.__error) errors.push({ id, err: errMsg(r.__error) });
              else succeededIds.push(id);
            });
            // Strip only the patches that actually succeeded. Failed edits
            // (and pre-flight failures) stay in ui.edited so the pill keeps
            // showing them and the user can fix / retry. This is the bug
            // the v1.0.0 release had — loadAll() was always called, which
            // unconditionally cleared ui.edited and wiped failed retries.
            succeededIds.forEach((id) => ui.edited.delete(id));
            const totalErr = errors.length + fetchErrors.length;
            const okMsg = 'Committed ' + succeededIds.length + ' edit' + (succeededIds.length === 1 ? '' : 's');
            if (totalErr === 0) {
              setStatus(okMsg + '. Refreshing…', 'success');
              showToast({ kind: 'success', message: okMsg });
              // Full reload only on full success — picks up new `version`s.
              // loadAll() clears ui.edited, which is fine here (nothing left).
              loadAll();
            } else {
              setStatus(okMsg + ' · ' + totalErr + ' failed. Edits kept — see console.', 'error');
              showToast({ kind: 'error', message: totalErr + ' edit(s) failed — kept staged for retry.' });
              errors.concat(fetchErrors).forEach((e) => console.error('[samlBulkEditor] save error', e));
              // Refresh ONLY the succeeded users in-place so their server
              // version updates, without nuking the pending-edit map. The
              // surviving entries in ui.edited still point at the original
              // users[] entries, which is what we want for retry.
              if (succeededIds.length) {
                refreshUsersById(succeededIds, myGen).finally(() => { render(); });
              } else {
                render();
              }
            }
          })
          .catch((err) => {
            if (isStale(myGen) || isCancelled(err)) return;
            // Total transport failure — nothing committed. Keep all edits.
            setStatus('Save failed: ' + errMsg(err) + '. Edits kept.', 'error');
            showToast({ kind: 'error', message: 'Save failed — edits kept staged.' });
            render();
          });
      })
      .catch((err) => {
        if (isStale(myGen) || isCancelled(err)) return;
        setStatus('Pre-flight Get failed: ' + errMsg(err) + '. Edits kept.', 'error');
        showToast({ kind: 'error', message: 'Pre-flight Get failed — edits kept staged.' });
        render();
      });
  }

  // Refresh a specific subset of User entities in place. Used after a
  // partial-success save so we pick up the new server `version` for the
  // users we successfully wrote, without disturbing the rest of ui.users
  // or ui.edited.
  function refreshUsersById(ids, gen) {
    if (!ids || !ids.length) return Promise.resolve();
    const calls = ids.map((id) => ['Get', { typeName: 'User', search: { id } }]);
    return apiMultiCall(calls, { gen, label: 'Refreshing saved users' })
      .then((results) => {
        if (isStale(gen)) return;
        ids.forEach((id, i) => {
          const r = results[i];
          if (!r || r.__cancelled || r.__error || !Array.isArray(r) || !r.length) return;
          const fresh = r[0];
          const idx = ui.users.findIndex((u) => u.id === id);
          if (idx >= 0) ui.users[idx] = fresh;
        });
      })
      .catch((err) => {
        if (isStale(gen) || isCancelled(err)) return;
        console.warn('[samlBulkEditor] post-save refresh failed', err);
      });
  }

  // ── CSV export ─────────────────────────────────────────────────────────
  function csvEscape(v) {
    if (v == null) return '';
    const s = String(v);
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }
  function exportCsv(rowsOnly) {
    const rows = rowsOnly || virtualRows;
    if (!rows.length) { showToast({ kind: 'error', message: 'Nothing to export.' }); return; }
    const headers = ['id', 'name', 'firstName', 'lastName', 'authType', 'certificateId', 'certificateName', 'lastAccess', 'active'];
    const lines = [headers.join(',')];
    rows.forEach((r) => {
      lines.push([r.id, r.name, r.firstName, r.lastName, r.authType, r.certId || '', r.certName || '', r.lastAccess || '', r.active ? 'true' : 'false']
        .map(csvEscape).join(','));
    });
    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'saml-users-' + new Date().toISOString().replace(/[:.]/g, '-') + '.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  // ── Event wiring ───────────────────────────────────────────────────────
  function bindControls() {
    $('sbe-load').addEventListener('click', loadAll);
    $('sbe-refresh').addEventListener('click', loadAll);
    $('sbe-export').addEventListener('click', () => exportCsv());
    $('sbe-export-selected').addEventListener('click', () => {
      exportCsv(virtualRows.filter((r) => ui.selected.has(r.id)));
    });
    ['sbe-search', 'sbe-filter-auth', 'sbe-filter-cert', 'sbe-filter-active'].forEach((id) => {
      const el = $(id);
      if (!el) return;
      el.addEventListener('input', render);
      el.addEventListener('change', render);
    });
    $('sbe-check-all').addEventListener('change', (e) => {
      if (e.target.checked) virtualRows.forEach((r) => ui.selected.add(r.id));
      else ui.selected.clear();
      render();
    });
    $('sbe-sel-clear').addEventListener('click', () => { ui.selected.clear(); render(); });
    // Tile / pill / action-bar quick-filter shortcuts. Toggles the matching
    // filter so the user can immediately see which rows the count refers
    // to. Tiles toggle their own mode; the pill / action-bar review
    // buttons always activate (clear-then-set) to avoid a confusing
    // "click Review and nothing changes because it was already on".
    const pendBtn = $('sbe-tile-pending-btn');
    if (pendBtn) pendBtn.addEventListener('click', () => setQuickFilter('pending'));
    const selBtn = $('sbe-tile-selected-btn');
    if (selBtn) selBtn.addEventListener('click', () => setQuickFilter('selected'));
    $('sbe-review-pending').addEventListener('click', () => {
      ui.quickFilter = ''; setQuickFilter('pending');
    });
    $('sbe-sel-review').addEventListener('click', () => {
      ui.quickFilter = ''; setQuickFilter('selected');
    });
    $('sbe-bulk-saml').addEventListener('click', openBulkSaml);
    $('sbe-bulk-basic').addEventListener('click', openBulkBasic);
    $('sbe-save-edits').addEventListener('click', saveEdits);
    $('sbe-discard-edits').addEventListener('click', () => {
      if (!confirm('Discard ' + ui.edited.size + ' staged edit(s)?')) return;
      ui.edited.clear();
      render();
    });

    // Modal close
    document.querySelectorAll('[data-modal-close]').forEach((el) =>
      el.addEventListener('click', closeModal));
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        if (ui.activeEdit) { e.preventDefault(); cancelCellEdit(); return; }
        const modal = $('sbe-modal');
        if (modal && !modal.hidden) { e.preventDefault(); closeModal(); }
      }
    });

    // Table delegation: row check, edit/revert buttons, cell-edit click, sort
    const tbody = $('sbe-tbody');
    tbody.addEventListener('click', (e) => {
      const tr = e.target.closest('tr[data-id]');
      if (!tr) return;
      const id = tr.dataset.id;
      if (e.target.classList && e.target.classList.contains('sbe-row-check')) {
        if (e.target.checked) ui.selected.add(id); else ui.selected.delete(id);
        render();
        return;
      }
      const btn = e.target.closest('[data-action]');
      if (btn) {
        const action = btn.dataset.action;
        if (action === 'edit')   { openRowEdit(id); return; }
        if (action === 'revert') { ui.edited.delete(id); render(); return; }
      }
      const td = e.target.closest('td[data-edit-field]');
      if (td) {
        const field = td.dataset.editField;
        if (EDITABLE_CELL_MAP[field]) enterCellEdit(id, field);
      }
    });
    tbody.addEventListener('change', (e) => {
      if (e.target.classList && e.target.classList.contains('sbe-cell-input') && ui.activeEdit) {
        commitCellEdit(ui.activeEdit.id, ui.activeEdit.field, e.target.value, 0);
      }
    });
    tbody.addEventListener('keydown', (e) => {
      if (!ui.activeEdit) return;
      if (!(e.target.classList && e.target.classList.contains('sbe-cell-input'))) return;
      if (e.key === 'Enter')  { e.preventDefault(); commitCellEdit(ui.activeEdit.id, ui.activeEdit.field, e.target.value, 0); }
      else if (e.key === 'Escape') { e.preventDefault(); cancelCellEdit(); }
      else if (e.key === 'Tab')    { e.preventDefault(); commitCellEdit(ui.activeEdit.id, ui.activeEdit.field, e.target.value, e.shiftKey ? -1 : 1); }
    });
    tbody.addEventListener('focusout', (e) => {
      if (ui.suppressBlurCommit) return;
      if (!ui.activeEdit) return;
      if (!(e.target.classList && e.target.classList.contains('sbe-cell-input'))) return;
      // setTimeout so a click on another editable cell registers first and
      // the activeEdit transition is handled by enterCellEdit instead.
      setTimeout(() => {
        if (!ui.activeEdit) return;
        const stillFocusedInput = document.activeElement &&
          document.activeElement.classList &&
          document.activeElement.classList.contains('sbe-cell-input');
        if (stillFocusedInput) return;
        commitCellEdit(ui.activeEdit.id, ui.activeEdit.field, e.target.value, 0);
      }, 0);
    });

    // Sort
    document.querySelectorAll('#sbe-table thead th[data-sort]').forEach((th) => {
      th.addEventListener('click', () => toggleSort(th.dataset.sort));
      th.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSort(th.dataset.sort); }
      });
    });
  }
  function toggleSort(key) {
    if (ui.sortKey === key) ui.sortDir = ui.sortDir === 'asc' ? 'desc' : 'asc';
    else { ui.sortKey = key; ui.sortDir = 'asc'; }
    document.querySelectorAll('#sbe-table thead th[data-sort]').forEach((th) => {
      th.setAttribute('aria-sort', th.dataset.sort === ui.sortKey ? (ui.sortDir === 'asc' ? 'ascending' : 'descending') : 'none');
    });
    render();
  }

  // ── Public lifecycle ───────────────────────────────────────────────────
  return {
    initialize: function (freshApi, freshState, callback) {
      api = freshApi;
      state = freshState;
      try {
        if (!ui.initialized) {
          bindControls();
          ui.initialized = true;
        }
      } catch (err) {
        console.error('[samlBulkEditor] initialize failed', err);
      }
      if (typeof callback === 'function') callback();
    },
    focus: function (freshApi, freshState) {
      api = freshApi;
      state = freshState;
      // Initial load on first focus; refresh otherwise is user-triggered.
      if (!ui.users.length) loadAll();
      const statusEl = $('sbe-status');
      if (statusEl && /…$/.test(statusEl.textContent || '')) {
        setStatus('Resumed. Previous operation was cancelled when the tab lost focus.');
      }
    },
    blur: function () {
      ui.opGen++;
      ui.inflight.forEach((c) => { try { c.abort(); } catch (_) {} });
      ui.inflight.clear();
    },
    unload: function () {
      this.blur();
      ui.users = [];
      ui.certs = [];
      ui.certById.clear();
      ui.edited.clear();
      ui.selected.clear();
      const tbody = $('sbe-tbody');
      if (tbody) tbody.innerHTML = '';
      api = null;
      state = null;
      ui.initialized = false;
    }
  };
};

// ── Standalone bootstrap (file:// preview) ───────────────────────────────
(function () {
  if (typeof window === 'undefined' || typeof document === 'undefined') return;
  if (window.geotab && window.geotab.api) return;
  document.addEventListener('DOMContentLoaded', () => {
    if (window.__samlBulkEditorBootstrapped) return;
    window.__samlBulkEditorBootstrapped = true;
    const lifecycle = window.geotab.addin.samlBulkEditor();
    const stubApi = { getSession: (cb) => cb(null) };
    lifecycle.initialize(stubApi, {}, () => lifecycle.focus(stubApi, {}));
  });
})();
