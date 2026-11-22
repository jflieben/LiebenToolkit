// Job persistence. Job records (config, stats, status, site admin grants) live in
// localStorage; the per-file item list goes to IndexedDB so jobs with tens of thousands
// of files do not hit the localStorage quota. Neither ever holds a token.
(() => {
  const KEY = 'SPOMove-jobs';
  const ELEV_KEY = 'SPOMove-elevations';
  const DB_NAME = 'SPOMove';
  const STORE = 'jobItems';
  const FINISHED = new Set(['completed', 'completed_with_errors', 'cancelled', 'failed']);
  const listeners = new Set();

  function readJson(key) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : [];
    } catch (e) {
      Log.warn(`Could not read ${key}, starting empty:`, e.message);
      return [];
    }
  }
  function writeJson(key, value) {
    if (MoveRules.containsSecret(value)) throw new Error('Refusing to store job data that contains a token.');
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (e) {
      Log.err(`Failed to store ${key}:`, e.message);
      throw new Error('Browser storage is full. Clear finished jobs and try again.');
    }
  }

  // ---------- job records ----------
  function loadAll() { return readJson(KEY); }
  function persist(jobs) {
    writeJson(KEY, jobs);
    listeners.forEach((fn) => { try { fn(jobs); } catch { /* listener failed */ } });
  }

  // The item list is stored in IndexedDB, never in the record.
  function stripItems(job) { const clone = { ...job }; delete clone.items; return clone; }

  function list() { return loadAll().sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0)); }
  function get(id) { return loadAll().find((j) => j.id === id) || null; }

  function save(job) {
    const jobs = loadAll();
    job.updatedAt = Date.now();
    const record = stripItems(job);
    const idx = jobs.findIndex((j) => j.id === job.id);
    if (idx >= 0) jobs[idx] = record; else jobs.push(record);
    persist(jobs);
    return job;
  }

  // Site admin grants of a dropped job stay on record until they are removed.
  function keepElevations(jobs) {
    const left = jobs.flatMap((j) => activeElevations(j).map((e) => ({ ...e, jobId: null, fromJob: j.id })));
    if (left.length) setPending([...pendingElevations(), ...left]);
  }

  function remove(id) {
    const all = loadAll();
    keepElevations(all.filter((j) => j.id === id));
    persist(all.filter((j) => j.id !== id));
    idbDelete(id).catch(() => {});
  }

  function clearFinished() {
    const all = loadAll();
    const finished = all.filter((j) => FINISHED.has(j.status));
    keepElevations(finished);
    for (const j of finished) idbDelete(j.id).catch(() => {});
    persist(all.filter((j) => !FINISHED.has(j.status)));
  }

  function clearAll() {
    const all = loadAll();
    keepElevations(all);
    for (const j of all) idbDelete(j.id).catch(() => {});
    persist([]);
  }

  function newId() {
    return crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function create(config, { elevations = [] } = {}) {
    return save({
      id: newId(),
      createdAt: Date.now(),
      updatedAt: Date.now(),
      status: 'draft',
      config,
      elevations,
      stats: { total: 0, done: 0, failed: 0, skipped: 0, bytesTotal: 0, bytesDone: 0 },
      startedAt: null,
      finishedAt: null,
      lastError: null,
    });
  }

  function onChange(fn) { listeners.add(fn); return () => listeners.delete(fn); }

  // ---------- site admin grants made by SPOMove ----------
  // Entry: { id, siteUrl, upn, tenantId, tenantHost, tenantName, role, grantedAt, revokedAt }
  function activeElevations(job) { return ((job && job.elevations) || []).filter((e) => !e.revokedAt); }
  function pendingElevations() { return readJson(ELEV_KEY).filter((e) => !e.revokedAt); }
  function setPending(list) { writeJson(ELEV_KEY, list.filter((e) => !e.revokedAt)); }

  function addPendingElevation(entry) {
    const e = { id: newId(), grantedAt: Date.now(), revokedAt: null, ...entry };
    setPending([...pendingElevations(), e]);
    return e;
  }
  function dropPendingElevation(id) { setPending(pendingElevations().filter((e) => e.id !== id)); }

  // Move the pending grants for these sites onto a new job.
  function takePendingElevations(match) {
    const all = pendingElevations();
    const taken = all.filter(match);
    if (taken.length) setPending(all.filter((e) => !taken.includes(e)));
    return taken;
  }

  function markElevationRevoked(entryId, jobId) {
    if (jobId) {
      const job = get(jobId);
      if (!job) return;
      const e = (job.elevations || []).find((x) => x.id === entryId);
      if (e) { e.revokedAt = Date.now(); save(job); }
      return;
    }
    dropPendingElevation(entryId);
  }

  // ---------- IndexedDB item storage ----------
  let dbPromise = null;
  function idb() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise((resolve, reject) => {
      let req;
      try { req = indexedDB.open(DB_NAME, 1); } catch (e) { reject(e); return; }
      req.onupgradeneeded = () => { if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE); };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbPromise;
  }

  async function idbPut(key, value) {
    const db = await idb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error || new Error('IndexedDB write aborted (quota?)'));
    });
  }
  async function idbGet(key) {
    const db = await idb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const r = tx.objectStore(STORE).get(key);
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
  }
  async function idbDelete(key) {
    const db = await idb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function saveItems(id, items) { await idbPut(id, items); }
  async function loadItems(id) { return (await idbGet(id).catch(() => null)) || []; }

  window.Jobs = {
    list, get, save, remove, clearFinished, clearAll, create, onChange, saveItems, loadItems, stripItems,
    activeElevations, pendingElevations, addPendingElevation, dropPendingElevation, takePendingElevations, markElevationRevoked,
  };
})();
