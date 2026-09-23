// Job persistence. Lightweight job records (config, stats, status) live in
// localStorage; the potentially huge per-file `items` array is stored separately
// in IndexedDB so large jobs (tens of thousands of files) don't blow the ~5MB
// localStorage quota.
(() => {
  const KEY = 'SPOMove-jobs';
  const DB_NAME = 'SPOMove';
  const STORE = 'jobItems';
  const listeners = new Set();

  // ---------- localStorage records ----------
  function loadAll() {
    try {
      const raw = localStorage.getItem(KEY);
      return raw ? JSON.parse(raw) : [];
    } catch (e) {
      Log.warn('Could not parse stored jobs, resetting:', e.message);
      return [];
    }
  }

  function persist(jobs) {
    try {
      localStorage.setItem(KEY, JSON.stringify(jobs));
    } catch (e) {
      Log.err('Failed to persist job records:', e.message);
      throw new Error('Browser storage is full. Clear finished jobs and try again.');
    }
    listeners.forEach(fn => { try { fn(jobs); } catch {} });
  }

  // The item list is stored in IndexedDB, never in the localStorage record.
  function stripItems(job) { const clone = { ...job }; delete clone.items; return clone; }

  function list() { return loadAll().sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0)); }
  function get(id) { return loadAll().find(j => j.id === id) || null; }

  function save(job) {
    const jobs = loadAll();
    job.updatedAt = Date.now();
    const record = stripItems(job);
    const idx = jobs.findIndex(j => j.id === job.id);
    if (idx >= 0) jobs[idx] = record; else jobs.push(record);
    persist(jobs);
    return job;
  }

  function remove(id) {
    persist(loadAll().filter(j => j.id !== id));
    idbDelete(id).catch(() => {});
  }

  function clearFinished() {
    const finished = new Set(['completed', 'completed_with_errors', 'cancelled', 'failed']);
    const all = loadAll();
    for (const j of all) if (finished.has(j.status)) idbDelete(j.id).catch(() => {});
    persist(all.filter(j => !finished.has(j.status)));
  }

  function newId() {
    return (crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`);
  }

  function create(config) {
    const job = {
      id: newId(),
      createdAt: Date.now(),
      updatedAt: Date.now(),
      status: 'draft',
      config,
      stats: { total: 0, done: 0, failed: 0, skipped: 0, bytesTotal: 0, bytesDone: 0 },
      startedAt: null,
      finishedAt: null,
      lastError: null,
    };
    return save(job);
  }

  function onChange(fn) { listeners.add(fn); return () => listeners.delete(fn); }

  // ---------- IndexedDB item storage ----------
  let _dbPromise = null;
  function idb() {
    if (_dbPromise) return _dbPromise;
    _dbPromise = new Promise((resolve, reject) => {
      let req;
      try { req = indexedDB.open(DB_NAME, 1); }
      catch (e) { reject(e); return; }
      req.onupgradeneeded = () => { if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE); };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return _dbPromise;
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

  window.Jobs = { list, get, save, remove, clearFinished, create, onChange, saveItems, loadItems };
})();
