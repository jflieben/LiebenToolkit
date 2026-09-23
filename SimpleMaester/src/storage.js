// Per-browser scan history using IndexedDB (falls back to localStorage if IDB blocked).
// Each scan record looks like:
//   { id, tenantId, tenantDomain, startedAt, durationMs, results: [...], summary: {...} }
//
// Storage management:
//   • IndexedDB: no hard browser quota for IDB alone, but the storage API cap is shared
//     across origins (~10% of disk or a browser-set limit). We use the StorageManager API
//     (navigator.storage.estimate) to track usage and warn / auto-prune.
//   • localStorage fallback: hard 5 MB cap; we cap at 50 scans and prune on every write.
//   • Auto-prune threshold: when usage > 80% of quota, oldest scans are deleted until
//     usage drops below 60% (keeping at least 1 scan always).
//   • A storage warning is emitted on window ('SimpleMaester:storageWarning') so the UI
//     can react without tight coupling.
(() => {
  const DB_NAME  = 'SimpleMaesterDB';
  const DB_VER   = 1;
  const STORE    = 'scans';
  const PRUNE_HIGH_PCT = 0.80;  // start pruning at 80% usage
  const PRUNE_LOW_PCT  = 0.60;  // stop pruning once below 60%
  const LS_MAX_SCANS   = 50;
  let dbPromise = null;

  function openDB() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise((resolve, reject) => {
      if (!('indexedDB' in window)) return reject(new Error('No IndexedDB'));
      const req = indexedDB.open(DB_NAME, DB_VER);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          const os = db.createObjectStore(STORE, { keyPath: 'id' });
          os.createIndex('tenantId', 'tenantId', { unique: false });
          os.createIndex('startedAt', 'startedAt', { unique: false });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    }).catch(e => { Log.warn('IndexedDB unavailable, falling back to localStorage:', e.message); return null; });
    return dbPromise;
  }

  function _lsAll() {
    try { return JSON.parse(localStorage.getItem('SimpleMaester:scans') || '[]'); } catch { return []; }
  }
  function _lsSet(arr) {
    try {
      localStorage.setItem('SimpleMaester:scans', JSON.stringify(arr));
    } catch (e) {
      // QuotaExceededError — force-prune half the scans and retry once
      Log.warn('localStorage quota hit, pruning oldest half of scans');
      const pruned = arr.slice(Math.ceil(arr.length / 2));
      try { localStorage.setItem('SimpleMaester:scans', JSON.stringify(pruned)); }
      catch (e2) { Log.err('localStorage save failed even after pruning:', e2.message); }
      window.dispatchEvent(new CustomEvent('SimpleMaester:storageWarning', {
        detail: { reason: 'quota', message: 'localStorage almost full — older scans were pruned automatically.' },
      }));
    }
  }

  // ── Storage statistics ──────────────────────────────────────────────────────

  // Returns { usedBytes, quotaBytes, usedPct, scanCount, backend } or null if unavailable.
  async function getStorageStats() {
    const db = await openDB();
    const scanCount = db ? await _idbCount(db) : _lsAll().length;
    const backend = db ? 'IndexedDB' : 'localStorage';

    if (navigator.storage && navigator.storage.estimate) {
      try {
        const est = await navigator.storage.estimate();
        const usedBytes = est.usage || 0;
        const quotaBytes = est.quota || 0;
        const usedPct = quotaBytes > 0 ? usedBytes / quotaBytes : 0;
        return { usedBytes, quotaBytes, usedPct, scanCount, backend };
      } catch (e) {
        Log.warn('storage.estimate() failed:', e.message);
      }
    }
    // Fallback for browsers without StorageManager: estimate localStorage size
    if (!db) {
      try {
        const raw = localStorage.getItem('SimpleMaester:scans') || '';
        const usedBytes = new Blob([raw]).size;
        const quotaBytes = 5 * 1024 * 1024; // 5 MB typical limit
        return { usedBytes, quotaBytes, usedPct: usedBytes / quotaBytes, scanCount, backend };
      } catch {}
    }
    return { usedBytes: 0, quotaBytes: 0, usedPct: 0, scanCount, backend };
  }

  function _idbCount(db) {
    return new Promise((resolve) => {
      try {
        const tx = db.transaction(STORE, 'readonly');
        const req = tx.objectStore(STORE).count();
        req.onsuccess = () => resolve(req.result || 0);
        req.onerror = () => resolve(0);
      } catch { resolve(0); }
    });
  }

  // ── Auto-prune ──────────────────────────────────────────────────────────────

  // Deletes oldest scans until storage usage drops below PRUNE_LOW_PCT.
  // Keeps at least 1 scan. Returns number of scans pruned.
  async function pruneOldScans() {
    const db = await openDB();
    if (db) return _idbPrune(db);
    return _lsPrune();
  }

  async function _idbPrune(db) {
    if (!navigator.storage || !navigator.storage.estimate) return 0;
    let pruned = 0;
    while (true) {
      const est = await navigator.storage.estimate().catch(() => null);
      if (!est || !est.quota) break;
      const pct = (est.usage || 0) / est.quota;
      if (pct < PRUNE_LOW_PCT) break;
      // Get the oldest scan by startedAt
      const oldest = await _idbOldest(db);
      if (!oldest) break;
      // Always keep at least 1 scan
      const count = await _idbCount(db);
      if (count <= 1) break;
      await _idbDelete(db, oldest.id);
      pruned++;
    }
    return pruned;
  }

  function _idbOldest(db) {
    return new Promise((resolve) => {
      try {
        const tx = db.transaction(STORE, 'readonly');
        const req = tx.objectStore(STORE).index('startedAt').openCursor(null, 'next');
        req.onsuccess = () => resolve(req.result ? req.result.value : null);
        req.onerror = () => resolve(null);
      } catch { resolve(null); }
    });
  }

  function _idbDelete(db, id) {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(id);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  }

  function _lsPrune() {
    const all = _lsAll();
    if (all.length <= 1) return 0;
    all.sort((a, b) => (a.startedAt || '').localeCompare(b.startedAt || ''));
    const keep = Math.max(1, Math.ceil(all.length / 2));
    const pruned = all.length - keep;
    _lsSet(all.slice(pruned));
    return pruned;
  }

  // ── Check-and-warn after each save ─────────────────────────────────────────

  async function _checkQuotaAfterSave() {
    if (!navigator.storage || !navigator.storage.estimate) return;
    try {
      const est = await navigator.storage.estimate();
      if (!est.quota) return;
      const pct = (est.usage || 0) / est.quota;
      if (pct >= PRUNE_HIGH_PCT) {
        Log.warn(`Storage at ${Math.round(pct * 100)}% — auto-pruning old scans`);
        const pruned = await pruneOldScans();
        window.dispatchEvent(new CustomEvent('SimpleMaester:storageWarning', {
          detail: {
            reason: 'auto-prune',
            message: `Storage was at ${Math.round(pct * 100)}% — ${pruned} old scan(s) were removed automatically.`,
            pct: Math.round(pct * 100),
          },
        }));
      } else if (pct >= 0.70) {
        window.dispatchEvent(new CustomEvent('SimpleMaester:storageWarning', {
          detail: {
            reason: 'warning',
            message: `Storage is at ${Math.round(pct * 100)}% — consider clearing old scans from the History tab.`,
            pct: Math.round(pct * 100),
          },
        }));
      }
    } catch (e) {
      Log.warn('Post-save quota check failed:', e.message);
    }
  }

  // ── Public CRUD ─────────────────────────────────────────────────────────────

  async function saveScan(scan) {
    // Strip functions / non-cloneable refs (e.g. test.run) so structured-clone won't choke.
    let safe;
    try { safe = JSON.parse(JSON.stringify(scan)); }
    catch (e) { Log.err('saveScan: could not serialise scan:', e.message); return; }
    const db = await openDB();
    if (!db) {
      const all = _lsAll();
      all.push(safe);
      // Cap localStorage at LS_MAX_SCANS to avoid blowing the quota.
      if (all.length > LS_MAX_SCANS) all.splice(0, all.length - LS_MAX_SCANS);
      _lsSet(all);
      _checkQuotaAfterSave();
      return;
    }
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(safe);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
    _checkQuotaAfterSave();
  }

  async function getAllScans() {
    const db = await openDB();
    if (!db) return _lsAll();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  }

  async function getScan(id) {
    const all = await getAllScans();
    return all.find(s => s.id === id);
  }

  async function deleteScan(id) {
    const db = await openDB();
    if (!db) {
      _lsSet(_lsAll().filter(s => s.id !== id));
      return;
    }
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(id);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  }

  async function clearAll() {
    const db = await openDB();
    if (!db) { localStorage.removeItem('SimpleMaester:scans'); return; }
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).clear();
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  }

  window.History = { saveScan, getAllScans, getScan, deleteScan, clearAll, getStorageStats, pruneOldScans };
})();

