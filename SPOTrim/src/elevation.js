// Force scan: temporary site collection admin, with a ledger in localStorage so rights the
// tool added can still be removed after a crash, a closed tab or a failed clean-up.
// The ledger holds site URL, account, tenant id and a timestamp. Never tokens.
(() => {
  const KEY = 'spotrim-elevations';
  const norm = (u) => String(u || '').trim().replace(/\/+$/, '').toLowerCase();
  const lc = (s) => String(s || '').trim().toLowerCase();

  // Outcome of GET {site}/_api/web/currentuser: 'admin' | 'notAdmin' | 'unknown'.
  // Only an answer from SharePoint counts. 403 means no access at all, so not an admin.
  // 401 (token trouble), 429 after retries, 5xx and network errors prove nothing.
  function adminState(data, error) {
    if (error) return error.status === 403 ? 'notAdmin' : 'unknown';
    if (data && data.IsSiteAdmin === true) return 'admin';
    if (data && data.IsSiteAdmin === false) return 'notAdmin';
    return 'unknown';
  }

  // checks: [{ state, error }] in the same order as sites.
  function plan(sites, checks, { includeOneDrive = false, isOneDrive = () => false } = {}) {
    const out = { elevate: [], admin: [], unknown: [], oneDriveSkipped: [] };
    sites.forEach((site, i) => {
      const c = checks[i] || { state: 'unknown' };
      if (c.state === 'admin') out.admin.push(site);
      else if (c.state !== 'notAdmin') out.unknown.push({ site, error: c.error || 'Check did not complete' });
      else if (!includeOneDrive && isOneDrive(site)) out.oneDriveSkipped.push(site);
      else out.elevate.push(site);
    });
    return out;
  }

  // ---------- ledger ----------
  function sanitize(e) {
    if (!e || typeof e !== 'object') return null;
    const siteUrl = String(e.siteUrl || '');
    const account = String(e.account || '');
    if (!/^https:\/\/[^/\s]+/i.test(siteUrl) || !account) return null;
    return { siteUrl, account, tenantId: String(e.tenantId || ''), at: String(e.at || '') };
  }
  function load(store) {
    try {
      const v = JSON.parse(store.getItem(KEY) || '[]');
      return Array.isArray(v) ? v.map(sanitize).filter(Boolean) : [];
    } catch { return []; }
  }
  function save(store, list) {
    if (list.length) store.setItem(KEY, JSON.stringify(list));
    else store.removeItem(KEY);
  }
  const same = (a, b) => norm(a.siteUrl) === norm(b.siteUrl) && lc(a.account) === lc(b.account) && lc(a.tenantId) === lc(b.tenantId);
  function record(store, entry, now = new Date()) {
    const e = sanitize({ ...entry, at: now.toISOString() });
    if (!e) throw new Error('Invalid elevation record');
    const list = load(store).filter((x) => !same(x, e));
    list.push(e);
    save(store, list);
    return e;
  }
  function forget(store, entry) { save(store, load(store).filter((x) => !same(x, entry))); }
  function mine(list, account, tenantId) { return list.filter((e) => lc(e.account) === lc(account) && lc(e.tenantId) === lc(tenantId)); }
  function storageWorks(store) {
    try { store.setItem(`${KEY}-probe`, '1'); store.removeItem(`${KEY}-probe`); return true; } catch { return false; }
  }

  // A grant that got an answer from the server failed for sure. A network error or abort
  // may still have reached SharePoint, so that site stays on the clean-up list.
  function definiteFailure(e) { return !!e && (typeof e.status === 'number' || e.csom === true); }

  // ---------- runs ----------
  // currentUser(siteUrl) resolves with the currentuser JSON or rejects with the HTTP error.
  async function check(sites, { currentUser, concurrency = 4, cancelToken, onProgress } = {}) {
    const { results } = await Concurrency.pmap(sites, async (site) => {
      if (cancelToken && cancelToken.cancelled) return { state: 'unknown', error: 'Cancelled' };
      try { return { state: adminState(await currentUser(site.webUrl)) }; }
      catch (e) { return { state: adminState(null, e), error: e.message }; }
    }, { concurrency, onProgress });
    return results.map((r) => (r && r.state ? r : { state: 'unknown', error: 'Check did not complete' }));
  }

  // Each site is written to the ledger before SetSiteAdmin is sent, and dropped again only
  // when SharePoint clearly refused. Returns what has to be revoked later.
  async function grant(sites, { setAdmin, store, account, tenantId, concurrency = 4, cancelToken, onProgress } = {}) {
    const toRevoke = [];
    const failed = [];
    let skipped = 0;
    await Concurrency.pmap(sites, async (site) => {
      if (cancelToken && cancelToken.cancelled) { skipped++; return; }
      const entry = record(store, { siteUrl: site.webUrl, account, tenantId });
      try {
        await setAdmin(site.webUrl, true, account);
        toRevoke.push(entry);
      } catch (e) {
        if (definiteFailure(e)) forget(store, entry);
        else toRevoke.push(entry);
        failed.push({ site, error: e });
      }
    }, { concurrency, onProgress });
    return { toRevoke, failed, skipped };
  }

  async function revoke(entries, { setAdmin, store, concurrency = 4, onProgress } = {}) {
    const removed = [];
    const failed = [];
    await Concurrency.pmap(entries, async (e) => {
      try {
        await setAdmin(e.siteUrl, false, e.account);
        forget(store, e);
        removed.push(e);
      } catch (err) {
        failed.push({ entry: e, error: err });
      }
    }, { concurrency, onProgress });
    return { removed, failed };
  }

  window.Elevation = { KEY, adminState, plan, load, record, forget, mine, storageWorks, definiteFailure, check, grant, revoke };
})();
