// SPOTrim UI: discovery, analysis, force scan, results.
(() => {
  const { esc, h, toast, confirm, busy, setUser } = Toolkit;

  window.Auth = ToolkitAuth.create({ scopes: { read: ['User.Read', 'Sites.Read.All'] } });

  const $ = (id) => document.getElementById(id);
  const store = (() => { try { return window.localStorage; } catch { return null; } })();
  const plural = (n, word, many) => `${n} ${n === 1 ? word : (many || `${word}s`)}`;
  const dash = (v) => (v == null || v === '' ? '-' : v);
  const fmtDate = (d) => Sites.fmtDate(d) || '-';

  const state = {
    sites: [],
    selected: new Set(),
    results: [],
    shownResults: [],
    forceReport: null,
    cancelToken: null,
    running: false,
    elevationStarted: false,
    sortSites: { key: null, dir: 1, type: 'string' },
    sortResults: { key: null, dir: 1, type: 'string' },
    // Column filters: values to hide per column. Empty = show all.
    colFilters: { type: new Set(), advice: new Set() },
  };
  let activateTab = () => {};

  // ---------- sorting ----------
  function wireSortHeaders(tableSelector, sortState, rerender) {
    const ths = document.querySelectorAll(`${tableSelector} thead th[data-sort]`);
    ths.forEach((th) => {
      th.classList.add('sortable-col');
      th.addEventListener('click', (e) => {
        if (e.target.closest('.col-filter-btn')) return;
        const key = th.dataset.sort;
        if (sortState.key === key) sortState.dir = -sortState.dir;
        else { sortState.key = key; sortState.dir = 1; }
        sortState.type = th.dataset.sortType || 'string';
        ths.forEach((x) => x.classList.remove('sort-asc', 'sort-desc'));
        th.classList.add(sortState.dir === 1 ? 'sort-asc' : 'sort-desc');
        rerender();
      });
    });
  }
  function applySort(items, sortState, valueOf) {
    if (!sortState.key) return items;
    const { dir, type } = sortState;
    const num = (v) => (type === 'date' ? (v instanceof Date ? v.getTime() : new Date(v).getTime()) : Number(v));
    return items.slice().sort((a, b) => {
      let av = valueOf(a, sortState.key);
      let bv = valueOf(b, sortState.key);
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      if (type === 'date' || type === 'number') {
        av = num(av); bv = num(bv);
        if (isNaN(av) && isNaN(bv)) return 0;
        if (isNaN(av)) return 1;
        if (isNaN(bv)) return -1;
      } else {
        av = String(av).toLowerCase();
        bv = String(bv).toLowerCase();
      }
      return av < bv ? -dir : av > bv ? dir : 0;
    });
  }

  // ---------- auth and tokens ----------
  function refreshAuthUI() {
    const signedIn = !!Auth.getAccount();
    setUser(Auth.userInfo(onSignOut));
    $('connectCard').hidden = signedIn;
    $('discoveryCard').hidden = !signedIn;
  }
  async function onSignIn() {
    try { await Auth.signIn('read'); }
    catch (e) { Log.err('Sign-in failed:', e); toast(`Sign-in failed: ${e.message}`, 'err'); }
  }
  async function onSignOut() {
    const t = Auth.tenantInfo();
    const left = store && t ? Elevation.mine(Elevation.load(store), t.upn, t.tenantId) : [];
    if (left.length) {
      const ok = await confirm({
        title: 'Sign out with admin rights still in place?',
        body: `${plural(left.length, 'site')} still ${left.length === 1 ? 'has' : 'have'} the admin rights a force scan added for ${t.upn}. You can remove them after signing in again.`,
        items: left.map((e) => e.siteUrl),
        okLabel: 'Sign out anyway',
      });
      if (!ok) return;
    }
    state.sites = []; state.results = []; state.selected.clear();
    await Auth.signOut();
  }

  async function ensureHosts() {
    if (SharePoint.getHosts()) return SharePoint.getHosts();
    const hosts = await Graph.detectHosts();
    SharePoint.setHosts(hosts);
    Log.info(`SharePoint host ${hosts.tenant}, OneDrive host ${hosts.my}`);
    return hosts;
  }

  // Silent check first. For what is missing a dialog asks first, then one Microsoft popup
  // consents to every resource at once (Auth.consent), on a fresh click so it is not
  // blocked. Runs take tokens silently afterwards and never navigate away.
  async function ensureConsent(sets, what) {
    const missing = [];
    for (const s of sets) if (!(await Auth.hasConsent(s))) missing.push(s);
    if (!missing.length) return true;
    const ok = await confirm({
      title: 'Permission needed',
      body: `${what} needs the delegated permissions below. One Microsoft window opens to sign in or consent.`,
      items: missing.flat(),
      okLabel: 'Continue',
    });
    if (!ok) return false;
    await Auth.consent(missing, { noRedirect: true });
    return true;
  }

  // Graph first (it gives the SharePoint host), then all SharePoint resources in one popup.
  async function prepare({ spo = true, admin = false, what = 'SPOTrim' } = {}) {
    if (!(await ensureConsent([Auth.resolve('read')], 'Microsoft Graph'))) return false;
    await ensureHosts();
    const sets = [];
    if (spo) sets.push(SharePoint.readScopes());
    if (admin) sets.push(SharePoint.adminScopes());
    return sets.length ? ensureConsent(sets, what) : true;
  }

  // ---------- progress ----------
  let progressStart = 0;
  function showProgress(label, total, { cancellable = true } = {}) {
    $('progressBox').hidden = false;
    $('progressLabel').textContent = label;
    $('progressEta').textContent = '';
    $('cancelBtn').hidden = !cancellable;
    $('progressBar').classList.toggle('indeterminate', total == null);
    progressStart = Date.now();
    if (total == null) { $('progressFill').style.width = ''; $('progressCount').textContent = ''; }
    else updateProgress(0, total);
  }
  function updateProgress(done, total) {
    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
    $('progressFill').style.width = `${pct}%`;
    $('progressCount').textContent = `${done} of ${total} (${pct}%)`;
    $('progressEta').textContent = Concurrency.eta(progressStart, done, total);
  }
  function hideProgress() { $('progressBox').hidden = true; }

  function setRunning(on) {
    state.running = on;
    ['discoverBtn', 'selectAllBtn', 'selectNoneBtn', 'forceScan', 'forceScanOneDrive', 'recoverBtn', 'forgetBtn'].forEach((id) => { $(id).disabled = on; });
    if (!on) {
      $('selectAllBtn').disabled = !state.sites.length;
      $('selectNoneBtn').disabled = !state.sites.length;
      $('forceScanOneDrive').disabled = !$('forceScan').checked;
    }
    updateSelCount();
  }

  function warnUnload(e) { e.preventDefault(); e.returnValue = ''; }

  // ---------- discovery ----------
  function filteredSites() {
    const q = ($('siteFilter').value || '').toLowerCase();
    const tf = $('typeFilter').value;
    const incOd = $('includeOneDrive').checked;
    const incSp = $('includeSharePoint').checked;
    return state.sites.filter((s) => {
      const cat = Sites.categoryOf(s);
      if (cat === 'OneDrive' && !incOd) return false;
      if (cat === 'SharePoint' && !incSp) return false;
      if (tf && Sites.classifyType(s) !== tf) return false;
      if (q && !((s.displayName || '').toLowerCase().includes(q) || (s.webUrl || '').toLowerCase().includes(q))) return false;
      return true;
    });
  }
  const selectedVisible = () => filteredSites().filter((s) => state.selected.has(s.id));

  async function discover() {
    const incOd = $('includeOneDrive').checked;
    const incSp = $('includeSharePoint').checked;
    if (!incOd && !incSp) { toast('Pick SharePoint sites, OneDrive sites or both.', 'warn'); return; }
    await busy($('discoverBtn'), async () => {
      let odAllowed = incOd;
      try {
        let ready = false;
        try { ready = await prepare({ admin: incOd, what: 'Discovery' }); }
        catch (e) { if (!incOd) throw e; Log.warn(`Consent for discovery with OneDrive listing failed: ${e.message}`); }
        // Without the admin permission, discovery still works through search.
        if (!ready && incOd) {
          odAllowed = false;
          if (!(await prepare({ what: 'Discovery' }))) return;
          toast('Without the admin permission only OneDrives you can open are listed.', 'warn');
        } else if (!ready) return;
      } catch (e) {
        Log.err('Preparing discovery failed:', e);
        toast(`Could not get started: ${e.message}`, 'err');
        return;
      }

      setRunning(true);
      state.cancelToken = new Concurrency.CancelToken();
      const signal = state.cancelToken.signal;
      showProgress('Discovering sites', null);
      try {
        const collected = [];
        const seen = new Set();
        const addAll = (list) => {
          let added = 0;
          for (const s of list) {
            const k = (s.webUrl || '').toLowerCase().replace(/\/+$/, '');
            if (!k || seen.has(k)) continue;
            seen.add(k);
            collected.push(s);
            added++;
          }
          return added;
        };

        Log.info(`Discovering via SharePoint Search (SharePoint=${incSp}, OneDrive=${incOd})`);
        try {
          const found = await SharePoint.searchAllSites({
            includeOneDrive: incOd, includeSharePoint: incSp, signal,
            onProgress: (n, total) => { $('progressCount').textContent = `${n} found by search${total ? ` of about ${total}` : ''}`; },
          });
          addAll(found);
          Log.info(`Search returned ${found.length} sites.`);
        } catch (e) {
          if (e.message === 'Cancelled') throw e;
          Log.warn(`SharePoint Search discovery failed: ${e.message}`);
          toast(`Search discovery failed: ${e.message}`, 'warn');
        }

        if (incOd && odAllowed) {
          $('progressLabel').textContent = 'Listing OneDrives through the tenant admin API';
          try {
            const od = await SharePoint.listOneDriveSitesViaTenantApi({ signal });
            Log.info(`Tenant admin API returned ${od.length} OneDrives, ${addAll(od)} new.`);
          } catch (e) {
            if (e.message === 'Cancelled') throw e;
            Log.warn(`OneDrive listing failed: ${e.message}`);
            toast(`OneDrive listing failed: ${e.message}`, 'err');
          }
        }

        state.sites = collected;
        state.selected.clear();
        const od = collected.filter((s) => s.isPersonalSite).length;
        Log.info(`Discovered ${collected.length} sites (${collected.length - od} SharePoint, ${od} OneDrive).`);
        toast(`Found ${collected.length} sites: ${collected.length - od} SharePoint, ${od} OneDrive.`, 'ok');
        $('siteListWrap').hidden = false;
        renderSites();
      } catch (e) {
        if (e.message === 'Cancelled') toast('Discovery cancelled', 'warn');
        else { Log.err('Discovery failed:', e); toast(`Discovery failed: ${e.message}`, 'err'); }
      } finally {
        hideProgress();
        setRunning(false);
      }
    });
  }

  function populateTypeFilter() {
    const sel = $('typeFilter');
    const cur = sel.value;
    while (sel.options.length > 1) sel.remove(1);
    [...new Set(state.sites.map(Sites.classifyType))].sort().forEach((v) => sel.appendChild(h('option', { value: v, text: v })));
    if (cur && [...sel.options].some((o) => o.value === cur)) sel.value = cur;
  }

  function renderSites() {
    populateTypeFilter();
    const visible = filteredSites();
    // Only what is visible stays selected.
    const ids = new Set(visible.map((s) => s.id));
    for (const id of [...state.selected]) if (!ids.has(id)) state.selected.delete(id);
    const list = applySort(visible, state.sortSites, (s, k) => {
      switch (k) {
        case 'title': return s.displayName || s.name || '';
        case 'url': return s.webUrl || '';
        case 'type': return Sites.classifyType(s);
        case 'template': return (s.root && s.root['@odata.type']) || '';
        case 'created': return s.createdDateTime || null;
        default: return '';
      }
    });
    document.querySelector('#siteTable tbody').innerHTML = list.map((s) => {
      const url = s.webUrl || '';
      return `<tr>
        <td class="col-check"><input type="checkbox" data-id="${esc(s.id)}" aria-label="Select ${esc(s.displayName || url)}" ${state.selected.has(s.id) ? 'checked' : ''}/></td>
        <td>${esc(s.displayName || s.name || '')}</td>
        <td class="url-cell" title="${esc(url)}"><a href="${esc(Toolkit.safeUrl(url))}" target="_blank" rel="noopener">${esc(url)}</a></td>
        <td><span class="badge ${Sites.categoryOf(s) === 'OneDrive' ? 'od' : 'spo'}">${esc(Sites.classifyType(s))}</span></td>
        <td>${esc((s.root && s.root['@odata.type']) || '')}</td>
        <td class="nowrap">${esc(fmtDate(s.createdDateTime))}</td>
      </tr>`;
    }).join('');
    $('siteHint').textContent = `${plural(visible.length, 'site')} shown of ${state.sites.length} discovered.`;
    $('selAll').checked = list.length > 0 && list.every((s) => state.selected.has(s.id));
    updateSelCount();
  }

  function updateSelCount() {
    const n = state.sites.length ? selectedVisible().length : 0;
    $('selCount').textContent = String(n);
    $('analyzeBtn').disabled = state.running || n === 0;
  }

  // ---------- force scan ----------
  function forceScanReady(acct) {
    if (!store || !Elevation.storageWorks(store)) {
      toast('Force scan keeps a list of the rights it adds in local storage, and that is blocked in this browser.', 'err');
      return false;
    }
    if (Elevation.mine(Elevation.load(store), acct.upn, acct.tenantId).length) {
      toast('Remove the admin rights from the earlier run first, see the notice at the top.', 'warn');
      renderRecovery();
      return false;
    }
    return true;
  }

  // Checks, confirms and grants. Returns false when the run should stop.
  async function forceScanGrant(selected, { concurrency, includeOneDrive }) {
    const acct = Auth.tenantInfo();
    if (!acct || !acct.upn) { toast('Cannot tell which account is signed in.', 'err'); return false; }
    if (!forceScanReady(acct)) return false;

    showProgress('Force scan: checking where you are site collection admin', selected.length);
    const checks = await Elevation.check(selected, {
      currentUser: SharePoint.getCurrentUser, concurrency, cancelToken: state.cancelToken,
      onProgress: (d, t) => updateProgress(d, t),
    });
    if (state.cancelToken.cancelled) return false;
    const p = Elevation.plan(selected, checks, { includeOneDrive, isOneDrive: (s) => Sites.categoryOf(s) === 'OneDrive' });
    p.unknown.forEach((u) => Log.warn(`Force scan: admin status of ${u.site.webUrl} unknown, left alone: ${u.error}`));
    Log.info(`Force scan check: ${p.elevate.length} to elevate, ${p.admin.length} already admin, ${p.unknown.length} unknown, ${p.oneDriveSkipped.length} OneDrives skipped.`);
    state.forceReport = { plan: p, failed: [], added: 0 };
    hideProgress();
    if (!p.elevate.length) {
      toast('Force scan: no site needs temporary admin rights.', 'info');
      return true;
    }

    const tenant = acct.domain ? `${acct.domain} (${acct.tenantId})` : acct.tenantId;
    const body = h('div', null, [
      h('p', null, ['Account ', h('strong', { text: acct.upn }), ' in tenant ', h('strong', { text: tenant }), '.']),
      h('p', { text: `SPOTrim adds this account as site collection admin on the ${plural(p.elevate.length, 'site')} below, analyzes them and removes the account again at the end, also when you cancel. Nothing else is changed. The changes show up in the audit log.` }),
      h('ul.small', null, [
        h('li', { text: `${p.admin.length} already admin: left alone, nothing is removed there.` }),
        h('li', { text: `${p.unknown.length} could not be checked: left alone and listed after the run.` }),
        includeOneDrive ? null : h('li', { text: `${p.oneDriveSkipped.length} OneDrives where you are not admin: skipped.` }),
      ]),
      h('p.muted.small', { text: 'The list is kept in this browser until the rights are removed, so an interrupted run can be cleaned up on the next visit.' }),
    ]);
    const ok = await confirm({
      title: `Temporary site admin on ${plural(p.elevate.length, 'site')}?`,
      body, items: p.elevate.map((s) => s.webUrl), itemsMax: 50, wide: true, danger: true,
      okLabel: `Add admin on ${plural(p.elevate.length, 'site')}`,
    });
    if (!ok) { toast('Force scan cancelled, nothing changed.', 'info'); return false; }

    state.elevationStarted = true;
    window.addEventListener('beforeunload', warnUnload);
    showProgress(`Force scan: adding temporary admin on ${plural(p.elevate.length, 'site')}`, p.elevate.length);
    const g = await Elevation.grant(p.elevate, {
      setAdmin: (url, on, upn) => SharePoint.setSiteAdmin(url, upn, on),
      store, account: acct.upn, tenantId: acct.tenantId, concurrency, cancelToken: state.cancelToken,
      onProgress: (d, t) => updateProgress(d, t),
    });
    g.failed.forEach((f) => Log.warn(`Force scan: adding admin on ${f.site.webUrl} failed: ${f.error.message}`));
    state.forceReport.failed = g.failed;
    state.forceReport.added = g.toRevoke.length;
    Log.info(`Force scan: admin added on ${g.toRevoke.length}, ${g.failed.length} failed, ${g.skipped} skipped after cancel.`);
    if (g.failed.length) toast(`Force scan: adding admin failed on ${plural(g.failed.length, 'site')}, see the debug log.`, 'warn');
    return true;
  }

  async function cleanup(entries, concurrency) {
    if (!entries.length) {
      if (!Elevation.load(store).length) window.removeEventListener('beforeunload', warnUnload);
      return;
    }
    showProgress(`Removing temporary admin from ${plural(entries.length, 'site')}`, entries.length, { cancellable: false });
    const r = await Elevation.revoke(entries, {
      setAdmin: (url, on, upn) => SharePoint.setSiteAdmin(url, upn, on),
      store, concurrency, onProgress: (d, t) => updateProgress(d, t),
    });
    r.removed.forEach((e) => Log.info(`Removed temporary admin on ${e.siteUrl}`));
    r.failed.forEach((f) => Log.err(`Could not remove admin on ${f.entry.siteUrl}: ${f.error.message}`));
    if (r.failed.length) toast(`${plural(r.failed.length, 'admin right')} not removed. Use the notice on the first tab to try again.`, 'err', { timeout: 15000 });
    else toast(`Removed temporary admin from ${plural(r.removed.length, 'site')}.`, 'ok');
    if (!Elevation.load(store).length) window.removeEventListener('beforeunload', warnUnload);
    renderRecovery();
  }

  // ---------- analysis ----------
  async function analyze() {
    const selected = selectedVisible();
    if (!selected.length) return;
    const concurrency = Math.min(32, Math.max(1, parseInt($('concurrency').value, 10) || 8));
    const staleDays = Math.min(3650, Math.max(1, parseInt($('staleDays').value, 10) || 365));
    const force = $('forceScan').checked;
    const includeOneDrive = $('forceScanOneDrive').checked;

    try {
      if (!(await prepare({ admin: force, what: force ? 'Force scan' : 'Analysis' }))) return;
    } catch (e) {
      Log.err('Preparing analysis failed:', e);
      toast(`Could not get started: ${e.message}`, 'err');
      return;
    }

    setRunning(true);
    state.cancelToken = new Concurrency.CancelToken();
    state.forceReport = null;
    state.elevationStarted = false;
    const acct = Auth.tenantInfo();
    let analysed = false;
    try {
      if (force && !(await forceScanGrant(selected, { concurrency, includeOneDrive }))) return;
      if (state.cancelToken.cancelled) return;

      showProgress(`Analyzing ${plural(selected.length, 'site')}`, selected.length);
      const { results } = await Concurrency.pmap(selected, async (site) => {
        state.cancelToken.throwIfCancelled();
        try {
          return await Analysis.analyzeSite(site, { staleDays });
        } catch (e) {
          Log.err(`Analysis failed for ${site.webUrl}:`, e.message);
          return { site, signals: {}, lastModified: null, lastViewed: null, advice: 'Unknown', reason: `Analysis error: ${e.message}`, errors: [e.message] };
        }
      }, { concurrency, cancelToken: state.cancelToken, onProgress: (d, t) => updateProgress(d, t) });
      state.results = results.filter((r) => r && !r.__error);
      analysed = true;
      Log.info(`Analysis done: ${state.results.length} results${state.cancelToken.cancelled ? ' (cancelled part way)' : ''}.`);
    } catch (e) {
      Log.err('Analysis run failed:', e);
      toast(`Analysis failed: ${e.message}`, 'err');
    } finally {
      // Everything this account has in the ledger was added by this run: a run does
      // not start while older entries exist.
      if (state.elevationStarted && store && acct) {
        try { await cleanup(Elevation.mine(Elevation.load(store), acct.upn, acct.tenantId), concurrency); }
        catch (e) { Log.err('Clean-up failed:', e); toast(`Clean-up failed: ${e.message}. Use the notice on the first tab.`, 'err'); }
      }
      hideProgress();
      setRunning(false);
    }
    if (!analysed) return;
    $('exportXlsxBtn').disabled = !state.results.length;
    $('exportCsvBtn').disabled = !state.results.length;
    renderResults();
    activateTab('results');
  }

  // ---------- recovery of an interrupted force scan ----------
  function renderRecovery() {
    const list = store ? Elevation.load(store) : [];
    const box = $('recoveryNotice');
    if (!list.length) { box.hidden = true; return; }
    const t = Auth.tenantInfo();
    const mine = t ? Elevation.mine(list, t.upn, t.tenantId) : [];
    const others = [...new Set(list.filter((e) => !mine.includes(e)).map((e) => e.account))];
    const parts = [];
    if (mine.length) parts.push(`${plural(mine.length, 'site')} for ${t.upn}.`);
    if (others.length) parts.push(`${plural(list.length - mine.length, 'site')} for ${others.join(', ')}: sign in as that account to remove them.`);
    $('recoveryText').textContent = parts.join(' ');
    $('recoverBtn').hidden = !mine.length;
    box.hidden = false;
  }

  async function recover() {
    const t = Auth.tenantInfo();
    if (!t) return;
    const mine = Elevation.mine(Elevation.load(store), t.upn, t.tenantId);
    if (!mine.length) { renderRecovery(); return; }
    const ok = await confirm({
      title: `Remove admin rights on ${plural(mine.length, 'site')}?`,
      body: `A force scan added ${t.upn} as site collection admin on these sites in tenant ${t.domain || t.tenantId} and did not get to remove it. This removes the account as site collection admin again.`,
      items: mine.map((e) => `${e.siteUrl}  (added ${String(e.at).substring(0, 16).replace('T', ' ')})`),
      okLabel: `Remove from ${plural(mine.length, 'site')}`,
      danger: true,
    });
    if (!ok) return;
    await busy($('recoverBtn'), async () => {
      try {
        if (!(await prepare({ spo: false, admin: true, what: 'Removing admin rights' }))) return;
        setRunning(true);
        await cleanup(mine, 4);
      } catch (e) {
        Log.err('Removing admin rights failed:', e);
        toast(`Removing admin rights failed: ${e.message}`, 'err');
      } finally {
        hideProgress();
        setRunning(false);
        renderRecovery();
      }
    });
  }

  async function forgetLedger() {
    const list = Elevation.load(store);
    const ok = await confirm({
      title: 'Forget this list?',
      body: 'Removes the list from this browser without changing anything in SharePoint. Only do this when the rights were removed some other way.',
      items: list.map((e) => `${e.siteUrl}  (${e.account})`),
      okLabel: 'Forget',
      danger: true,
    });
    if (!ok) return;
    list.forEach((e) => Elevation.forget(store, e));
    Log.info(`Forgot ${plural(list.length, 'force scan entry', 'force scan entries')}.`);
    renderRecovery();
  }

  // ---------- results ----------
  const BADGE = { 'Clean up': 'cleanup', Review: 'review', Keep: 'keep', 'No access': 'noaccess', Unknown: 'unknown' };
  const badgeClass = (advice) => BADGE[advice] || 'unknown';
  const lastViewedLabel = (r) => (r.lastViewed ? (r.lastViewedApprox ? `≤7d (${fmtDate(r.lastViewed)})` : fmtDate(r.lastViewed)) : '-');

  function trimSiteUrlPrefix(url) {
    if (!url) return '';
    try {
      let p = new URL(url).pathname || '/';
      if (p.length > 1) p = p.replace(/\/+$/, '');
      return p || '/';
    } catch {
      return url.replace(/^https?:\/\/[^/]+/i, '') || url;
    }
  }

  function filteredResults() {
    const q = ($('resultFilter').value || '').toLowerCase();
    const list = state.results.filter((r) => {
      if (state.colFilters.type.has(Sites.classifyType(r.site))) return false;
      if (state.colFilters.advice.has(r.advice || '')) return false;
      if (q && !`${r.site.displayName || ''} ${r.site.webUrl || ''}`.toLowerCase().includes(q)) return false;
      return true;
    });
    return applySort(list, state.sortResults, (r, k) => {
      switch (k) {
        case 'title': return r.site.displayName || r.site.name || '';
        case 'url': return r.site.webUrl || '';
        case 'type': return Sites.classifyType(r.site);
        case 'owners': return r.owners && r.owners[0] ? Sites.ownerName(r.owners[0]) : null;
        case 'lastModified': return r.lastModified || null;
        case 'lastViewed': return r.lastViewed || null;
        case 'daysIdle': return r.daysSinceModified == null ? null : Number(r.daysSinceModified);
        case 'storageMB': return r.storageMB == null ? null : Number(r.storageMB);
        case 'advice': return r.advice || '';
        default: return '';
      }
    });
  }

  function ownerCell(owners) {
    if (!owners || !owners.length) return '<span class="muted">-</span>';
    const first = owners[0];
    const more = owners.length > 1 ? ` <span class="muted small">+${owners.length - 1}</span>` : '';
    return `<span class="owner-cell" title="${esc(owners.map(Sites.ownerName).join(', '))}">${esc(Sites.ownerName(first))}</span>${more}`;
  }

  function renderResults() {
    renderResultNotices();
    refreshColumnFilterIndicators();
    const list = filteredResults();
    state.shownResults = list;
    document.querySelector('#resultTable tbody').innerHTML = list.map((r, i) => {
      const url = r.site.webUrl || '';
      return `<tr class="clickable-row" data-idx="${i}">
        <td><span class="badge ${badgeClass(r.advice)}">${esc(r.advice)}</span></td>
        <td>${esc(r.site.displayName || r.site.name || '')}</td>
        <td class="url-cell" title="${esc(url)}"><a href="${esc(Toolkit.safeUrl(url))}" target="_blank" rel="noopener">${esc(trimSiteUrlPrefix(url))}</a></td>
        <td><span class="badge ${Sites.categoryOf(r.site) === 'OneDrive' ? 'od' : 'spo'}">${esc(Sites.classifyType(r.site))}</span></td>
        <td class="owner-col">${ownerCell(r.owners)}</td>
        <td class="nowrap">${esc(fmtDate(r.lastModified))}</td>
        <td class="nowrap">${esc(lastViewedLabel(r))}</td>
        <td class="num">${esc(dash(r.daysSinceModified))}</td>
        <td class="num">${esc(dash(r.storageMB))}</td>
      </tr>`;
    }).join('');
    $('resultHint').textContent = state.results.length
      ? `${plural(list.length, 'site')} shown of ${state.results.length}. Exports contain the rows shown.`
      : 'No analysis yet. Discover sites, select some and analyze them.';
    $('exportXlsxBtn').disabled = !list.length;
    $('exportCsvBtn').disabled = !list.length;
  }

  function renderResultNotices() {
    const host = $('resultNotices');
    host.replaceChildren();
    const add = (kind, iconName, text, items) => {
      const content = [h('div', { text })];
      if (items && items.length) {
        content.push(h('ul.small.notice-list', null, items.slice(0, 20).map((t) => h('li', { text: t }))));
        if (items.length > 20) content.push(h('div.small', { text: `+ ${items.length - 20} more in the debug log` }));
      }
      host.appendChild(h('div.notice', { class: `notice-${kind}` }, [Toolkit.icon(iconName), h('div', null, content)]));
    };
    const denied = state.results.filter((r) => r && r.advice === 'No access');
    if (denied.length) add('info', 'i-lock', `${plural(denied.length, 'site')} could not be read (HTTP 401/403). Filter the recommendation on "No access" to see them.`);
    const fr = state.forceReport;
    if (fr) {
      if (fr.plan.unknown.length) add('warn', 'i-alert', `Force scan: the admin check did not complete on ${plural(fr.plan.unknown.length, 'site')}. They were not elevated.`, fr.plan.unknown.map((u) => `${u.site.webUrl}: ${u.error}`));
      if (fr.failed.length) add('warn', 'i-alert', `Force scan: adding admin failed on ${plural(fr.failed.length, 'site')}.`, fr.failed.map((f) => `${f.site.webUrl}: ${f.error.message}`));
      if (fr.plan.oneDriveSkipped.length) add('info', 'i-info', `Force scan: ${plural(fr.plan.oneDriveSkipped.length, 'OneDrive')} skipped because OneDrives were not ticked.`);
    }
  }

  // ---------- column filters ----------
  function wireColumnFilterHeaders() {
    document.querySelectorAll('#resultTable thead th[data-filter]').forEach((th) => {
      th.classList.add('filterable-col');
      const btn = h('button.col-filter-btn', { type: 'button', title: 'Filter', 'aria-label': 'Filter column' }, Toolkit.icon('i-filter', 'sm'));
      btn.addEventListener('click', (e) => { e.stopPropagation(); openColumnFilter(th.dataset.filter, btn); });
      th.appendChild(btn);
    });
    document.addEventListener('click', (e) => {
      const pop = $('columnFilterPopover');
      if (!pop.hidden && !pop.contains(e.target) && !e.target.closest('.col-filter-btn')) pop.hidden = true;
    });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') $('columnFilterPopover').hidden = true; });
  }
  function refreshColumnFilterIndicators() {
    document.querySelectorAll('#resultTable thead th[data-filter]').forEach((th) => {
      th.classList.toggle('filter-active', state.colFilters[th.dataset.filter].size > 0);
    });
  }
  function openColumnFilter(colKey, anchor) {
    const pop = $('columnFilterPopover');
    const valueOf = colKey === 'type' ? (r) => Sites.classifyType(r.site) : (r) => r.advice || '';
    const counts = new Map();
    for (const r of state.results) { const v = valueOf(r) || ''; counts.set(v, (counts.get(v) || 0) + 1); }
    const values = [...counts.keys()].sort();
    const hidden = state.colFilters[colKey];
    const apply = () => {
      hidden.clear();
      pop.querySelectorAll('input[data-val]').forEach((cb) => { if (!cb.checked) hidden.add(cb.dataset.val); });
      renderResults();
    };
    const all = h('input', { type: 'checkbox', checked: values.every((v) => !hidden.has(v)) });
    all.addEventListener('change', () => { pop.querySelectorAll('input[data-val]').forEach((cb) => { cb.checked = all.checked; }); apply(); });
    pop.replaceChildren(
      h('div.popover-title', { text: `Filter ${colKey === 'type' ? 'type' : 'recommendation'}` }),
      h('label.checkbox.col-filter-all', null, [all, 'Select all']),
      h('div.col-filter-list', null, values.map((v) => {
        const cb = h('input', { type: 'checkbox', checked: !hidden.has(v), dataset: { val: v } });
        cb.addEventListener('change', apply);
        return h('label.checkbox', null, [cb, v || '(empty)', h('span.muted', { text: ` (${counts.get(v)})` })]);
      })),
    );
    pop.hidden = false;
    const r = anchor.getBoundingClientRect();
    const pr = pop.getBoundingClientRect();
    let left = r.left + window.scrollX;
    if (r.left + pr.width > window.innerWidth - 12) left = Math.max(8, window.innerWidth - pr.width - 12) + window.scrollX;
    pop.style.left = `${left}px`;
    pop.style.top = `${r.bottom + 4 + window.scrollY}px`;
  }

  // ---------- detail dialog: how the verdict was reached ----------
  function showDetail(r) {
    const url = r.site.webUrl || '';
    const owners = r.owners && r.owners.length
      ? h('ul.owners-list', null, r.owners.map((o) => h('li', null, [
        h('strong', { text: Sites.ownerName(o) }),
        o.email ? h('div.muted.small', null, h('a', { href: `mailto:${o.email}`, text: o.email })) : null,
      ])))
      : h('span.muted', { text: 'No owners found' });
    const rows = [
      ['URL', h('a', { href: url, target: '_blank', rel: 'noopener', text: url })],
      ['Type', h('span.badge', { class: Sites.categoryOf(r.site) === 'OneDrive' ? 'od' : 'spo', text: Sites.classifyType(r.site) })],
      ['Recommendation', h('span.badge', { class: badgeClass(r.advice), text: r.advice })],
      ['Why', r.reason || ''],
      ['Owners', owners],
      ['Last modified', fmtDate(r.lastModified)],
      ['Last viewed', lastViewedLabel(r)],
      ['Days since modified', dash(r.daysSinceModified)],
      ['Days since viewed', dash(r.daysSinceViewed)],
      ['Page views (all time)', dash(r.pageViews)],
      ['Views (last 7 days)', dash(r.recentViews)],
      ['Views (last 30 days)', dash(r.last30Views)],
      ['Storage (MB)', dash(r.storageMB)],
    ];
    const dl = (pairs, cls) => h(`dl.detail-list${cls ? `.${cls}` : ''}`, null, pairs.flatMap(([k, v]) => [h('dt', { text: k }), h('dd', null, v instanceof Node ? v : String(v))]));
    const signals = Object.entries(r.signals || {}).map(([k, v]) => [k, v instanceof Date ? v.toISOString() : dash(v)]);
    const body = h('div', null, [
      dl(rows),
      h('details.detail-signals', null, [h('summary', { text: 'Raw signals' }), dl(signals, 'small')]),
      r.errors && r.errors.length
        ? h('details.detail-errors', { open: true }, [h('summary', { text: 'Errors' }), h('ul', null, r.errors.map((e) => h('li', { text: e })))])
        : null,
    ]);
    Toolkit.modal({ title: r.site.displayName || r.site.name || '(untitled site)', body, wide: true });
  }

  function onResultClick(e) {
    if (e.target.closest('a')) return;
    const tr = e.target.closest('tr[data-idx]');
    const r = tr && state.shownResults[Number(tr.dataset.idx)];
    if (r) showDetail(r);
  }

  // ---------- export ----------
  const exportRows = () => state.shownResults.map(Sites.exportRow);
  const stamp = () => new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);

  // ---------- boot ----------
  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    activateTab = Toolkit.tabs('#tabs');

    $('signInBtn').addEventListener('click', onSignIn);
    $('discoverBtn').addEventListener('click', discover);
    $('analyzeBtn').addEventListener('click', analyze);
    $('cancelBtn').addEventListener('click', () => { if (state.cancelToken) { state.cancelToken.cancel(); toast('Cancelling', 'info'); } });
    $('selectAllBtn').addEventListener('click', () => { filteredSites().forEach((s) => state.selected.add(s.id)); renderSites(); });
    $('selectNoneBtn').addEventListener('click', () => { state.selected.clear(); renderSites(); });
    $('selAll').addEventListener('change', (e) => {
      filteredSites().forEach((s) => { if (e.target.checked) state.selected.add(s.id); else state.selected.delete(s.id); });
      renderSites();
    });
    document.querySelector('#siteTable tbody').addEventListener('change', (e) => {
      const cb = e.target.closest('input[data-id]');
      if (!cb) return;
      if (cb.checked) state.selected.add(cb.dataset.id); else state.selected.delete(cb.dataset.id);
      $('selAll').checked = false;
      updateSelCount();
    });
    ['siteFilter', 'typeFilter', 'includeOneDrive', 'includeSharePoint'].forEach((id) => $(id).addEventListener(id === 'siteFilter' ? 'input' : 'change', renderSites));
    $('forceScan').addEventListener('change', () => {
      $('forceScanOneDrive').disabled = !$('forceScan').checked;
      if (!$('forceScan').checked) $('forceScanOneDrive').checked = false;
    });
    wireSortHeaders('#siteTable', state.sortSites, renderSites);

    $('resultFilter').addEventListener('input', renderResults);
    $('clearFiltersBtn').addEventListener('click', () => { state.colFilters.type.clear(); state.colFilters.advice.clear(); renderResults(); });
    document.querySelector('#resultTable tbody').addEventListener('click', onResultClick);
    wireSortHeaders('#resultTable', state.sortResults, renderResults);
    wireColumnFilterHeaders();

    $('exportXlsxBtn').addEventListener('click', () => busy($('exportXlsxBtn'), () =>
      Export.xlsx([{ name: 'StaleSites', rows: exportRows(), columns: Sites.EXPORT_COLUMNS }], `SPOTrim-StaleSites-${stamp()}.xlsx`)
        .catch((e) => toast(`Export failed: ${e.message}`, 'err'))));
    $('exportCsvBtn').addEventListener('click', () => Export.csv(exportRows(), { columns: Sites.EXPORT_COLUMNS, filename: `SPOTrim-StaleSites-${stamp()}.csv` }));

    $('recoverBtn').addEventListener('click', recover);
    $('forgetBtn').addEventListener('click', forgetLedger);

    try {
      await Auth.init();
      const cfg = Auth.getAuthConfig();
      $('authNote').textContent = `Redirect URI ${Auth.getRedirectUri()}, client id ${cfg.clientId}`;
      $('authNote').hidden = false;
      refreshAuthUI();
      if (Auth.getAccount()) {
        try { await ensureHosts(); }
        catch (e) { Log.dbg(`SharePoint host not detected yet: ${e.message}`); }
      }
      Log.info('SPOTrim ready.');
    } catch (e) {
      Log.err('Start-up failed:', e);
      toast(`Start-up failed: ${e.message}`, 'err');
      $('connectCard').querySelector('div:last-child').appendChild(h('p.notice.notice-err', { text: `Could not start: ${e.message}` }));
    }
    renderRecovery();
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
