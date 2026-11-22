// SPNRoleMgr UI. Nothing changes without a confirm that names the tenant, the
// service principal and every permission involved.
(() => {
  const { esc, h, toast, busy, setUser } = Toolkit;

  window.Auth = ToolkitAuth.create({
    scopes: {
      read: ['User.Read', 'Application.Read.All', 'Directory.Read.All'],
      write: ['AppRoleAssignment.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All'],
      // Only Graph scope that can manage /sites/{id}/permissions. Asked when the site manager opens.
      sites: ['Sites.FullControl.All'],
    },
  });

  const state = {
    spns: [],
    filteredSpns: [],
    apiSpns: [],          // service principals that expose assignable permissions
    selectedSpn: null,
    appRoles: [],
    oauth2Grants: [],
    resourceCache: new Map(),       // resource object id -> service principal
    resourceIndexCache: new Map(),  // resource object id -> permission index
    addApiSpn: null,
    addApiKind: 'AppRole',
    addSelected: new Set(),
    selRoles: new Set(),    // app role assignment ids ticked for revoke
    selScopes: new Set(),   // Spns.scopeKey(grantId, scope) ticked for revoke
    sort: { col: 'name', dir: 'asc' },
    cancelToken: null,
    elevations: [],         // temporary site admin records, mirrored to localStorage
    siteManager: {
      allSites: null,       // every site collection (admin API), cached for this page
      source: null,         // 'admin' | 'search'
      sites: [],            // current filtered list
      rendered: [],         // what the list shows, sorted
      selectedSite: null,   // site collection the detail pane works on (with Graph id)
      selectedKey: null,
      sitePerms: [],
      readable: false,      // its permissions could be read
      selectedRoles: new Set(),
      selectedPerms: new Set(), // permission ids ticked for removal
      access: new Map(),    // siteKey -> { status: 'yes'|'no'|'denied'|'error', roles }
      ids: new Map(),       // siteKey -> resolved Graph site
      scan: null,           // { done, total } while checking access
      searchToken: 0,
      scanToken: 0,
      detailToken: 0,
      optimisticUnlock: false, // right after granting Sites.Selected, before Graph shows it
      returnFocus: null,
    },
  };
  const $ = (id) => document.getElementById(id);
  const MAX_ROWS = 1000;
  const plural = (n, word) => `${n} ${word}${n === 1 ? '' : 's'}`;
  const NOTICE_ICON = { ok: 'i-check-circle', warn: 'i-alert', err: 'i-error', info: 'i-info' };

  function setNotice(el, message, kind) {
    el.className = `notice notice-${kind}`;
    el.replaceChildren(Toolkit.icon(NOTICE_ICON[kind] || 'i-info'), h('div', { text: message }));
    el.hidden = false;
  }
  function clearNotice(el) { el.hidden = true; el.replaceChildren(); }

  // ---------- confirm ----------
  // Every change goes through here: tenant, service principal, API or site, and the list.
  function confirmChange({ title, api, site, items, note, noteKind = 'warn', okLabel, danger = false, showSpn = true }) {
    const facts = Spns.confirmFacts({ tenant: Auth.tenantInfo(), spn: showSpn ? state.selectedSpn : null, api, site });
    const body = h('div', null, [
      h('dl.confirm-facts', null, facts.flatMap(([k, v]) => [h('dt', { text: k }), h('dd', { text: v })])),
      note ? h('div', { class: `notice notice-${noteKind}` }, [Toolkit.icon(NOTICE_ICON[noteKind]), h('div', { text: note })]) : null,
    ]);
    return Toolkit.confirm({ title, body, items, itemsMax: 500, okLabel, danger, wide: true });
  }

  // Write consent right after the confirm, before the first call.
  async function ensureWrite() {
    try { await Graph.tokens.write(); return true; }
    catch (e) { Log.err('Write token failed:', e); toast(`No write permission: ${e.message}`, 'err'); return false; }
  }

  // ---------- auth ----------
  function tenantText() {
    const t = Auth.tenantInfo();
    return t ? `Tenant ${t.domain || '(unknown domain)'} (${t.tenantId})` : '';
  }
  function refreshAuthUI() {
    const signedIn = !!Auth.getAccount();
    setUser(Auth.userInfo(onSignOut));
    $('connectCard').hidden = signedIn;
    $('browseCard').hidden = !signedIn;
    $('selTenant').textContent = tenantText();
  }
  async function onSignIn() {
    try { await Auth.signIn('read'); }
    catch (e) { Log.err('Sign-in failed:', e); toast(`Sign-in failed: ${e.message}`, 'err'); }
  }
  async function onSignOut() {
    state.spns = []; state.filteredSpns = []; state.apiSpns = []; state.selectedSpn = null;
    await Auth.signOut();
  }

  // ---------- load service principals ----------
  async function loadSpns() {
    const box = $('progressBox');
    box.hidden = false;
    $('progressCount').textContent = '0 loaded';
    state.cancelToken = new Concurrency.CancelToken();
    $('cancelBtn').onclick = () => state.cancelToken.cancel();
    const t0 = Date.now();
    await busy($('loadSpnsBtn'), async () => {
      try {
        const list = await Graph.listAllServicePrincipals((loaded, page) => {
          $('progressCount').textContent = `${loaded} loaded, page ${page}`;
        }, state.cancelToken.signal);
        state.spns = list;
        state.apiSpns = sortApiSpns(list.filter(Spns.hasAssignablePermissions));
        const dt = ((Date.now() - t0) / 1000).toFixed(1);
        $('loadStatus').textContent = `${list.length} service principals loaded in ${dt}s. Click one to see and change its permissions.`;
        Log.info(`Loaded ${list.length} service principals in ${dt}s`);
        populateTypeFilter();
        applySpnFilter();
      } catch (e) {
        if (e.message === 'Cancelled') { toast('Loading cancelled', 'warn'); return; }
        Log.err('Load failed:', e);
        toast(`Loading failed: ${e.message}`, 'err');
      } finally {
        box.hidden = true;
      }
    });
  }

  function populateTypeFilter() {
    const sel = $('typeFilter');
    const types = [...new Set(state.spns.map(Spns.classifySpn))].sort();
    sel.replaceChildren(h('option', { value: '', text: 'All types' }), ...types.map((t) => h('option', { value: t, text: t })));
  }

  function applySpnFilter() {
    const q = $('spnFilter').value.trim().toLowerCase();
    const type = $('typeFilter').value;
    const hideMs = $('hideMs').checked;
    state.filteredSpns = state.spns.filter((sp) => {
      if (hideMs && Spns.isMicrosoftFirstParty(sp)) return false;
      if (type && Spns.classifySpn(sp) !== type) return false;
      if (!q) return true;
      return `${sp.displayName || ''}\n${sp.appId || ''}\n${sp.appDisplayName || ''}`.toLowerCase().includes(q);
    });
    renderSpnTable();
  }

  function sortSpns(rows) {
    const { col, dir } = state.sort;
    const m = dir === 'asc' ? 1 : -1;
    const k = (sp) => {
      switch (col) {
        case 'name': return (sp.displayName || '').toLowerCase();
        case 'type': return Spns.classifySpn(sp);
        case 'appId': return sp.appId || '';
        case 'enabled': return sp.accountEnabled ? 1 : 0;
        default: return '';
      }
    };
    return rows.slice().sort((a, b) => (k(a) > k(b) ? m : k(a) < k(b) ? -m : 0));
  }

  function renderSpnTable() {
    const has = state.spns.length > 0;
    ['spnFilterRow', 'spnTableWrap', 'spnHint', 'spnActionBar'].forEach((id) => { $(id).hidden = !has; });
    const tbody = document.querySelector('#spnTable tbody');
    if (!has) { tbody.replaceChildren(); return; }
    const rows = sortSpns(state.filteredSpns).slice(0, MAX_ROWS);
    $('spnHint').textContent = `${state.filteredSpns.length} match${state.filteredSpns.length === 1 ? '' : 'es'}` +
      (state.filteredSpns.length > rows.length ? `, showing the first ${rows.length}` : '');
    tbody.innerHTML = rows.map((sp) => `
      <tr data-id="${esc(sp.id)}" class="clickable-row">
        <td><strong>${esc(sp.displayName || '(no name)')}</strong></td>
        <td>${esc(Spns.classifySpn(sp))}</td>
        <td><code>${esc(sp.appId || '')}</code></td>
        <td>${sp.accountEnabled ? '<span class="badge ok">Yes</span>' : '<span class="badge neutral">No</span>'}</td>
      </tr>`).join('');
    document.querySelectorAll('#spnTable thead th').forEach((th) => {
      th.classList.remove('sort-asc', 'sort-desc');
      if (th.dataset.sort) {
        th.classList.add('sortable-col');
        if (th.dataset.sort === state.sort.col) th.classList.add(state.sort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
      }
    });
    $('exportSpnXlsxBtn').disabled = $('exportSpnCsvBtn').disabled = state.filteredSpns.length === 0;
  }

  function onSpnTableClick(e) {
    const tr = e.target.closest('tr[data-id]');
    if (!tr) return;
    const sp = state.spns.find((s) => s.id === tr.dataset.id);
    if (sp) selectSpn(sp);
  }
  function onSpnSortClick(e) {
    const th = e.target.closest('th[data-sort]');
    if (!th) return;
    const col = th.dataset.sort;
    if (state.sort.col === col) state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
    else state.sort = { col, dir: 'asc' };
    renderSpnTable();
  }

  // ---------- selected service principal ----------
  let activateTab = () => {};

  async function selectSpn(sp) {
    state.selectedSpn = sp;
    state.appRoles = []; state.oauth2Grants = [];
    state.selRoles.clear(); state.selScopes.clear();
    state.siteManager.optimisticUnlock = false;
    ['revokeResults', 'appRolesResults', 'oauth2Results'].forEach((id) => { $(id).hidden = true; });
    activateTab('manage');
    renderManageTab();
    await loadPermissions();
  }

  async function loadPermissions() {
    const sp = state.selectedSpn;
    if (!sp) return;
    ['appRolesCard', 'oauth2Card', 'addCard'].forEach((id) => { $(id).hidden = false; });
    const loading = '<tr><td colspan="5" class="muted">Loading...</td></tr>';
    document.querySelector('#appRolesTable tbody').innerHTML = loading;
    document.querySelector('#oauth2Table tbody').innerHTML = loading;
    try {
      const [appRoles, oauth2] = await Promise.all([Graph.listAppRoleAssignments(sp.id), Graph.listOAuth2Grants(sp.id)]);
      if (state.selectedSpn !== sp) return;
      state.appRoles = appRoles;
      state.oauth2Grants = oauth2;
      const resourceIds = new Set([...appRoles, ...oauth2].map((x) => x.resourceId).filter(Boolean));
      await Promise.all([...resourceIds].map(ensureResource));
      renderAppRoles();
      renderOAuth2();
      renderAddPanel();
      renderManageTab();
      await autoSelectDefaultApi();
    } catch (e) {
      Log.err('Loading permissions failed:', e);
      toast(`Loading permissions failed: ${e.message}`, 'err');
    }
  }

  async function ensureResource(resourceId) {
    if (state.resourceCache.has(resourceId)) return state.resourceCache.get(resourceId);
    try {
      const sp = await Graph.getServicePrincipal(resourceId);
      state.resourceCache.set(resourceId, sp);
      state.resourceIndexCache.set(resourceId, Spns.buildPermissionIndex(sp));
      return sp;
    } catch (e) {
      Log.warn(`Could not resolve resource service principal ${resourceId}: ${e.message}`);
      state.resourceCache.set(resourceId, null);
      state.resourceIndexCache.set(resourceId, new Map());
      return null;
    }
  }

  function renderManageTab() {
    const sp = state.selectedSpn;
    $('noSelectionCard').hidden = !!sp;
    $('spnHeaderCard').hidden = !sp;
    if (!sp) {
      ['appRolesCard', 'oauth2Card', 'addCard'].forEach((id) => { $(id).hidden = true; });
      return;
    }
    $('selSpnName').textContent = sp.displayName || '(no name)';
    $('selSpnMeta').textContent = [`appId ${sp.appId || '?'}`, Spns.classifySpn(sp), sp.publisherName, `Enabled: ${sp.accountEnabled ? 'yes' : 'no'}`].filter(Boolean).join(' · ');
    $('selTenant').textContent = tenantText();
    const hasAny = state.appRoles.length > 0 || state.oauth2Grants.length > 0;
    $('revokeAllBtn').disabled = !hasAny;
    $('exportPermsBtn').disabled = !hasAny;
    $('mailboxHint').hidden = !Spns.hasMailboxPermissions(state.appRoles, state.resourceCache, state.resourceIndexCache);
    updateManageSitesButton();
  }

  function renderAppRoles() {
    const tbody = document.querySelector('#appRolesTable tbody');
    $('appRolesCount').textContent = `(${state.appRoles.length})`;
    $('appRolesEmpty').hidden = state.appRoles.length > 0;
    tbody.innerHTML = state.appRoles.map((a) => {
      const res = state.resourceCache.get(a.resourceId);
      const meta = (state.resourceIndexCache.get(a.resourceId) || new Map()).get(a.appRoleId);
      const granted = a.createdDateTime ? new Date(a.createdDateTime).toLocaleString() : '';
      const name = meta ? meta.name : a.appRoleId;
      return `<tr data-id="${esc(a.id)}">
        <td class="col-check"><input type="checkbox" data-act="sel-role" aria-label="Select ${esc(name)}" ${state.selRoles.has(a.id) ? 'checked' : ''}/></td>
        <td>${esc((res && res.displayName) || a.resourceDisplayName || a.resourceId)}</td>
        <td><code>${esc(name)}</code></td>
        <td class="muted small">${esc(meta ? meta.description : '')}</td>
        <td class="muted small nowrap">${esc(granted)}</td>
      </tr>`;
    }).join('');
    updateRevokeSelection();
  }

  // One row per scope, so single scopes can be picked.
  function scopeRows() {
    const rows = [];
    for (const g of state.oauth2Grants) {
      const res = state.resourceCache.get(g.resourceId);
      const apiName = (res && res.displayName) || g.resourceId;
      const scopes = Spns.parseScopeString(g.scope);
      if (!scopes.length) rows.push({ apiName, grant: g, scope: '' });
      for (const s of scopes) rows.push({ apiName, grant: g, scope: s });
    }
    return rows.map((r) => ({ ...r, key: Spns.scopeKey(r.grant.id, r.scope) }));
  }

  function renderOAuth2() {
    const tbody = document.querySelector('#oauth2Table tbody');
    $('oauth2Count').textContent = `(${state.oauth2Grants.length})`;
    $('oauth2Empty').hidden = state.oauth2Grants.length > 0;
    tbody.innerHTML = scopeRows().map((r) => `
      <tr data-key="${esc(r.key)}">
        <td class="col-check"><input type="checkbox" data-act="sel-scope" aria-label="Select ${esc(r.scope || 'grant without scopes')}" ${state.selScopes.has(r.key) ? 'checked' : ''}/></td>
        <td>${esc(r.apiName)}</td>
        <td><code>${esc(r.scope || '(none)')}</code></td>
        <td>${esc(r.grant.consentType)}</td>
        <td class="muted small">${esc(r.grant.principalId || (r.grant.consentType === 'AllPrincipals' ? '(all users)' : '(unknown)'))}</td>
      </tr>`).join('');
    updateRevokeSelection();
  }

  function setAllBox(cb, n, total) {
    cb.disabled = total === 0;
    cb.checked = total > 0 && n === total;
    cb.indeterminate = n > 0 && n < total;
  }

  // Selections that no longer exist after a reload are dropped.
  function updateRevokeSelection() {
    const roleIds = new Set(state.appRoles.map((a) => a.id));
    for (const id of [...state.selRoles]) if (!roleIds.has(id)) state.selRoles.delete(id);
    const keys = new Set(scopeRows().map((r) => r.key));
    for (const k of [...state.selScopes]) if (!keys.has(k)) state.selScopes.delete(k);
    $('appRolesSelCount').textContent = String(state.selRoles.size);
    $('revokeAppsBtn').disabled = state.selRoles.size === 0;
    $('appRolesBar').hidden = roleIds.size === 0;
    setAllBox($('appRolesAll'), state.selRoles.size, roleIds.size);
    $('oauth2SelCount').textContent = String(state.selScopes.size);
    $('revokeScopesBtn').disabled = state.selScopes.size === 0;
    $('oauth2Bar').hidden = keys.size === 0;
    setAllBox($('oauth2All'), state.selScopes.size, keys.size);
  }

  function onRoleCheck(e) {
    const cb = e.target.closest('input[data-act="sel-role"]');
    if (!cb) return;
    const id = cb.closest('tr').dataset.id;
    if (cb.checked) state.selRoles.add(id); else state.selRoles.delete(id);
    updateRevokeSelection();
  }
  function onScopeCheck(e) {
    const cb = e.target.closest('input[data-act="sel-scope"]');
    if (!cb) return;
    const key = cb.closest('tr').dataset.key;
    if (cb.checked) state.selScopes.add(key); else state.selScopes.delete(key);
    updateRevokeSelection();
  }
  function onAllRoles(e) {
    state.selRoles = new Set(e.target.checked ? state.appRoles.map((a) => a.id) : []);
    renderAppRoles();
  }
  function onAllScopes(e) {
    state.selScopes = new Set(e.target.checked ? scopeRows().map((r) => r.key) : []);
    renderOAuth2();
  }

  // ---------- revoke ----------
  // Runs jobs ({ label, run }) with a live result list. Returns { done, failed }.
  async function runJobs(jobs, listEl, { concurrency = 4, doneVerb = 'Removed' } = {}) {
    listEl.hidden = false;
    listEl.replaceChildren();
    let done = 0;
    const res = await Concurrency.pmap(jobs, async (job) => {
      const li = h('li', null, [h('span.warn-mark', { text: '...' }), ` ${job.label}`]);
      listEl.appendChild(li);
      try {
        await job.run();
        done++;
        li.replaceChildren(Toolkit.icon('i-check', 'sm ok-mark'), ` ${doneVerb}: ${job.label}`);
      } catch (e) {
        li.replaceChildren(Toolkit.icon('i-error', 'sm err-mark'), ` ${job.label}: ${e.message}`);
        Log.err(`Failed: ${job.label}:`, e);
        throw e;
      }
    }, { concurrency });
    const failed = res.errors.length;
    toast(`${doneVerb} ${done} of ${jobs.length}${failed ? `, ${failed} failed (see the list and the debug log)` : ''}`, failed ? 'warn' : 'ok');
    return { done, failed };
  }

  async function onRevokeAppsClick() {
    const sp = state.selectedSpn;
    const picks = state.appRoles.filter((a) => state.selRoles.has(a.id));
    if (!sp || !picks.length) return;
    const label = (a) => Spns.appRoleLabel(a, state.resourceCache, state.resourceIndexCache);
    const ok = await confirmChange({
      title: `Revoke ${plural(picks.length, 'application permission')}?`,
      items: picks.map((a) => `Remove application permission ${label(a)}`),
      note: 'The service principal can stop working if it relies on these permissions.',
      okLabel: `Revoke ${picks.length}`,
      danger: true,
    });
    if (!ok || !(await ensureWrite())) return;
    await busy($('revokeAppsBtn'), async () => {
      await runJobs(picks.map((a) => ({ label: `Application permission ${label(a)}`, run: () => Graph.removeAppRoleAssignment(sp.id, a.id) })), $('appRolesResults'));
      state.selRoles.clear();
      state.siteManager.optimisticUnlock = false;
      await loadPermissions();
    });
  }

  async function onRevokeScopesClick() {
    const sp = state.selectedSpn;
    const plan = Spns.planScopeRevocation(state.oauth2Grants, state.selScopes);
    if (!sp || !plan.length) return;
    const apiName = (g) => { const r = state.resourceCache.get(g.resourceId); return (r && r.displayName) || g.resourceId; };
    const who = (g) => (g.consentType === 'AllPrincipals' ? 'all users' : `user ${g.principalId || '?'}`);
    const items = plan.map((p) => (p.action === 'delete'
      ? `Remove delegated grant ${Spns.grantLabel(p.grant, state.resourceCache)}`
      : `Remove ${apiName(p.grant)}: ${p.remove.join(' ')} (${who(p.grant)}), keep ${p.keep.join(' ')}`));
    const deletes = plan.filter((p) => p.action === 'delete').length;
    const scopeCount = plan.reduce((n, p) => n + Math.max(1, p.remove.length), 0);
    const ok = await confirmChange({
      title: `Revoke ${plural(scopeCount, 'delegated scope')}?`,
      items,
      note: deletes ? `${plural(deletes, 'grant')} lose${deletes === 1 ? 's' : ''} its last scope and ${deletes === 1 ? 'is' : 'are'} deleted.` : null,
      okLabel: `Revoke ${scopeCount}`,
      danger: true,
    });
    if (!ok || !(await ensureWrite())) return;
    await busy($('revokeScopesBtn'), async () => {
      await runJobs(plan.map((p, i) => ({
        label: items[i].replace(/^Remove /, ''),
        run: () => (p.action === 'delete' ? Graph.deleteOAuth2Grant(p.grant.id) : Graph.updateOAuth2Grant(p.grant.id, { scope: Spns.joinScopeString(p.keep) })),
      })), $('oauth2Results'));
      state.selScopes.clear();
      await loadPermissions();
    });
  }

  async function onRevokeAllClick() {
    const sp = state.selectedSpn;
    if (!sp) return;
    const appRoles = state.appRoles.slice();
    const grants = state.oauth2Grants.slice();
    const summary = Spns.summarizePermissionRemoval(appRoles, grants);
    if (!summary.totalPermissionEntries) return;
    const ok = await confirmChange({
      title: 'Revoke all permissions?',
      items: Spns.revokeAllItems(appRoles, grants, state.resourceCache, state.resourceIndexCache),
      note: `${plural(summary.appRoleCount, 'application permission')} and ${plural(summary.grantCount, 'delegated grant')} (${plural(summary.delegatedScopeCount, 'scope')}) are removed. There is no undo; if the workload still needs access you grant it again.`,
      noteKind: 'err',
      okLabel: `Revoke all ${summary.totalPermissionEntries}`,
      danger: true,
    });
    if (!ok || !(await ensureWrite())) return;

    const jobs = [
      ...appRoles.map((a) => ({ label: `Application permission ${Spns.appRoleLabel(a, state.resourceCache, state.resourceIndexCache)}`, run: () => Graph.removeAppRoleAssignment(sp.id, a.id) })),
      ...grants.map((g) => ({ label: `Delegated grant ${Spns.grantLabel(g, state.resourceCache)}`, run: () => Graph.deleteOAuth2Grant(g.id) })),
    ];
    await busy($('revokeAllBtn'), async () => {
      await runJobs(jobs, $('revokeResults'));
      state.selRoles.clear(); state.selScopes.clear();
      state.siteManager.optimisticUnlock = false;
      await loadPermissions();
    });
  }

  // ---------- add permission ----------
  function sortApiSpns(spns) {
    return spns.slice().sort((a, b) => (a.displayName || '').toLowerCase().localeCompare((b.displayName || '').toLowerCase()) || (a.appId || '').localeCompare(b.appId || ''));
  }

  function renderAddPanel() {
    $('apiSuggestions').replaceChildren(...state.apiSpns.map((sp) => h('option', { value: Spns.apiOptionLabel(sp) })));
    state.addApiSpn = null;
    state.addSelected.clear();
    document.querySelector('#addPermTable tbody').innerHTML = '<tr><td colspan="4" class="muted">Pick an API to see its permissions.</td></tr>';
    updateAddSelection();
    $('addResults').hidden = true;
    clearNotice($('addFeedback'));
    $('apiLoadStatus').textContent = `${state.apiSpns.length} APIs with assignable permissions. Pick one from the list, it shows the appId.`;
  }

  async function autoSelectDefaultApi() {
    const graph = state.apiSpns.find((s) => (s.appId || '').toLowerCase() === Spns.GRAPH_APP_ID);
    if (!graph) return;
    $('apiSearch').value = Spns.apiOptionLabel(graph);
    await loadApiForQuery($('apiSearch').value);
  }

  async function loadApiForQuery(query) {
    const status = $('apiLoadStatus');
    const r = Spns.resolveApiQuery(state.apiSpns, query);
    if (!r.sp) {
      status.textContent = r.error;
      state.addApiSpn = null;
      state.addSelected.clear();
      updateAddSelection();
      renderAddPermTable();
      return null;
    }
    let api = r.sp;
    try {
      if (!Array.isArray(api.appRoles) || !Array.isArray(api.oauth2PermissionScopes)) api = await Graph.getServicePrincipal(api.id);
    } catch (e) {
      Log.err('Loading API failed:', e);
      status.textContent = `Loading the API failed: ${e.message}`;
      return null;
    }
    state.addApiSpn = api;
    state.resourceCache.set(api.id, api);
    state.resourceIndexCache.set(api.id, Spns.buildPermissionIndex(api));
    $('apiSearch').value = Spns.apiOptionLabel(api);
    const t = Auth.tenantInfo();
    status.textContent = `${api.displayName}, appId ${api.appId}, owned by ${Spns.ownerLabel(api, t && t.tenantId)}: ` +
      `${plural((api.appRoles || []).length, 'app role')}, ${plural((api.oauth2PermissionScopes || []).length, 'delegated scope')}.`;
    state.addSelected.clear();
    updateAddSelection();
    renderAddPermTable();
    return api;
  }

  function visibleAddItems() {
    const api = state.addApiSpn;
    if (!api) return [];
    const filter = $('permFilter').value.trim().toLowerCase();
    const isAppRole = state.addApiKind === 'AppRole';
    const items = (isAppRole ? api.appRoles : api.oauth2PermissionScopes) || [];
    return items.filter((it) => {
      if (it.isEnabled === false) return false;
      if (isAppRole && !(it.allowedMemberTypes || []).includes('Application')) return false;
      if (!filter) return true;
      return `${it.value || ''}\n${it.displayName || it.adminConsentDisplayName || ''}\n${it.description || it.adminConsentDescription || ''}`.toLowerCase().includes(filter);
    }).sort((a, b) => (a.value || '').localeCompare(b.value || ''));
  }

  function renderAddPermTable() {
    const tbody = document.querySelector('#addPermTable tbody');
    if (!state.addApiSpn) { tbody.innerHTML = '<tr><td colspan="4" class="muted">Pick an API above.</td></tr>'; return; }
    const items = visibleAddItems();
    // Only what is visible stays selected.
    const visible = new Set(items.map((it) => it.id));
    for (const id of [...state.addSelected]) if (!visible.has(id)) state.addSelected.delete(id);
    updateAddSelection();
    if (!items.length) { tbody.innerHTML = '<tr><td colspan="4" class="muted">No matching permissions.</td></tr>'; return; }
    const isAppRole = state.addApiKind === 'AppRole';
    tbody.innerHTML = items.map((it) => {
      const name = it.value || it.displayName || it.adminConsentDisplayName || it.id;
      return `<tr data-id="${esc(it.id)}">
        <td class="col-check"><input type="checkbox" data-act="sel-perm" aria-label="Select ${esc(name)}" ${state.addSelected.has(it.id) ? 'checked' : ''}/></td>
        <td><code>${esc(name)}</code><br><span class="muted small">${esc(it.displayName || it.adminConsentDisplayName || '')}</span></td>
        <td>${isAppRole ? 'Application' : 'Delegated'}</td>
        <td class="muted small">${esc(it.description || it.adminConsentDescription || it.userConsentDescription || '')}</td>
      </tr>`;
    }).join('');
  }

  function updateAddSelection() {
    $('addSelCount').textContent = String(state.addSelected.size);
    $('addPermsBtn').disabled = state.addSelected.size === 0;
  }

  function onAddPermTableChange(e) {
    const cb = e.target.closest('input[data-act="sel-perm"]');
    if (!cb) return;
    const id = cb.closest('tr').dataset.id;
    if (cb.checked) state.addSelected.add(id); else state.addSelected.delete(id);
    updateAddSelection();
  }

  async function onAddPermsClick() {
    const sp = state.selectedSpn;
    const api = state.addApiSpn;
    if (!sp || !api || !state.addSelected.size) return;
    const isAppRole = state.addApiKind === 'AppRole';
    const picks = visibleAddItems().filter((it) => state.addSelected.has(it.id));
    if (!picks.length) return;
    const existing = isAppRole ? null : state.oauth2Grants.find((g) => g.resourceId === api.id && g.consentType === 'AllPrincipals');
    const kind = isAppRole ? 'application permission' : 'delegated permission';
    const ok = await confirmChange({
      title: `Grant ${plural(picks.length, kind)}?`,
      api,
      items: picks.map((p) => `Add ${kind} ${api.displayName}: ${p.value || p.displayName || p.id}`),
      note: isAppRole
        ? 'Application permissions work tenant-wide without a signed-in user. Make sure this service principal should have them.'
        : `Delegated grants are added with consentType AllPrincipals, which covers every user in the tenant${existing ? ', merged into the existing grant for this API' : ''}.`,
      noteKind: isAppRole ? 'warn' : 'info',
      okLabel: `Grant ${picks.length}`,
    });
    if (!ok || !(await ensureWrite())) return;

    const resultsEl = $('addResults');
    resultsEl.hidden = false;
    resultsEl.replaceChildren();
    clearNotice($('addFeedback'));
    let success = 0;
    let failed = 0;
    const line = (iconName, cls, text) => h('li', null, [Toolkit.icon(iconName, `sm ${cls}`), ` ${text}`]);

    await busy($('addPermsBtn'), async () => {
      if (isAppRole) {
        for (const p of picks) {
          const label = p.value || p.id;
          const li = h('li', null, [h('span.warn-mark', { text: '...' }), ` ${label}`]);
          resultsEl.appendChild(li);
          try {
            await Graph.addAppRoleAssignment(sp.id, { resourceId: api.id, appRoleId: p.id });
            li.replaceWith(line('i-check', 'ok-mark', `Granted ${label}`));
            success++;
          } catch (e) {
            li.replaceWith(line('i-error', 'err-mark', `${label}: ${e.message}`));
            Log.err('Grant failed:', e);
            failed++;
          }
        }
      } else {
        // One AllPrincipals grant per client and resource: patch it, or create it.
        const newScopes = picks.map((p) => p.value).filter(Boolean);
        try {
          if (existing) {
            await Graph.updateOAuth2Grant(existing.id, { scope: Spns.joinScopeString([...Spns.parseScopeString(existing.scope), ...newScopes]) });
            resultsEl.appendChild(line('i-check', 'ok-mark', `Added to the existing grant: ${newScopes.join(' ')}`));
          } else {
            await Graph.createOAuth2Grant({ clientId: sp.id, resourceId: api.id, scope: Spns.joinScopeString(newScopes), consentType: 'AllPrincipals' });
            resultsEl.appendChild(line('i-check', 'ok-mark', `Created a grant with: ${newScopes.join(' ')}`));
          }
          success += newScopes.length;
        } catch (e) {
          resultsEl.appendChild(line('i-error', 'err-mark', e.message));
          Log.err('Grant failed:', e);
          failed += newScopes.length;
        }
      }
    });

    const fb = $('addFeedback');
    if (success && !failed) { setNotice(fb, `${plural(success, 'permission')} granted.`, 'ok'); toast(`Granted ${plural(success, 'permission')}`, 'ok'); }
    else if (success) { setNotice(fb, `${success} granted, ${failed} failed. Details below and in the debug log.`, 'warn'); toast(`${success} granted, ${failed} failed`, 'warn'); }
    else { setNotice(fb, 'Nothing was granted. Details below and in the debug log.', 'err'); toast('Grant failed, see the result list', 'err'); }

    state.addSelected.clear();
    updateAddSelection();
    // Graph takes a while to show a new assignment; unlock the site manager right away.
    if (isAppRole && success && [Spns.GRAPH_APP_ID, Spns.SPO_APP_ID].includes((api.appId || '').toLowerCase()) &&
        picks.some((p) => (p.value || '').toLowerCase() === 'sites.selected')) {
      state.siteManager.optimisticUnlock = true;
    }
    // Reloading resets the panel to Graph; stay on the API the user was working with.
    const keepApi = Spns.apiOptionLabel(api);
    await loadPermissions();
    if ($('apiSearch').value !== keepApi) await loadApiForQuery(keepApi);
  }

  // ---------- site manager (Sites.Selected) ----------
  function canManageSites() {
    return state.siteManager.optimisticUnlock || Spns.hasSitesSelected(state.appRoles, state.resourceCache, state.resourceIndexCache);
  }
  function updateManageSitesButton() {
    const btn = $('manageSitesBtn');
    const can = canManageSites();
    btn.disabled = !can;
    btn.title = can
      ? 'Grant this service principal access to specific SharePoint sites (Sites.Selected).'
      : 'Needs the Sites.Selected application permission on Microsoft Graph or SharePoint Online first. Add it below.';
  }

  const SITE_ROWS_MAX = 300;     // listed, and checked for access, per filter
  const SCAN_CONCURRENCY = 6;
  const ACCESS_WAIT_TRIES = 8;   // after SetSiteAdmin, until Graph lets you in
  const ACCESS_WAIT_MS = 4000;
  let siteSearchTimer = null;
  let siteRenderTimer = null;
  const keyOf = (site) => Spns.siteKey(site && site.webUrl);
  const siteName = (site) => (site && (site.displayName || site.name || site.webUrl)) || '(no name)';
  const spnName = () => state.selectedSpn.displayName || state.selectedSpn.appId;
  const modalOpen = () => !$('sitePickerModal').hidden;

  async function onManageSitesClick() {
    const sp = state.selectedSpn;
    if (!sp || !canManageSites()) return;
    if (!sp.appId) { toast('This service principal has no appId, so it cannot get site access.', 'err'); return; }
    await busy($('manageSitesBtn'), async () => {
      try { await Graph.tokens.sites(); }
      catch (e) {
        Log.err('Sites token failed:', e);
        toast(`No SharePoint permission: ${e.message}`, 'err');
        return;
      }
      openSiteManager();
    });
    updateManageSitesButton();
  }

  function openSiteManager() {
    const sm = state.siteManager;
    sm.selectedSite = null; sm.selectedKey = null; sm.sitePerms = []; sm.readable = false;
    sm.selectedRoles = new Set(); sm.selectedPerms = new Set();
    sm.access = new Map(); sm.scan = null; sm.scanToken++; sm.detailToken++;
    sm.returnFocus = document.activeElement;
    $('siteModalSpn').textContent = spnName();
    $('siteSearch').value = '';
    $('siteDetail').hidden = true;
    $('siteResults').replaceChildren();
    $('sitePickerModal').hidden = false;
    renderRolePicker();
    $('siteSearch').focus();
    if (sm.allSites) applySiteFilter(); else loadAllSites();
  }

  async function closeSiteManager() {
    const sm = state.siteManager;
    if (!modalOpen()) return;
    $('sitePickerModal').hidden = true;
    sm.searchToken++; sm.scanToken++; sm.detailToken++;
    sm.scan = null;
    if (siteSearchTimer) { clearTimeout(siteSearchTimer); siteSearchTimer = null; }
    if (sm.returnFocus && sm.returnFocus.focus) try { sm.returnFocus.focus(); } catch { /* gone */ }
    const open = myElevations();
    if (!open.length) return;
    const ok = await Toolkit.confirm({
      title: 'Remove your temporary site admin rights?',
      body: 'This tool made you site collection admin of these sites. Remove that now, or keep it and remove it later from the notice at the top of the page.',
      items: open.map((r) => r.siteUrl),
      okLabel: 'Remove my admin rights',
      cancelLabel: 'Keep for now',
    });
    if (ok) await demote(open.map((r) => r.siteUrl));
  }

  // Every site collection from the admin API; Graph search when that is refused.
  async function loadAllSites() {
    const sm = state.siteManager;
    const status = $('siteSearchStatus');
    const btn = $('siteLoadAllBtn');
    btn.hidden = true;
    status.textContent = 'Loading all sites from the SharePoint admin API...';
    try {
      sm.allSites = await Spo.listSites((n) => { if (modalOpen()) status.textContent = `Loading all sites: ${n} so far...`; });
      sm.source = 'admin';
      Log.info(`${sm.allSites.length} site collections from the SharePoint admin API`);
    } catch (e) {
      Log.warn(`Full site list not available, using Graph search: ${e.message}`);
      sm.source = 'search';
      btn.hidden = false;
      btn.title = `The SharePoint admin API refused the site list: ${e.message}`;
    }
    if (modalOpen()) await applySiteFilter();
  }

  function onSiteSearchInput() {
    if (siteSearchTimer) clearTimeout(siteSearchTimer);
    siteSearchTimer = setTimeout(applySiteFilter, state.siteManager.source === 'admin' ? 120 : 300);
  }

  async function applySiteFilter() {
    const sm = state.siteManager;
    const status = $('siteSearchStatus');
    const q = $('siteSearch').value.trim();
    const token = ++sm.searchToken;
    const isUrl = /^[a-z][a-z0-9+.-]*:\/\//i.test(q);
    if (isUrl && !Spns.siteCollectionPath(q)) {
      sm.sites = [];
      renderSiteResults();
      status.textContent = 'Only https URLs on *.sharepoint.com.';
      return;
    }
    let sites = [];
    if (sm.source === 'admin') {
      sites = Spns.filterSites(sm.allSites, q);
    } else if (!isUrl) {
      status.textContent = 'Searching...';
      try { sites = await Graph.searchSitesLoose(q || '*'); }
      catch (e) {
        if (token !== sm.searchToken) return;
        Log.err('Site search failed:', e);
        status.textContent = `Search failed: ${e.message}`;
        return;
      }
    }
    // A pasted URL that is not listed (search hides sites you cannot open): try it directly.
    if (isUrl && !sites.length) sites = [{ id: null, displayName: Spns.siteKey(q), webUrl: Spns.siteKey(q) }];
    if (token !== sm.searchToken) return;
    sm.sites = sites;
    renderSiteResults();
    updateSiteStatus();
    scanAccess(false);
    if (isUrl && sites.length === 1) selectSite(sites[0]);
  }

  function renderSiteResults() {
    const sm = state.siteManager;
    const ul = $('siteResults');
    const focused = ul.contains(document.activeElement) ? document.activeElement.dataset.key : null;
    sm.rendered = Spns.sortSitesByAccess(sm.sites, sm.access).slice(0, SITE_ROWS_MAX);
    if (!sm.rendered.length) { ul.innerHTML = '<li class="muted site-empty">No sites match.</li>'; return; }
    ul.innerHTML = sm.rendered.map((s, i) => {
      const key = keyOf(s);
      const a = sm.access.get(key);
      let mark = '';
      if (a && a.status === 'yes') mark = `<span class="pill brand site-mark">${esc(a.roles.join(', ') || 'access')}</span>`;
      else if (a && a.status === 'denied') mark = `<span class="site-mark muted" title="Not checked: you are not an admin of this site">${Toolkit.iconHtml('i-lock', 'sm')}</span>`;
      else if (a && a.status === 'error') mark = `<span class="site-mark muted" title="${esc(`Check failed: ${a.message || ''}`)}">${Toolkit.iconHtml('i-alert', 'sm')}</span>`;
      return `<li class="site-list-item${a && a.status === 'yes' ? ' has-access' : ''}${sm.selectedKey === key ? ' active' : ''}" data-idx="${i}" data-key="${esc(key || '')}" tabindex="0">
        <span class="site-name">${esc(siteName(s))}${mark}</span>
        <span class="muted small site-url">${esc(s.webUrl || '')}</span>
      </li>`;
    }).join('');
    if (focused) {
      const li = [...ul.querySelectorAll('li[data-key]')].find((x) => x.dataset.key === focused);
      if (li) li.focus();
    }
  }

  // Re-sorting on every finished check makes the list jump; batch it.
  function scheduleSiteRender() {
    if (siteRenderTimer) return;
    siteRenderTimer = setTimeout(() => {
      siteRenderTimer = null;
      if (modalOpen()) { renderSiteResults(); updateSiteStatus(); }
    }, 500);
  }

  function updateSiteStatus() {
    const sm = state.siteManager;
    const total = sm.sites.length;
    let yes = 0; let denied = 0; let unchecked = 0;
    for (const s of sm.sites) {
      const a = sm.access.get(keyOf(s));
      if (!a) unchecked++;
      else if (a.status === 'yes') yes++;
      else if (a.status === 'denied') denied++;
    }
    const parts = [`${plural(total, 'site')}${total > SITE_ROWS_MAX ? `, the first ${SITE_ROWS_MAX} listed` : ''}.`];
    if (sm.source === 'search') parts.push('From Graph search, which matches whole words and prefixes only.');
    if (sm.scan) parts.push(`Checking access: ${sm.scan.done} of ${sm.scan.total}...`);
    if (total) parts.push(`${spnName()} has access to ${yes}${unchecked ? ` of the ${total - unchecked} checked` : ''}.`);
    if (denied) parts.push(`${denied} could not be checked because you are not their admin.`);
    $('siteSearchStatus').textContent = parts.join(' ');
    const more = $('siteCheckAllBtn');
    more.hidden = !!sm.scan || unchecked === 0;
    more.textContent = `Check ${unchecked} more`;
  }

  // Background check of which listed sites the SPN can reach. all: beyond the listed ones.
  async function scanAccess(all) {
    const sm = state.siteManager;
    const token = ++sm.scanToken;
    const seen = new Set();
    const pool = (all ? Spns.sortSitesByAccess(sm.sites, sm.access) : sm.rendered).filter((s) => {
      const k = keyOf(s);
      if (!k || sm.access.has(k) || seen.has(k)) return false;
      seen.add(k);
      return true;
    });
    if (!pool.length) { sm.scan = null; updateSiteStatus(); return; }
    const scan = { done: 0, total: pool.length };
    sm.scan = scan;
    updateSiteStatus();
    const ct = new Concurrency.CancelToken();
    await Concurrency.pmap(pool, async (site) => {
      if (token !== sm.scanToken) { ct.cancel(); return; }
      await checkSiteAccess(site);
      scan.done++;
      if (token === sm.scanToken) scheduleSiteRender();
    }, { concurrency: SCAN_CONCURRENCY, cancelToken: ct });
    if (token !== sm.scanToken) return;
    sm.scan = null;
    renderSiteResults();
    updateSiteStatus();
  }

  async function resolveSite(site, { fresh = false } = {}) {
    const sm = state.siteManager;
    const key = keyOf(site);
    if (!fresh && sm.ids.has(key)) return sm.ids.get(key);
    const root = await Graph.getSiteCollectionByUrl(site.webUrl);
    if (!root || !Spns.isValidSiteId(root.id)) throw new Error(`No SharePoint Online site id for ${site.webUrl}`);
    sm.ids.set(key, root);
    return root;
  }

  function setAccessFromPerms(key, perms) {
    const mine = Spns.findAppPermissionOnSite(perms, state.selectedSpn.appId);
    state.siteManager.access.set(key, mine ? { status: 'yes', roles: mine.roles } : { status: 'no' });
  }

  // 404 counts as no access too: Graph hides sites you cannot open.
  const deniedOrHidden = (e) => Spns.isAccessDenied(e) || (e && e.status === 404);

  async function checkSiteAccess(site) {
    const key = keyOf(site);
    try {
      const root = await resolveSite(site);
      setAccessFromPerms(key, await Graph.getSitePermissions(root.id));
    } catch (e) {
      const denied = deniedOrHidden(e);
      state.siteManager.access.set(key, { status: denied ? 'denied' : 'error', message: e.message });
      if (!denied) Log.warn(`Access check failed for ${site.webUrl}: ${e.message}`);
    }
  }

  function onSiteResultClick(e) {
    const li = e.target.closest('li[data-idx]');
    if (!li) return;
    if (e.type === 'keydown' && e.key !== 'Enter' && e.key !== ' ') return;
    if (e.type === 'keydown') e.preventDefault();
    const site = state.siteManager.rendered[Number(li.dataset.idx)];
    if (site) selectSite(site);
  }

  function setSiteDetailHead(site) {
    $('siteDetailName').textContent = siteName(site);
    const a = $('siteDetailUrl');
    a.textContent = site.webUrl || '';
    if (Spns.parseSharePointUrl(site.webUrl)) a.href = Toolkit.safeUrl(site.webUrl);
    else a.removeAttribute('href');
  }

  // kind: loading | waiting | ready | denied | error
  function setSiteState(kind, text) {
    const sm = state.siteManager;
    $('siteDeniedNote').hidden = kind !== 'denied';
    $('siteElevatedNote').hidden = !(sm.selectedSite && isElevated(sm.selectedSite.webUrl));
    $('sitePermsBox').hidden = kind === 'denied' && !sm.sitePerms.length;
    $('siteGrantBox').hidden = kind !== 'ready';
    const tbody = document.querySelector('#sitePermsTable tbody');
    if (kind === 'loading' || kind === 'waiting' || kind === 'error') {
      const msg = kind === 'loading' ? 'Loading...' : kind === 'waiting' ? 'Waiting for SharePoint to apply your admin rights...' : `Failed: ${text || ''}`;
      tbody.innerHTML = `<tr><td colspan="4" class="muted">${esc(msg)}</td></tr>`;
      $('sitePermsEmpty').hidden = true;
      $('sitePermsBar').hidden = true;
    }
    if (kind === 'denied') $('siteDeniedText').textContent = text || 'You are not an admin of this site, so its app permissions cannot be read or changed.';
  }

  async function selectSite(site) {
    const sm = state.siteManager;
    const token = ++sm.detailToken;
    sm.selectedKey = keyOf(site);
    sm.selectedSite = site;
    sm.sitePerms = []; sm.readable = false; sm.selectedPerms = new Set();
    renderSiteResults();
    $('siteDetail').hidden = false;
    $('siteScopeNote').hidden = true;
    setSiteDetailHead(site);
    clearNotice($('siteGrantFeedback'));
    await loadSiteDetail(site, token);
  }

  // waitForAccess: right after SetSiteAdmin, retry while Graph still refuses.
  async function loadSiteDetail(site, token, { waitForAccess = false } = {}) {
    const sm = state.siteManager;
    setSiteState(waitForAccess ? 'waiting' : 'loading');
    const tries = waitForAccess ? ACCESS_WAIT_TRIES : 1;
    for (let i = 1; ; i++) {
      try {
        const root = await resolveSite(site, { fresh: waitForAccess });
        if (token !== sm.detailToken) return;
        const perms = await Graph.getSitePermissions(root.id);
        if (token !== sm.detailToken) return;
        sm.selectedSite = { ...root, webUrl: root.webUrl || site.webUrl };
        setSiteDetailHead(sm.selectedSite);
        const picked = String(site.webUrl || '').replace(/\/+$/, '').toLowerCase();
        if (root.webUrl && picked && root.webUrl.replace(/\/+$/, '').toLowerCase() !== picked) {
          $('siteScopeText').textContent = `You picked a subsite or library. Sites.Selected is granted per site collection, so this shows "${siteName(root)}". Access granted here covers all its subsites.`;
          $('siteScopeNote').hidden = false;
        }
        sm.sitePerms = perms;
        sm.readable = true;
        setAccessFromPerms(sm.selectedKey, perms);
        renderSiteResults();
        updateSiteStatus();
        setSiteState('ready');
        renderSitePerms();
        return;
      } catch (e) {
        if (token !== sm.detailToken) return;
        const denied = deniedOrHidden(e);
        if (denied && i < tries) { await Concurrency.sleep(ACCESS_WAIT_MS); continue; }
        if (!denied) { Log.err('Loading site permissions failed:', e); setSiteState('error', e.message); return; }
        sm.access.set(sm.selectedKey, { status: 'denied', message: e.message });
        renderSiteResults();
        updateSiteStatus();
        setSiteState('denied', waitForAccess
          ? 'SharePoint still refuses access after making you site admin. Try again in a minute, or check the debug log.'
          : e.status === 404
            ? 'Graph cannot open this site with your account: it does not exist, or you have no access to it.'
            : 'You are not an admin of this site, so its app permissions cannot be read or changed.');
        return;
      }
    }
  }

  function renderSitePerms() {
    const sm = state.siteManager;
    const rows = Spns.extractSiteAppPermissions(sm.sitePerms);
    const ids = new Set(rows.map((r) => r.permissionId));
    for (const id of [...sm.selectedPerms]) if (!ids.has(id)) sm.selectedPerms.delete(id);
    const selfAppId = (state.selectedSpn.appId || '').toLowerCase();
    $('sitePermsCount').textContent = `(${rows.length})`;
    $('sitePermsEmpty').hidden = rows.length > 0;
    $('sitePermsBar').hidden = rows.length === 0;
    document.querySelector('#sitePermsTable tbody').innerHTML = rows.map((r) => {
      const isSelf = r.appId.toLowerCase() === selfAppId;
      return `<tr class="${isSelf ? 'self-perm' : ''}">
        <td class="col-check"><input type="checkbox" data-act="sel-site-perm" data-perm="${esc(r.permissionId)}" aria-label="Select ${esc(r.displayName || r.appId)}" ${sm.selectedPerms.has(r.permissionId) ? 'checked' : ''} ${sm.readable ? '' : 'disabled'}/></td>
        <td>${esc(r.displayName || '(unknown app)')}${isSelf ? ' <span class="pill brand">this one</span>' : ''}</td>
        <td><code>${esc(r.appId)}</code></td>
        <td><code>${esc(r.roles.join(', ') || '(none)')}</code></td>
      </tr>`;
    }).join('');
    updateSitePermSelection();
    if (sm.readable) renderGrantBox();
  }

  function updateSitePermSelection() {
    const sm = state.siteManager;
    const total = new Set(Spns.extractSiteAppPermissions(sm.sitePerms).map((r) => r.permissionId)).size;
    $('sitePermsSelCount').textContent = String(sm.selectedPerms.size);
    $('sitePermsRemoveBtn').disabled = !sm.readable || sm.selectedPerms.size === 0;
    setAllBox($('sitePermsAll'), sm.selectedPerms.size, sm.readable ? total : 0);
  }

  function syncSitePermBoxes() {
    const sel = state.siteManager.selectedPerms;
    document.querySelectorAll('#sitePermsTable tbody input[data-act="sel-site-perm"]').forEach((x) => { x.checked = sel.has(x.dataset.perm); });
    updateSitePermSelection();
  }
  function onSitePermCheck(e) {
    const cb = e.target.closest('input[data-act="sel-site-perm"]');
    if (!cb) return;
    const sel = state.siteManager.selectedPerms;
    if (cb.checked) sel.add(cb.dataset.perm); else sel.delete(cb.dataset.perm);
    syncSitePermBoxes();
  }
  function onAllSitePerms(e) {
    const sm = state.siteManager;
    sm.selectedPerms = new Set(e.target.checked ? Spns.extractSiteAppPermissions(sm.sitePerms).map((r) => r.permissionId) : []);
    syncSitePermBoxes();
  }

  function renderRolePicker() {
    const sm = state.siteManager;
    $('siteRolePicker').innerHTML = Spns.SITE_ROLE_OPTIONS.map((opt) => `
      <label class="checkbox role-option" title="${esc(opt.hint)}${opt.extended ? ' (not supported in every tenant)' : ''}">
        <input type="checkbox" data-role="${esc(opt.value)}" ${sm.selectedRoles.has(opt.value) ? 'checked' : ''}/>${esc(opt.label)}
      </label>`).join('');
  }

  function currentSitePerm() {
    return Spns.findAppPermissionOnSite(state.siteManager.sitePerms, state.selectedSpn.appId);
  }

  function renderGrantBox() {
    const sm = state.siteManager;
    const existing = currentSitePerm();
    sm.selectedRoles = new Set(existing ? existing.roles : []);
    $('siteGrantTitle').textContent = existing ? `Change access of ${spnName()}` : `Grant access to ${spnName()}`;
    $('siteRemoveBtn').hidden = !existing;
    $('siteGrantHint').textContent = existing ? `Now: ${existing.roles.join(', ') || '(none)'}` : 'No access to this site yet.';
    renderRolePicker();
    updateGrantButtonState();
  }

  function onSiteRoleToggle(e) {
    const cb = e.target.closest('input[data-role]');
    if (!cb) return;
    const sm = state.siteManager;
    if (cb.checked) sm.selectedRoles.add(cb.dataset.role); else sm.selectedRoles.delete(cb.dataset.role);
    updateGrantButtonState();
  }

  function updateGrantButtonState() {
    const existing = currentSitePerm();
    const roles = [...state.siteManager.selectedRoles];
    const btn = $('siteGrantBtn');
    btn.textContent = existing ? 'Change roles' : 'Grant';
    let changed = roles.length > 0;
    if (existing) {
      const cur = new Set(existing.roles);
      changed = roles.length > 0 && (roles.length !== cur.size || roles.some((r) => !cur.has(r)));
    }
    btn.disabled = !changed || !state.siteManager.readable;
  }

  // ---------- site changes ----------
  // After every change: re-read, then offer to drop temporary admin rights.
  async function afterSiteChange() {
    const sm = state.siteManager;
    const site = sm.selectedSite;
    try {
      sm.sitePerms = await Graph.getSitePermissions(site.id);
      setAccessFromPerms(sm.selectedKey, sm.sitePerms);
    } catch (e) { Log.warn(`Re-reading site permissions failed: ${e.message}`); }
    renderSitePerms();
    renderSiteResults();
    updateSiteStatus();
    await offerDemotion(site);
  }

  // A 403 on a change means you are not (or no longer) site admin: offer to become one.
  async function onSiteChangeFailed(e, what) {
    Log.err(`${what} failed:`, e);
    const sm = state.siteManager;
    if (!Spns.isAccessDenied(e)) {
      setNotice($('siteGrantFeedback'), `${what} failed: ${e.message}`, 'err');
      toast(`${what} failed: ${e.message}`, 'err');
      return;
    }
    const site = sm.selectedSite;
    const token = sm.detailToken;
    sm.readable = false;
    renderSitePerms();
    setSiteState('denied', 'SharePoint refused the change because you are not an admin of this site.');
    if (!(await elevate(site)) || token !== sm.detailToken) return;
    await loadSiteDetail(site, token, { waitForAccess: true });
    if (sm.readable) setNotice($('siteGrantFeedback'), 'You are site admin now and the permissions are reloaded. Check them and repeat the change.', 'info');
  }

  async function onGrantSiteClick() {
    const sm = state.siteManager;
    const sp = state.selectedSpn;
    const site = sm.selectedSite;
    const roles = [...sm.selectedRoles];
    if (!site || !sm.readable || !roles.length || !Spns.isValidSiteId(site.id)) return;
    const existing = currentSitePerm();
    const ok = await confirmChange({
      title: existing ? 'Change site access?' : 'Grant site access?',
      site,
      items: existing
        ? [`Change roles on ${siteName(site)} from ${existing.roles.join(', ') || '(none)'} to ${roles.join(', ')}`]
        : roles.map((r) => `Add site role ${r} on ${siteName(site)}`),
      note: 'Application access to this one site. The app uses it without a signed-in user.',
      noteKind: 'info',
      okLabel: existing ? 'Change' : 'Grant',
    });
    if (!ok) return;
    clearNotice($('siteGrantFeedback'));
    await busy($('siteGrantBtn'), async () => {
      try {
        if (existing) await Graph.updateSitePermission(site.id, existing.permissionId, { roles });
        else await Graph.addSitePermission(site.id, { appId: sp.appId, displayName: sp.displayName || sp.appId, roles });
      } catch (e) { await onSiteChangeFailed(e, existing ? 'Change' : 'Grant'); return; }
      Log.info(`${existing ? 'Changed' : 'Granted'} site access for ${sp.appId} on ${site.id}: ${roles.join(', ')}`);
      setNotice($('siteGrantFeedback'), `${existing ? 'Changed' : 'Granted'}: ${roles.join(', ')}`, 'ok');
      toast(existing ? 'Site access changed' : 'Site access granted', 'ok');
      await afterSiteChange();
    });
    updateGrantButtonState();
  }

  async function onRemoveSiteClick() {
    const sm = state.siteManager;
    const sp = state.selectedSpn;
    const site = sm.selectedSite;
    const existing = currentSitePerm();
    if (!site || !sm.readable || !existing || !Spns.isValidSiteId(site.id)) return;
    const ok = await confirmChange({
      title: 'Remove site access?',
      site,
      items: [`Remove site access (${existing.roles.join(', ') || 'no roles'}) on ${siteName(site)}`],
      note: 'The app loses access to this site. Workloads that rely on it break.',
      okLabel: 'Remove',
      danger: true,
    });
    if (!ok) return;
    await busy($('siteRemoveBtn'), async () => {
      try { await Graph.deleteSitePermission(site.id, existing.permissionId); }
      catch (e) { await onSiteChangeFailed(e, 'Remove'); return; }
      Log.info(`Removed site access for ${sp.appId} on ${site.id}`);
      toast('Site access removed', 'ok');
      await afterSiteChange();
    });
  }

  // Removes ticked grants of any app on the open site, not only the selected SPN's.
  async function onRemoveSelectedSitePermsClick() {
    const sm = state.siteManager;
    const site = sm.selectedSite;
    const picks = Spns.sitePermsForRemoval(sm.sitePerms, sm.selectedPerms);
    if (!site || !sm.readable || !picks.length || !Spns.isValidSiteId(site.id)) return;
    const appText = (p) => p.apps.map((a) => `${a.displayName || '(unknown app)'} (appId ${a.appId})`).join(', ');
    const ok = await confirmChange({
      title: `Remove ${plural(picks.length, 'app permission')} on this site?`,
      site,
      showSpn: false,
      items: picks.map((p) => `Remove ${appText(p)}: ${p.roles.join(', ') || 'no roles'}`),
      note: 'These apps lose their access to this site. Workloads that rely on it break.',
      okLabel: `Remove ${picks.length}`,
      danger: true,
    });
    if (!ok) return;
    await busy($('sitePermsRemoveBtn'), async () => {
      let done = 0;
      for (const p of picks) {
        try {
          await Graph.deleteSitePermission(site.id, p.permissionId);
          done++;
          Log.info(`Removed site permission ${p.permissionId} (${appText(p)}) on ${site.id}`);
        } catch (e) {
          if (Spns.isAccessDenied(e)) { await onSiteChangeFailed(e, 'Remove'); return; }
          Log.err(`Removing ${appText(p)} failed:`, e);
          toast(`Removing ${appText(p)} failed: ${e.message}`, 'err');
        }
      }
      if (done) toast(`Removed ${plural(done, 'app permission')}${done < picks.length ? `, ${picks.length - done} failed` : ''}`, done === picks.length ? 'ok' : 'warn');
      sm.selectedPerms.clear();
      await afterSiteChange();
    });
  }

  // ---------- temporary site admin ----------
  const ELEVATION_KEY = 'SPNRoleMgr.siteAdmin';
  function loadElevations() {
    try {
      const v = JSON.parse(localStorage.getItem(ELEVATION_KEY) || '[]');
      state.elevations = Array.isArray(v) ? v : [];
    } catch { state.elevations = []; }
  }
  function saveElevations(list) {
    state.elevations = list;
    try { localStorage.setItem(ELEVATION_KEY, JSON.stringify(list)); }
    catch (e) { Log.warn(`Could not store the site admin record: ${e.message}`); }
  }
  function myElevations() {
    const t = Auth.tenantInfo();
    return t ? Spns.elevationsFor(state.elevations, t.tenantId, t.upn) : [];
  }
  function isElevated(url) {
    const k = Spns.siteKey(url);
    return !!k && myElevations().some((r) => r.siteUrl === k);
  }

  function renderElevations() {
    const list = myElevations();
    $('elevationBanner').hidden = list.length === 0;
    $('elevationList').replaceChildren(...list.map((r) => h('li', { text: `${r.siteUrl}${r.at ? `, since ${new Date(r.at).toLocaleString()}` : ''}` })));
    const sm = state.siteManager;
    $('siteElevatedNote').hidden = !(sm.selectedSite && isElevated(sm.selectedSite.webUrl));
  }

  async function elevate(site) {
    const t = Auth.tenantInfo();
    const url = keyOf(site);
    if (!t || !t.upn || !url) return false;
    const ok = await confirmChange({
      title: 'Make yourself site admin temporarily?',
      site: { displayName: siteName(site), webUrl: url },
      showSpn: false,
      items: [`Add ${t.upn} as site collection admin of ${url}`],
      note: 'This goes through the SharePoint admin API and needs SharePoint Administrator. It shows in the audit log. After each change you are asked whether to remove it again.',
      okLabel: 'Make me site admin',
    });
    if (!ok) return false;
    // On record before the call, so a closed tab never hides a grant.
    saveElevations(Spns.addElevation(state.elevations, { siteUrl: url, upn: t.upn, tenantId: t.tenantId, at: new Date().toISOString() }));
    try {
      await Spo.setSiteAdmin(url, t.upn, true);
      Log.info(`Made ${t.upn} site collection admin of ${url}`);
      toast('You are site admin now', 'ok');
      return true;
    } catch (e) {
      saveElevations(Spns.dropElevation(state.elevations, url, t.tenantId, t.upn));
      Log.err('SetSiteAdmin failed:', e);
      toast(Spns.isAccessDenied(e)
        ? `The SharePoint admin API refused it. This needs the SharePoint Administrator role. (${e.message})`
        : `Could not make you site admin: ${e.message}`, 'err');
      return false;
    } finally {
      renderElevations();
    }
  }

  async function demote(urls) {
    const t = Auth.tenantInfo();
    if (!t || !urls.length) return;
    let failed = 0;
    for (const url of urls) {
      try {
        await Spo.setSiteAdmin(url, t.upn, false);
        saveElevations(Spns.dropElevation(state.elevations, url, t.tenantId, t.upn));
        Log.info(`Removed ${t.upn} as site collection admin of ${url}`);
      } catch (e) {
        failed++;
        Log.err(`Removing site admin on ${url} failed:`, e);
        toast(`Could not remove your admin rights on ${url}: ${e.message}`, 'err');
      }
    }
    renderElevations();
    if (!failed) toast(urls.length === 1 ? 'Your temporary site admin rights are removed' : `Your admin rights on ${urls.length} sites are removed`, 'ok');
    const sm = state.siteManager;
    if (modalOpen() && sm.selectedSite && urls.includes(keyOf(sm.selectedSite)) && !isElevated(sm.selectedSite.webUrl)) {
      sm.readable = false;
      renderSitePerms();
      setSiteState('denied', 'You are no longer admin of this site. The list shows it as it was after your last change. Make yourself admin again for more changes.');
    }
  }

  async function offerDemotion(site) {
    if (!site || !isElevated(site.webUrl)) return;
    const ok = await Toolkit.confirm({
      title: 'Remove your temporary site admin rights?',
      body: `You were made site collection admin of ${siteName(site)} for this change. Remove that now, or keep it for more changes on this site. You are asked again after the next change.`,
      okLabel: 'Remove my admin rights',
      cancelLabel: 'Keep for now',
    });
    if (ok) await demote([keyOf(site)]);
  }

  async function onElevateClick() {
    const sm = state.siteManager;
    const site = sm.selectedSite;
    if (!site) return;
    const token = sm.detailToken;
    await busy($('siteElevateBtn'), async () => {
      if (!(await elevate(site)) || token !== sm.detailToken) return;
      await loadSiteDetail(site, token, { waitForAccess: true });
    });
  }

  // ---------- export ----------
  const SPN_COLUMNS = [
    { key: 'displayName', label: 'Display name' },
    { key: 'type', label: 'Type' },
    { key: 'appId', label: 'AppId' },
    { key: 'id', label: 'Object id' },
    { key: 'accountEnabled', label: 'Enabled' },
    { key: 'servicePrincipalType', label: 'Service principal type' },
    { key: 'appOwnerOrganizationId', label: 'Owner tenant' },
    { key: 'publisherName', label: 'Publisher' },
    { key: 'createdDateTime', label: 'Created' },
  ];
  const spnExportRows = () => sortSpns(state.filteredSpns).map((sp) => ({ ...sp, type: Spns.classifySpn(sp) }));

  const PERM_COLUMNS = [
    { key: 'kind', label: 'Kind' },
    { key: 'api', label: 'API' },
    { key: 'apiAppId', label: 'API appId' },
    { key: 'permission', label: 'Permission' },
    { key: 'consentType', label: 'Consent type' },
    { key: 'principalId', label: 'Principal' },
    { key: 'granted', label: 'Granted' },
    { key: 'id', label: 'Assignment or grant id' },
  ];
  function permExportRows() {
    const res = (id) => state.resourceCache.get(id) || {};
    return [
      ...state.appRoles.map((a) => {
        const meta = (state.resourceIndexCache.get(a.resourceId) || new Map()).get(a.appRoleId);
        return { kind: 'Application', api: res(a.resourceId).displayName || a.resourceDisplayName, apiAppId: res(a.resourceId).appId, permission: meta ? meta.name : a.appRoleId, granted: a.createdDateTime, id: a.id };
      }),
      ...state.oauth2Grants.map((g) => ({ kind: 'Delegated', api: res(g.resourceId).displayName || g.resourceId, apiAppId: res(g.resourceId).appId, permission: g.scope, consentType: g.consentType, principalId: g.principalId, id: g.id })),
    ];
  }
  function fileSafe(name) { return String(name || 'spn').replace(/[^\w.-]+/g, '-').substring(0, 60); }

  // ---------- boot ----------
  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    activateTab = Toolkit.tabs('#tabs');
    $('signInBtn').addEventListener('click', onSignIn);
    $('loadSpnsBtn').addEventListener('click', loadSpns);
    $('spnFilter').addEventListener('input', applySpnFilter);
    $('typeFilter').addEventListener('change', applySpnFilter);
    $('hideMs').addEventListener('change', applySpnFilter);
    document.querySelector('#spnTable thead').addEventListener('click', onSpnSortClick);
    document.querySelector('#spnTable tbody').addEventListener('click', onSpnTableClick);
    document.querySelector('#appRolesTable tbody').addEventListener('change', onRoleCheck);
    document.querySelector('#oauth2Table tbody').addEventListener('change', onScopeCheck);
    $('appRolesAll').addEventListener('change', onAllRoles);
    $('oauth2All').addEventListener('change', onAllScopes);
    $('revokeAppsBtn').addEventListener('click', onRevokeAppsClick);
    $('revokeScopesBtn').addEventListener('click', onRevokeScopesClick);
    $('revokeAllBtn').addEventListener('click', onRevokeAllClick);
    $('refreshPermsBtn').addEventListener('click', () => busy($('refreshPermsBtn'), loadPermissions));
    $('manageSitesBtn').addEventListener('click', onManageSitesClick);

    $('siteModalCloseBtn').addEventListener('click', closeSiteManager);
    $('sitePickerModal').addEventListener('mousedown', (e) => { if (e.target === $('sitePickerModal')) closeSiteManager(); });
    // Escape in a confirm on top of the site manager is handled (and prevented) by that confirm.
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !e.defaultPrevented && !$('sitePickerModal').hidden) closeSiteManager();
    });
    $('siteSearch').addEventListener('input', onSiteSearchInput);
    $('siteResults').addEventListener('click', onSiteResultClick);
    $('siteResults').addEventListener('keydown', onSiteResultClick);
    $('siteRolePicker').addEventListener('change', onSiteRoleToggle);
    $('siteGrantBtn').addEventListener('click', onGrantSiteClick);
    $('siteRemoveBtn').addEventListener('click', onRemoveSiteClick);
    $('siteLoadAllBtn').addEventListener('click', () => busy($('siteLoadAllBtn'), loadAllSites));
    $('siteCheckAllBtn').addEventListener('click', () => scanAccess(true));
    document.querySelector('#sitePermsTable tbody').addEventListener('change', onSitePermCheck);
    $('sitePermsAll').addEventListener('change', onAllSitePerms);
    $('sitePermsRemoveBtn').addEventListener('click', onRemoveSelectedSitePermsClick);
    $('siteElevateBtn').addEventListener('click', onElevateClick);
    $('siteDemoteBtn').addEventListener('click', () => busy($('siteDemoteBtn'), () => demote([keyOf(state.siteManager.selectedSite)])));
    $('elevationRemoveBtn').addEventListener('click', () => busy($('elevationRemoveBtn'), () => demote(myElevations().map((r) => r.siteUrl))));
    // A temporary site admin grant must not be forgotten by closing the tab.
    window.addEventListener('beforeunload', (e) => {
      if (!myElevations().length) return;
      e.preventDefault();
      e.returnValue = '';
    });

    const apiSearch = $('apiSearch');
    apiSearch.addEventListener('change', () => loadApiForQuery(apiSearch.value));
    apiSearch.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      loadApiForQuery(apiSearch.value);
    });
    document.querySelector('#addPermTable tbody').addEventListener('change', onAddPermTableChange);
    document.querySelectorAll('input[name="permKind"]').forEach((r) => r.addEventListener('change', (e) => {
      state.addApiKind = e.target.value;
      state.addSelected.clear();
      renderAddPermTable();
    }));
    $('permFilter').addEventListener('input', renderAddPermTable);
    $('addPermsBtn').addEventListener('click', onAddPermsClick);

    $('exportSpnXlsxBtn').addEventListener('click', () => busy($('exportSpnXlsxBtn'), () => Export.xlsx([{ name: 'Service principals', rows: spnExportRows(), columns: SPN_COLUMNS }], `service-principals-${Export.stamp()}.xlsx`).catch((e) => toast(e.message, 'err'))));
    $('exportSpnCsvBtn').addEventListener('click', () => Export.csv(spnExportRows(), { columns: SPN_COLUMNS, filename: `service-principals-${Export.stamp()}.csv` }));
    $('exportPermsBtn').addEventListener('click', () => busy($('exportPermsBtn'), () => Export.xlsx([{ name: 'Permissions', rows: permExportRows(), columns: PERM_COLUMNS }], `permissions-${fileSafe(state.selectedSpn && state.selectedSpn.displayName)}-${Export.stamp()}.xlsx`).catch((e) => toast(e.message, 'err'))));

    renderManageTab();
    try {
      await Auth.init();
      const t = Auth.tenantInfo();
      Graph.pinTenant(t && t.tenantId);
      refreshAuthUI();
      loadElevations();
      renderElevations();
      if (myElevations().length) Log.warn(`Still temporary site admin on: ${myElevations().map((r) => r.siteUrl).join(', ')}`);
      Log.info('SPNRoleMgr ready.');
    } catch (e) {
      Log.err('Start-up failed:', e);
      toast(`Start-up failed: ${e.message}`, 'err');
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
