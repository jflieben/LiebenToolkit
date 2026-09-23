// SPNRoleMgr - SPA wiring.
// Pattern follows DupedDevices: small global state, render functions per panel,
// confirm() before any mutation, dry-run by inspection (you only ever execute
// after picking individual rows + clicking the explicit grant/revoke button).
(() => {
  const state = {
    spns: [],            // all loaded service principals
    filteredSpns: [],
    apiSpns: [],         // SPNs that expose assignable API permissions
    selectedSpn: null,   // SPN object the user is managing
    appRoles: [],        // current appRoleAssignments on selected SPN
    oauth2Grants: [],    // current oauth2PermissionGrants on selected SPN
    resourceCache: new Map(), // resourceObjectId -> resource SPN with appRoles + scopes
    resourceIndexCache: new Map(), // resourceObjectId -> permission index
    addApiSpn: null,     // resource SPN currently chosen on the Add panel
    addApiKind: 'AppRole',
    addSelected: new Set(),
    sort: { col: 'name', dir: 'asc' },
    cancelToken: null,
    siteManager: {       // SharePoint site-access modal (Sites.Selected)
      sites: [],
      selectedSite: null,
      sitePerms: [],
      selectedRoles: new Set(),
      searchToken: 0,
      optimisticUnlock: false, // set right after granting Sites.Selected (eventual-consistency bridge)
    },
  };

  // --------- THEME ---------
  function initTheme() {
    const saved = localStorage.getItem('spnrolemgr-theme') || 'light';
    document.documentElement.setAttribute('data-theme', saved);
  }
  function toggleTheme() {
    const cur = document.documentElement.getAttribute('data-theme');
    const nxt = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', nxt);
    localStorage.setItem('spnrolemgr-theme', nxt);
  }

  // --------- TABS ---------
  function activateTab(id) {
    document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === id));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `tab-${id}`));
  }

  // --------- TOAST + CONFIRM MODAL ---------
  function toast(msg, kind = 'info') {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.className = `toast toast-${kind}`;
    el.classList.remove('hidden');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.add('hidden'), 4000);
  }
  function confirmModal(title, htmlBody, okLabel = 'Confirm') {
    return new Promise(resolve => {
      document.getElementById('confirmTitle').textContent = title;
      document.getElementById('confirmBody').innerHTML = htmlBody;
      const ok = document.getElementById('confirmOkBtn');
      ok.textContent = okLabel;
      const modal = document.getElementById('confirmModal');
      modal.classList.remove('hidden');
      const cleanup = (val) => {
        modal.classList.add('hidden');
        ok.removeEventListener('click', okFn);
        document.getElementById('confirmCancelBtn').removeEventListener('click', cancelFn);
        document.getElementById('confirmCloseBtn').removeEventListener('click', cancelFn);
        resolve(val);
      };
      const okFn = () => cleanup(true);
      const cancelFn = () => cleanup(false);
      ok.addEventListener('click', okFn);
      document.getElementById('confirmCancelBtn').addEventListener('click', cancelFn);
      document.getElementById('confirmCloseBtn').addEventListener('click', cancelFn);
    });
  }

  // --------- AUTH UI ---------
  async function refreshAuthUI() {
    const acc = Auth.getAccount();
    const userBox = document.getElementById('userBox');
    const userName = document.getElementById('userName');
    const connectCard = document.getElementById('connectCard');
    const browseCard = document.getElementById('browseCard');
    if (acc) {
      userBox.classList.remove('hidden');
      userName.textContent = acc.username || acc.name || 'Signed in';
      connectCard.classList.add('hidden');
      browseCard.classList.remove('hidden');
    } else {
      userBox.classList.add('hidden');
      connectCard.classList.remove('hidden');
      browseCard.classList.add('hidden');
    }
  }
  async function onSignIn() {
    try {
      await Auth.signIn(false);
      await refreshAuthUI();
    } catch (e) { Log.err('Sign-in failed:', e); toast('Sign-in failed: ' + e.message, 'err'); }
  }
  async function onSignOut() {
    await Auth.signOut();
    state.spns = []; state.filteredSpns = []; state.apiSpns = []; state.selectedSpn = null;
    renderSpnTable(); renderManageTab();
    await refreshAuthUI();
  }

  // --------- LOAD SPNs ---------
  async function loadSpns() {
    const btn = document.getElementById('loadSpnsBtn');
    btn.disabled = true;
    const status = document.getElementById('loadStatus');
    const box = document.getElementById('progressBox');
    const fill = document.getElementById('progressFill');
    const count = document.getElementById('progressCount');
    const label = document.getElementById('progressLabel');
    box.classList.remove('hidden');
    label.textContent = 'Loading service principals...';
    fill.style.width = '0%';
    count.textContent = '0';
    state.cancelToken = new Concurrency.CancelToken();
    document.getElementById('cancelBtn').onclick = () => state.cancelToken.cancel();
    const t0 = Date.now();
    try {
      const list = await Graph.listAllServicePrincipals((loaded, page) => {
        if (state.cancelToken.cancelled) throw new Error('Cancelled');
        count.textContent = `${loaded} loaded (${page} pages)`;
        // We don't know the total; use indeterminate-ish fill that grows with pages.
        fill.style.width = Math.min(95, page * 8) + '%';
      });
      fill.style.width = '100%';
      state.spns = list;
      state.filteredSpns = list.slice();
      state.apiSpns = sortApiSpns(list.filter(Spns.hasAssignablePermissions));
      const dt = ((Date.now() - t0) / 1000).toFixed(1);
      status.textContent = `Loaded ${list.length} service principals in ${dt}s`;
      Log.info(`Loaded ${list.length} service principals in ${dt}s`);
      populateTypeFilter();
      applySpnFilter();
    } catch (e) {
      Log.err('Load failed:', e);
      toast('Load failed: ' + e.message, 'err');
    } finally {
      btn.disabled = false;
      setTimeout(() => box.classList.add('hidden'), 600);
    }
  }

  function populateTypeFilter() {
    const sel = document.getElementById('typeFilter');
    const seen = new Set();
    for (const sp of state.spns) seen.add(Spns.classifySpn(sp));
    sel.innerHTML = '<option value="">All types</option>' + [...seen].sort().map(t => `<option value="${escAttr(t)}">${escHtml(t)}</option>`).join('');
  }

  function applySpnFilter() {
    const q = document.getElementById('spnFilter').value.trim().toLowerCase();
    const type = document.getElementById('typeFilter').value;
    const hideMs = document.getElementById('hideMs').checked;
    state.filteredSpns = state.spns.filter(sp => {
      if (hideMs && Spns.isMicrosoftFirstParty(sp)) return false;
      if (type && Spns.classifySpn(sp) !== type) return false;
      if (!q) return true;
      const hay = `${sp.displayName || ''}\n${sp.appId || ''}\n${sp.appDisplayName || ''}`.toLowerCase();
      return hay.includes(q);
    });
    renderSpnTable();
  }

  function sortSpns(rows) {
    const { col, dir } = state.sort;
    const m = dir === 'asc' ? 1 : -1;
    const k = sp => {
      switch (col) {
        case 'name': return (sp.displayName || '').toLowerCase();
        case 'type': return Spns.classifySpn(sp);
        case 'appId': return sp.appId || '';
        case 'enabled': return sp.accountEnabled ? 1 : 0;
        default: return '';
      }
    };
    return rows.slice().sort((a, b) => k(a) > k(b) ? m : k(a) < k(b) ? -m : 0);
  }

  function renderSpnTable() {
    const wrap = document.getElementById('spnTableWrap');
    const tbody = document.querySelector('#spnTable tbody');
    const hint = document.getElementById('spnHint');
    const filterRow = document.getElementById('spnFilterRow');
    if (!state.spns.length) {
      wrap.classList.add('hidden');
      filterRow.classList.add('hidden');
      hint.classList.add('hidden');
      return;
    }
    filterRow.classList.remove('hidden');
    wrap.classList.remove('hidden');
    hint.classList.remove('hidden');
    const rows = sortSpns(state.filteredSpns).slice(0, 1000);
    hint.textContent = `${state.filteredSpns.length} match${state.filteredSpns.length === 1 ? '' : 'es'}` + (state.filteredSpns.length > rows.length ? ` (showing first ${rows.length})` : '');
    tbody.innerHTML = rows.map(sp => `
        <tr data-id="${escAttr(sp.id)}" class="spn-row">
        <td><strong>${escHtml(sp.displayName || '(no name)')}</strong></td>
        <td>${escHtml(Spns.classifySpn(sp))}</td>
        <td><code>${escHtml(sp.appId || '')}</code></td>
        <td>${sp.accountEnabled ? 'Yes' : 'No'}</td>
      </tr>
    `).join('');
    // Highlight sort column
    document.querySelectorAll('#spnTable thead th').forEach(th => {
      th.classList.remove('sort-asc','sort-desc','sortable-col');
      if (th.dataset.sort) {
        th.classList.add('sortable-col');
        if (th.dataset.sort === state.sort.col) th.classList.add(state.sort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
      }
    });
  }

  function onSpnTableClick(e) {
    const tr = e.target.closest('tr[data-id]');
    if (!tr) return;
    const sp = state.spns.find(s => s.id === tr.dataset.id);
    if (sp) selectSpn(sp);
  }
  function onSpnSortClick(e) {
    const th = e.target.closest('th[data-sort]');
    if (!th) return;
    const col = th.dataset.sort;
    if (state.sort.col === col) state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
    else { state.sort.col = col; state.sort.dir = 'asc'; }
    renderSpnTable();
  }

  // --------- SELECT + LOAD PERMISSIONS ---------
  async function selectSpn(sp) {
    state.selectedSpn = sp;
    state.siteManager.optimisticUnlock = false;
    activateTab('manage');
    renderManageTab();
    await loadPermissions();
  }

  async function loadPermissions() {
    if (!state.selectedSpn) return;
    const sp = state.selectedSpn;
    document.getElementById('appRolesCard').classList.remove('hidden');
    document.getElementById('oauth2Card').classList.remove('hidden');
    document.getElementById('addCard').classList.remove('hidden');
    document.querySelector('#appRolesTable tbody').innerHTML = '<tr><td colspan="5" class="muted">Loading...</td></tr>';
    document.querySelector('#oauth2Table tbody').innerHTML = '<tr><td colspan="5" class="muted">Loading...</td></tr>';
    try {
      const [appRoles, oauth2] = await Promise.all([
        Graph.listAppRoleAssignments(sp.id),
        Graph.listOAuth2Grants(sp.id),
      ]);
      state.appRoles = appRoles;
      state.oauth2Grants = oauth2;

      // Pre-fetch resource SPNs referenced (so we can resolve permission names).
      const resourceIds = new Set();
      for (const r of appRoles) if (r.resourceId) resourceIds.add(r.resourceId);
      for (const g of oauth2) if (g.resourceId) resourceIds.add(g.resourceId);
      await Promise.all([...resourceIds].map(id => ensureResource(id)));

      renderAppRoles();
      renderOAuth2();
      renderAddPanel();
      renderManageTab();
      await autoSelectDefaultApi();
    } catch (e) {
      Log.err('Load permissions failed:', e);
      toast('Load permissions failed: ' + e.message, 'err');
    }
  }

  async function ensureResource(resourceObjectId) {
    if (state.resourceCache.has(resourceObjectId)) return state.resourceCache.get(resourceObjectId);
    try {
      const sp = await Graph.getServicePrincipal(resourceObjectId);
      state.resourceCache.set(resourceObjectId, sp);
      state.resourceIndexCache.set(resourceObjectId, Spns.buildPermissionIndex(sp));
      return sp;
    } catch (e) {
      Log.warn(`Could not resolve resource SPN ${resourceObjectId}: ${e.message}`);
      state.resourceCache.set(resourceObjectId, null);
      state.resourceIndexCache.set(resourceObjectId, new Map());
      return null;
    }
  }

  function renderManageTab() {
    const sp = state.selectedSpn;
    const noSel = document.getElementById('noSelectionCard');
    const head = document.getElementById('spnHeaderCard');
    if (!sp) {
      noSel.classList.remove('hidden');
      head.classList.add('hidden');
      document.getElementById('appRolesCard').classList.add('hidden');
      document.getElementById('oauth2Card').classList.add('hidden');
      document.getElementById('addCard').classList.add('hidden');
      return;
    }
    noSel.classList.add('hidden');
    head.classList.remove('hidden');
    document.getElementById('selSpnName').textContent = sp.displayName || '(no name)';
    document.getElementById('selSpnMeta').innerHTML =
      `<code>${escHtml(sp.appId || '')}</code> &middot; ${escHtml(Spns.classifySpn(sp))}` +
      (sp.publisherName ? ` &middot; ${escHtml(sp.publisherName)}` : '') +
      ` &middot; Enabled: ${sp.accountEnabled ? 'Yes' : 'No'}`;
    document.getElementById('revokeAllBtn').disabled = !state.appRoles.length && !state.oauth2Grants.length;
    updateManageSitesButton();
  }

  // --------- SHAREPOINT SITE ACCESS (Sites.Selected) ---------
  function canManageSites() {
    return state.siteManager.optimisticUnlock ||
      Spns.hasSitesSelected(state.appRoles, state.resourceCache, state.resourceIndexCache);
  }

  function updateManageSitesButton() {
    const btn = document.getElementById('manageSitesBtn');
    if (!btn) return;
    const canManage = canManageSites();
    btn.disabled = !canManage;
    btn.title = canManage
      ? 'Grant this service principal access to specific SharePoint sites (Sites.Selected model).'
      : 'This service principal needs the Sites.Selected application permission (Microsoft Graph or SharePoint Online) first. Add it from the "Add permission" panel further down this page.';
  }

  function renderAppRoles() {
    const sp = state.selectedSpn;
    const tbody = document.querySelector('#appRolesTable tbody');
    const empty = document.getElementById('appRolesEmpty');
    document.getElementById('appRolesCount').textContent = `(${state.appRoles.length})`;
    if (!state.appRoles.length) {
      tbody.innerHTML = '';
      empty.classList.remove('hidden');
      return;
    }
    empty.classList.add('hidden');
    tbody.innerHTML = state.appRoles.map(a => {
      const res = state.resourceCache.get(a.resourceId);
      const idx = state.resourceIndexCache.get(a.resourceId) || new Map();
      const meta = idx.get(a.appRoleId);
      const apiName = (res && res.displayName) || a.resourceDisplayName || a.resourceId;
      const permName = meta ? meta.name : a.appRoleId;
      const desc = meta ? meta.description : '';
      const granted = a.createdDateTime ? new Date(a.createdDateTime).toLocaleString() : '';
      return `<tr data-id="${escAttr(a.id)}">
        <td>${escHtml(apiName)}</td>
        <td><code>${escHtml(permName)}</code></td>
        <td class="muted small">${escHtml(desc)}</td>
        <td class="muted small">${escHtml(granted)}</td>
        <td><button class="ghost-btn danger-text" data-act="revoke-app">Revoke</button></td>
      </tr>`;
    }).join('');
  }

  async function onRevokeAppClick(e) {
    const btn = e.target.closest('button[data-act="revoke-app"]');
    if (!btn) return;
    const tr = btn.closest('tr');
    const a = state.appRoles.find(x => x.id === tr.dataset.id);
    if (!a) return;
    const idx = state.resourceIndexCache.get(a.resourceId) || new Map();
    const meta = idx.get(a.appRoleId);
    const ok = await confirmModal(
      'Revoke application permission?',
      `<p>This will permanently remove the <code>${escHtml(meta ? meta.name : a.appRoleId)}</code> application permission from <strong>${escHtml(state.selectedSpn.displayName)}</strong>.</p>
       <p>The service principal may stop working if it relies on this permission.</p>`,
      'Revoke'
    );
    if (!ok) return;
    btn.disabled = true;
    try {
      await Graph.removeAppRoleAssignment(state.selectedSpn.id, a.id);
      Log.info(`Revoked app role ${a.appRoleId} from ${state.selectedSpn.id}`);
      toast('Permission revoked', 'ok');
      state.siteManager.optimisticUnlock = false;
      await loadPermissions();
    } catch (e) {
      Log.err('Revoke failed:', e);
      toast('Revoke failed: ' + e.message, 'err');
      btn.disabled = false;
    }
  }

  function renderOAuth2() {
    const tbody = document.querySelector('#oauth2Table tbody');
    const empty = document.getElementById('oauth2Empty');
    document.getElementById('oauth2Count').textContent = `(${state.oauth2Grants.length})`;
    if (!state.oauth2Grants.length) {
      tbody.innerHTML = '';
      empty.classList.remove('hidden');
      return;
    }
    empty.classList.add('hidden');
    // Each grant can hold multiple scopes; render one row per scope so the user can revoke individual scopes.
    const rows = [];
    for (const g of state.oauth2Grants) {
      const res = state.resourceCache.get(g.resourceId);
      const apiName = (res && res.displayName) || g.resourceId;
      const scopes = Spns.parseScopeString(g.scope);
      if (!scopes.length) {
        rows.push({ apiName, grant: g, scope: '(none)' });
      } else {
        for (const s of scopes) rows.push({ apiName, grant: g, scope: s });
      }
    }
    tbody.innerHTML = rows.map((r, i) => `
      <tr data-grant="${escAttr(r.grant.id)}" data-scope="${escAttr(r.scope)}">
        <td>${escHtml(r.apiName)}</td>
        <td><code>${escHtml(r.scope)}</code></td>
        <td>${escHtml(r.grant.consentType)}</td>
        <td class="muted small">${escHtml(r.grant.principalId || (r.grant.consentType === 'AllPrincipals' ? '(all users)' : '(unknown)'))}</td>
        <td><button class="ghost-btn danger-text" data-act="revoke-scope">Revoke</button></td>
      </tr>
    `).join('');
  }

  function sortApiSpns(spns) {
    return spns.slice().sort((a, b) => {
      const an = (a.displayName || '').toLowerCase();
      const bn = (b.displayName || '').toLowerCase();
      if (an > bn) return 1;
      if (an < bn) return -1;
      return (a.appId || '').localeCompare(b.appId || '');
    });
  }

  function resolveApiSpn(q) {
    const needle = String(q || '').trim().toLowerCase();
    if (!needle) return null;
    let sp = state.apiSpns.find(x => x.appId && x.appId.toLowerCase() === needle);
    if (sp) return sp;
    sp = state.apiSpns.find(x => (x.displayName || '').toLowerCase() === needle);
    if (sp) return sp;
    return state.apiSpns.find(x => (x.displayName || '').toLowerCase().includes(needle)) || null;
  }

  async function ensureApiSpnDetail(sp) {
    if (!sp) return null;
    if (Array.isArray(sp.appRoles) && Array.isArray(sp.oauth2PermissionScopes)) return sp;
    const full = await Graph.getServicePrincipal(sp.id);
    return full;
  }

  function setApiSearchValue(sp) {
    if (!sp) return;
    document.getElementById('apiSearch').value = sp.displayName || sp.appId || '';
  }

  async function autoSelectDefaultApi() {
    const graph = state.apiSpns.find(s => (s.displayName || '').toLowerCase() === 'microsoft graph');
    if (!graph) return;
    setApiSearchValue(graph);
    await loadApiForQuery(graph.displayName || graph.appId || '');
  }

  async function onRevokeScopeClick(e) {
    const btn = e.target.closest('button[data-act="revoke-scope"]');
    if (!btn) return;
    const tr = btn.closest('tr');
    const grantId = tr.dataset.grant;
    const scope = tr.dataset.scope;
    const grant = state.oauth2Grants.find(g => g.id === grantId);
    if (!grant) return;
    const remaining = Spns.parseScopeString(grant.scope).filter(s => s !== scope);
    const ok = await confirmModal(
      'Revoke delegated scope?',
      `<p>Remove the <code>${escHtml(scope)}</code> delegated scope from <strong>${escHtml(state.selectedSpn.displayName)}</strong> (${grant.consentType}).</p>` +
      (remaining.length === 0 ? `<p>This is the last scope in the grant, so the entire <code>oauth2PermissionGrant</code> entry will be deleted.</p>` : `<p>Remaining scopes: <code>${escHtml(remaining.join(' '))}</code></p>`),
      'Revoke'
    );
    if (!ok) return;
    btn.disabled = true;
    try {
      if (remaining.length === 0) {
        await Graph.deleteOAuth2Grant(grantId);
      } else {
        await Graph.updateOAuth2Grant(grantId, { scope: Spns.joinScopeString(remaining) });
      }
      Log.info(`Removed scope "${scope}" from grant ${grantId}`);
      toast('Scope revoked', 'ok');
      await loadPermissions();
    } catch (e) {
      Log.err('Revoke scope failed:', e);
      toast('Revoke failed: ' + e.message, 'err');
      btn.disabled = false;
    }
  }

  async function onRevokeAllClick() {
    const sp = state.selectedSpn;
    if (!sp) return;

    const summary = Spns.summarizePermissionRemoval(state.appRoles, state.oauth2Grants);
    if (!summary.totalPermissionEntries) return;

    const ok = await confirmModal(
      'Revoke all permissions?',
      `<p>This will remove <strong>${summary.appRoleCount}</strong> application permission${summary.appRoleCount === 1 ? '' : 's'} and <strong>${summary.grantCount}</strong> delegated grant${summary.grantCount === 1 ? '' : 's'} from <strong>${escHtml(sp.displayName || '(no name)')}</strong>.</p>
       ${summary.delegatedScopeCount ? `<p>The delegated grants currently cover <strong>${summary.delegatedScopeCount}</strong> scope${summary.delegatedScopeCount === 1 ? '' : 's'} in total.</p>` : ''}
       <p>Delegated permissions are deleted per grant, so any scope bundle in the same <code>oauth2PermissionGrant</code> entry is removed together.</p>
       <p class="notice notice-warn">This cannot be undone from SPNRoleMgr. If the workload still needs access, you will have to grant it again later.</p>`,
      'Revoke all'
    );
    if (!ok) return;

    const btn = document.getElementById('revokeAllBtn');
    btn.disabled = true;

    let successCount = 0;
    let failCount = 0;

    for (const assignment of [...state.appRoles]) {
      try {
        await Graph.removeAppRoleAssignment(sp.id, assignment.id);
        successCount += 1;
      } catch (e) {
        Log.err('Bulk revoke app role failed:', e);
        failCount += 1;
      }
    }

    for (const grant of [...state.oauth2Grants]) {
      try {
        await Graph.deleteOAuth2Grant(grant.id);
        successCount += 1;
      } catch (e) {
        Log.err('Bulk revoke delegated grant failed:', e);
        failCount += 1;
      }
    }

    if (successCount > 0) {
      toast(
        failCount
          ? `Removed ${successCount} permission${successCount === 1 ? '' : 's'} with ${failCount} failure${failCount === 1 ? '' : 's'}.`
          : `Removed ${successCount} permission${successCount === 1 ? '' : 's'}.`,
        failCount ? 'warn' : 'ok'
      );
      await loadPermissions();
    } else if (failCount > 0) {
      toast(`Bulk revoke completed with ${failCount} failure${failCount === 1 ? '' : 's'}.`, 'err');
    } else {
      toast('No permissions were removed.', 'err');
    }
  }

  // --------- ADD PERMISSION PANEL ---------
  function renderAddPanel() {
    // Populate API suggestions only from SPNs that expose assignable permissions.
    const dl = document.getElementById('apiSuggestions');
    const seen = new Set();
    const opts = [];
    for (const sp of state.apiSpns) {
      if (!sp.displayName) continue;
      if (seen.has(sp.displayName)) continue;
      seen.add(sp.displayName);
      opts.push(`<option value="${escAttr(sp.displayName)}"></option>`);
        // Removed appId entry from API datalist
    }
    dl.innerHTML = opts.join('');
    state.addApiSpn = null;
    state.addSelected.clear();
    document.querySelector('#addPermTable tbody').innerHTML = '<tr><td colspan="4" class="muted">Pick an API to load permissions.</td></tr>';
    document.getElementById('addSelCount').textContent = '0';
    document.getElementById('addPermsBtn').disabled = true;
    document.getElementById('addResults').classList.add('hidden');
    clearAddFeedback();
    document.getElementById('apiLoadStatus').textContent = `${state.apiSpns.length} API SPN(s) with assignable permissions available.`;
  }

  async function loadApiForQuery(query) {
    const q = String(query || '').trim();
    const status = document.getElementById('apiLoadStatus');
    if (!q) { status.textContent = 'Enter an API name first.'; return; }
    status.textContent = 'Resolving...';
    try {
      const resourceSpn = resolveApiSpn(q);
      if (!resourceSpn) { status.textContent = 'No matching API with assignable permissions found.'; return; }
      const full = await ensureApiSpnDetail(resourceSpn);
      state.addApiSpn = full;
      state.resourceCache.set(full.id, full);
      state.resourceIndexCache.set(full.id, Spns.buildPermissionIndex(full));
      const ar = (full.appRoles || []).length;
      const os = (full.oauth2PermissionScopes || []).length;
      status.textContent = `${full.displayName}: ${ar} app role(s), ${os} delegated scope(s).`;
      state.addSelected.clear();
      renderAddPermTable();
      return full;
    } catch (e) {
      Log.err('Load API failed:', e);
      status.textContent = 'Failed: ' + e.message;
      return null;
    }
  }

  async function onLoadApi() {
    await loadApiForQuery(document.getElementById('apiSearch').value);
  }

  function renderAddPermTable() {
    const tbody = document.querySelector('#addPermTable tbody');
    if (!state.addApiSpn) { tbody.innerHTML = '<tr><td colspan="4" class="muted">Pick an API above.</td></tr>'; return; }
    const filter = document.getElementById('permFilter').value.trim().toLowerCase();
    const isAppRole = state.addApiKind === 'AppRole';
    const items = isAppRole ? (state.addApiSpn.appRoles || []) : (state.addApiSpn.oauth2PermissionScopes || []);
    // For app roles, only those with allowedMemberTypes containing "Application" are assignable to SPNs.
    const filtered = items.filter(it => {
      if (isAppRole && !(it.allowedMemberTypes || []).includes('Application')) return false;
      const enabled = it.isEnabled !== false;
      if (!enabled) return false;
      if (!filter) return true;
      const hay = `${it.value || ''}\n${it.displayName || it.adminConsentDisplayName || ''}\n${it.description || it.adminConsentDescription || ''}`.toLowerCase();
      return hay.includes(filter);
    }).sort((a, b) => (a.value || '').localeCompare(b.value || ''));

    if (!filtered.length) { tbody.innerHTML = '<tr><td colspan="4" class="muted">No matching permissions.</td></tr>'; return; }

    tbody.innerHTML = filtered.map(it => {
      const id = it.id;
      const checked = state.addSelected.has(id) ? 'checked' : '';
      const name = it.value || it.displayName || it.adminConsentDisplayName || id;
      const display = it.displayName || it.adminConsentDisplayName || '';
      const desc = it.description || it.adminConsentDescription || it.userConsentDescription || '';
      return `<tr data-id="${escAttr(id)}">
        <td><input type="checkbox" data-act="sel-perm" ${checked}/></td>
        <td><code>${escHtml(name)}</code><br><span class="muted small">${escHtml(display)}</span></td>
        <td>${isAppRole ? 'Application' : 'Delegated'}</td>
        <td class="muted small">${escHtml(desc)}</td>
      </tr>`;
    }).join('');
  }

  function onAddPermTableChange(e) {
    const cb = e.target.closest('input[data-act="sel-perm"]');
    if (!cb) return;
    const id = cb.closest('tr').dataset.id;
    if (cb.checked) state.addSelected.add(id); else state.addSelected.delete(id);
    document.getElementById('addSelCount').textContent = String(state.addSelected.size);
    document.getElementById('addPermsBtn').disabled = state.addSelected.size === 0;
  }

  function clearAddFeedback() {
    const el = document.getElementById('addFeedback');
    if (!el) return;
    el.textContent = '';
    el.className = 'inline-feedback hidden';
  }

  function setAddFeedback(message, kind) {
    const el = document.getElementById('addFeedback');
    if (!el) return;
    el.textContent = message;
    el.className = `inline-feedback feedback-${kind}`;
  }

  async function onAddPermsClick() {
    const sp = state.selectedSpn; const api = state.addApiSpn;
    if (!sp || !api || !state.addSelected.size) return;
    const isAppRole = state.addApiKind === 'AppRole';
    const items = (isAppRole ? api.appRoles : api.oauth2PermissionScopes) || [];
    const picks = items.filter(it => state.addSelected.has(it.id));

    const listHtml = picks.map(p => `<li><code>${escHtml(p.value || p.displayName || p.id)}</code></li>`).join('');
    const ok = await confirmModal(
      'Grant permissions?',
      `<p>You are about to grant <strong>${picks.length}</strong> ${isAppRole ? 'application' : 'delegated'} permission(s) on <strong>${escHtml(api.displayName)}</strong> to <strong>${escHtml(sp.displayName)}</strong>:</p>
       <ul>${listHtml}</ul>
       ${isAppRole
         ? '<p class="notice notice-warn">Application permissions take effect tenant-wide and need no user interaction. Make sure this SPN should really get this access.</p>'
         : '<p class="notice notice-info">Delegated grants will be added with consentType <code>AllPrincipals</code> (covers all users in your tenant).</p>'}`,
      'Grant'
    );
    if (!ok) return;
    const btn = document.getElementById('addPermsBtn');
    btn.disabled = true;
    clearAddFeedback();
    const resultsEl = document.getElementById('addResults');
    resultsEl.classList.remove('hidden');
    resultsEl.innerHTML = '';
    let successCount = 0;
    let failedCount = 0;

    if (isAppRole) {
      for (const p of picks) {
        const li = document.createElement('li');
        li.innerHTML = `<span class="ok-mark">...</span> <code>${escHtml(p.value || p.id)}</code>`;
        resultsEl.appendChild(li);
        try {
          await Graph.addAppRoleAssignment(sp.id, { resourceId: api.id, appRoleId: p.id });
          li.innerHTML = `<span class="ok-mark">&#10003;</span> Granted <code>${escHtml(p.value || p.id)}</code>`;
          successCount += 1;
        } catch (e) {
          li.innerHTML = `<span class="err-mark">&#10007;</span> <code>${escHtml(p.value || p.id)}</code> - ${escHtml(e.message)}`;
          Log.err('Grant failed:', e);
          failedCount += 1;
        }
      }
    } else {
      // Delegated: combine all picked scopes into a single grant entry per (client, resource, AllPrincipals).
      // If a grant already exists, PATCH its scope; otherwise POST a new one.
      const newScopes = picks.map(p => p.value).filter(Boolean);
      const existing = state.oauth2Grants.find(g => g.resourceId === api.id && g.consentType === 'AllPrincipals');
      const li = document.createElement('li');
      resultsEl.appendChild(li);
      try {
        if (existing) {
          const merged = Spns.joinScopeString([...Spns.parseScopeString(existing.scope), ...newScopes]);
          await Graph.updateOAuth2Grant(existing.id, { scope: merged });
          li.innerHTML = `<span class="ok-mark">&#10003;</span> Added scopes to existing grant: <code>${escHtml(newScopes.join(' '))}</code>`;
          successCount += newScopes.length;
        } else {
          await Graph.createOAuth2Grant({ clientId: sp.id, resourceId: api.id, scope: Spns.joinScopeString(newScopes), consentType: 'AllPrincipals' });
          li.innerHTML = `<span class="ok-mark">&#10003;</span> Created grant with scopes: <code>${escHtml(newScopes.join(' '))}</code>`;
          successCount += newScopes.length;
        }
      } catch (e) {
        li.innerHTML = `<span class="err-mark">&#10007;</span> ${escHtml(e.message)}`;
        Log.err('Grant failed:', e);
        failedCount += newScopes.length;
      }
    }

    if (successCount > 0 && failedCount === 0) {
      setAddFeedback(`Success: ${successCount} permission(s) granted.`, 'ok');
      toast(`Granted ${successCount} permission(s).`, 'ok');
    } else if (successCount > 0 && failedCount > 0) {
      setAddFeedback(`Partial success: ${successCount} granted, ${failedCount} failed. Check details below.`, 'warn');
      toast(`Partial success: ${successCount} granted, ${failedCount} failed.`, 'warn');
    } else {
      setAddFeedback('No permissions were granted. Check the errors below.', 'err');
      toast('Grant failed. See details in the result list.', 'err');
    }

    state.addSelected.clear();
    document.getElementById('addSelCount').textContent = '0';
    btn.disabled = true;
    // Bridge Graph's eventual consistency: unlock the site manager right away when we just
    // granted Sites.Selected on Graph or SharePoint Online, so no manual refresh is needed.
    if (isAppRole && successCount > 0 &&
        [Spns.GRAPH_APP_ID, Spns.SPO_APP_ID].includes((api.appId || '').toLowerCase()) &&
        picks.some(p => (p.value || '').toLowerCase() === 'sites.selected')) {
      state.siteManager.optimisticUnlock = true;
    }
    await loadPermissions();
  }

  // --------- SITE MANAGER: MODAL LIFECYCLE ---------
  const SITE_SEARCH_DEBOUNCE = 300;
  let _siteSearchTimer = null;

  async function onManageSitesClick() {
    const sp = state.selectedSpn;
    if (!sp) return;
    if (!canManageSites()) return;
    if (!sp.appId) { toast('This SPN has no appId, so it cannot be granted site access.', 'err'); return; }

    // Acquire the Sites.FullControl.All token up front. On first use this triggers
    // an interactive consent (redirect); the modal is opened once we have the token.
    const btn = document.getElementById('manageSitesBtn');
    btn.disabled = true;
    try {
      await Auth.getSitesToken();
    } catch (e) {
      Log.err('Sites token acquisition failed:', e);
      toast('Could not obtain SharePoint permission: ' + e.message, 'err');
      updateManageSitesButton();
      return;
    }
    updateManageSitesButton();
    openSiteManager();
  }

  function openSiteManager() {
    const sm = state.siteManager;
    sm.sites = [];
    sm.selectedSite = null;
    sm.sitePerms = [];
    sm.selectedRoles = new Set();
    document.getElementById('siteModalSpn').textContent = `- ${state.selectedSpn.displayName || state.selectedSpn.appId}`;
    document.getElementById('siteSearch').value = '';
    document.getElementById('siteDetail').classList.add('hidden');
    document.getElementById('siteResults').innerHTML = '';
    document.getElementById('sitePickerModal').classList.remove('hidden');
    renderRolePicker();
    document.getElementById('siteSearch').focus();
    searchSitesForQuery('*');
  }

  function closeSiteManager() {
    document.getElementById('sitePickerModal').classList.add('hidden');
    if (_siteSearchTimer) { clearTimeout(_siteSearchTimer); _siteSearchTimer = null; }
  }

  // --------- SITE MANAGER: SEARCH + LIST ---------
  function onSiteSearchInput() {
    if (_siteSearchTimer) clearTimeout(_siteSearchTimer);
    const q = document.getElementById('siteSearch').value.trim();
    _siteSearchTimer = setTimeout(() => searchSitesForQuery(q || '*'), SITE_SEARCH_DEBOUNCE);
  }

  async function searchSitesForQuery(query) {
    const sm = state.siteManager;
    const status = document.getElementById('siteSearchStatus');
    const token = ++sm.searchToken;
    status.textContent = 'Searching...';
    try {
      const sites = await Graph.searchSites(query);
      if (token !== sm.searchToken) return; // a newer search superseded this one
      sm.sites = sites;
      status.textContent = `${sites.length} site${sites.length === 1 ? '' : 's'}${sites.length >= 100 ? '+ (refine your search)' : ''}`;
      renderSiteResults();
    } catch (e) {
      if (token !== sm.searchToken) return;
      Log.err('Site search failed:', e);
      status.textContent = 'Search failed: ' + e.message;
    }
  }

  function renderSiteResults() {
    const ul = document.getElementById('siteResults');
    const sm = state.siteManager;
    if (!sm.sites.length) { ul.innerHTML = '<li class="muted site-empty">No sites match.</li>'; return; }
    ul.innerHTML = sm.sites.map(s => `
      <li class="site-list-item${sm.selectedSite && sm.selectedSite.id === s.id ? ' active' : ''}" data-id="${escAttr(s.id)}">
        <span class="site-name">${escHtml(s.displayName || s.name || '(no name)')}</span>
        <span class="muted small site-url">${escHtml(s.webUrl || '')}</span>
      </li>`).join('');
  }

  function onSiteResultClick(e) {
    const li = e.target.closest('li[data-id]');
    if (!li) return;
    const site = state.siteManager.sites.find(s => s.id === li.dataset.id);
    if (site) selectSite(site);
  }

  // --------- SITE MANAGER: DETAIL + PERMISSIONS ---------
  function setSiteDetailHead(site) {
    document.getElementById('siteDetailName').textContent = site.displayName || site.name || '(no name)';
    const urlEl = document.getElementById('siteDetailUrl');
    urlEl.textContent = site.webUrl || '';
    urlEl.href = site.webUrl || '#';
  }

  // Sites.Selected permissions live on the site collection, not subsites. Derive the
  // collection root from the webUrl (managed path + first segment, e.g. /sites/Marketing).
  async function resolveSiteCollectionSite(site) {
    if (!site.webUrl) return null;
    const u = new URL(site.webUrl);
    const segs = u.pathname.split('/').filter(Boolean);
    const rel = segs.length >= 2
      ? `/sites/${u.hostname}:/${encodeURI(segs[0] + '/' + segs[1])}`
      : `/sites/${u.hostname}`;
    return Graph.call(rel, { sites: true });
  }

  async function selectSite(site) {
    const sm = state.siteManager;
    sm.selectedSite = site;
    renderSiteResults();
    document.getElementById('siteDetail').classList.remove('hidden');
    document.getElementById('siteScopeNote').classList.add('hidden');
    setSiteDetailHead(site);
    document.querySelector('#sitePermsTable tbody').innerHTML = '<tr><td colspan="4" class="muted">Loading...</td></tr>';
    clearSiteGrantFeedback();
    try {
      sm.sitePerms = await Graph.getSitePermissions(site.id);
      renderSitePerms();
    } catch (e) {
      // A subsite returns 400 notSupported - resolve to its site collection and use that.
      if (e.status === 400) {
        const root = await resolveSiteCollectionSite(site).catch(err => { Log.warn('Site-collection resolve failed: ' + err.message); return null; });
        if (root && root.id && root.id !== site.id) {
          sm.selectedSite = root;
          setSiteDetailHead(root);
          const note = document.getElementById('siteScopeNote');
          note.textContent = `You picked a subsite. Sites.Selected is granted at the site-collection level, so this shows the parent site collection "${root.displayName || root.name || root.webUrl}" - access granted here applies to all of its subsites.`;
          note.classList.remove('hidden');
          try {
            sm.sitePerms = await Graph.getSitePermissions(root.id);
            renderSitePerms();
            return;
          } catch (e2) { e = e2; }
        }
      }
      Log.err('Load site permissions failed:', e);
      document.querySelector('#sitePermsTable tbody').innerHTML = `<tr><td colspan="4" class="muted">Failed: ${escHtml(e.message)}</td></tr>`;
    }
  }

  function renderSitePerms() {
    const sm = state.siteManager;
    const rows = Spns.extractSiteAppPermissions(sm.sitePerms);
    const tbody = document.querySelector('#sitePermsTable tbody');
    const empty = document.getElementById('sitePermsEmpty');
    document.getElementById('sitePermsCount').textContent = `(${rows.length})`;
    const selfAppId = (state.selectedSpn.appId || '').toLowerCase();
    if (!rows.length) {
      tbody.innerHTML = '';
      empty.classList.remove('hidden');
    } else {
      empty.classList.add('hidden');
      tbody.innerHTML = rows.map(r => {
        const isSelf = r.appId.toLowerCase() === selfAppId;
        return `<tr class="${isSelf ? 'self-perm' : ''}">
          <td>${escHtml(r.displayName || '(unknown app)')}${isSelf ? ' <span class="pill">this SPN</span>' : ''}</td>
          <td><code>${escHtml(r.appId)}</code></td>
          <td><code>${escHtml(r.roles.join(', ') || '(none)')}</code></td>
          <td></td>
        </tr>`;
      }).join('');
    }
    renderGrantBox();
  }

  function renderRolePicker() {
    const box = document.getElementById('siteRolePicker');
    const sm = state.siteManager;
    box.innerHTML = Spns.SITE_ROLE_OPTIONS.map(opt => `
      <label class="checkbox role-option${opt.extended ? ' extended' : ''}" title="${escAttr(opt.hint)}${opt.extended ? ' (may not be supported in every tenant)' : ''}">
        <input type="checkbox" data-role="${escAttr(opt.value)}" ${sm.selectedRoles.has(opt.value) ? 'checked' : ''}/>
        ${escHtml(opt.label)}
      </label>`).join('');
  }

  function renderGrantBox() {
    const sm = state.siteManager;
    const existing = Spns.findAppPermissionOnSite(sm.sitePerms, state.selectedSpn.appId);
    const title = document.getElementById('siteGrantTitle');
    const removeBtn = document.getElementById('siteRemoveBtn');
    const hint = document.getElementById('siteGrantHint');
    sm.selectedRoles = new Set(existing ? existing.roles : []);
    const spName = state.selectedSpn.displayName || state.selectedSpn.appId;
    if (existing) {
      title.textContent = `Update ${spName}\u2019s access`;
      removeBtn.classList.remove('hidden');
      hint.textContent = `Currently granted: ${existing.roles.join(', ') || '(none)'}`;
    } else {
      title.textContent = `Grant access to ${spName}`;
      removeBtn.classList.add('hidden');
      hint.textContent = `${spName} has no access to this site yet.`;
    }
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
    const sm = state.siteManager;
    const existing = Spns.findAppPermissionOnSite(sm.sitePerms, state.selectedSpn.appId);
    const btn = document.getElementById('siteGrantBtn');
    const roles = [...sm.selectedRoles];
    btn.textContent = existing ? 'Update roles' : 'Grant';
    // Disable when nothing is selected, or when the selection equals the existing roles.
    let changed = roles.length > 0;
    if (existing) {
      const cur = new Set(existing.roles);
      changed = roles.length > 0 && (roles.length !== cur.size || roles.some(r => !cur.has(r)));
    }
    btn.disabled = !changed;
  }

  function clearSiteGrantFeedback() {
    const el = document.getElementById('siteGrantFeedback');
    el.textContent = '';
    el.className = 'inline-feedback hidden';
  }
  function setSiteGrantFeedback(msg, kind) {
    const el = document.getElementById('siteGrantFeedback');
    el.textContent = msg;
    el.className = `inline-feedback feedback-${kind}`;
  }

  async function onGrantSiteClick() {
    const sm = state.siteManager;
    const sp = state.selectedSpn;
    const site = sm.selectedSite;
    if (!site) return;
    const roles = [...sm.selectedRoles];
    if (!roles.length) return;
    const existing = Spns.findAppPermissionOnSite(sm.sitePerms, sp.appId);

    const ok = await confirmModal(
      existing ? 'Update site access?' : 'Grant site access?',
      `<p>${existing ? 'Change' : 'Grant'} <strong>${escHtml(sp.displayName || sp.appId)}</strong>'s access on <strong>${escHtml(site.displayName || site.name || site.webUrl)}</strong> to role(s): <code>${escHtml(roles.join(', '))}</code>.</p>
       <p class="notice notice-info">This changes application access on a single SharePoint site. The app can act on this site without a signed-in user once granted.</p>`,
      existing ? 'Update' : 'Grant'
    );
    if (!ok) return;

    const btn = document.getElementById('siteGrantBtn');
    btn.disabled = true;
    clearSiteGrantFeedback();
    try {
      if (existing) {
        await Graph.updateSitePermission(site.id, existing.permissionId, { roles });
      } else {
        await Graph.addSitePermission(site.id, { appId: sp.appId, displayName: sp.displayName || sp.appId, roles });
      }
      Log.info(`${existing ? 'Updated' : 'Granted'} site access for ${sp.appId} on ${site.id}: ${roles.join(', ')}`);
      setSiteGrantFeedback(`${existing ? 'Updated' : 'Granted'}: ${roles.join(', ')}`, 'ok');
      toast(existing ? 'Site access updated.' : 'Site access granted.', 'ok');
      sm.sitePerms = await Graph.getSitePermissions(site.id);
      renderSitePerms();
    } catch (e) {
      Log.err('Grant site access failed:', e);
      setSiteGrantFeedback('Failed: ' + e.message, 'err');
      toast('Grant failed: ' + e.message, 'err');
      updateGrantButtonState();
    }
  }

  async function onRemoveSiteClick() {
    const sm = state.siteManager;
    const sp = state.selectedSpn;
    const site = sm.selectedSite;
    if (!site) return;
    const existing = Spns.findAppPermissionOnSite(sm.sitePerms, sp.appId);
    if (!existing) return;

    const ok = await confirmModal(
      'Remove site access?',
      `<p>Remove <strong>${escHtml(sp.displayName || sp.appId)}</strong>'s access (<code>${escHtml(existing.roles.join(', '))}</code>) from <strong>${escHtml(site.displayName || site.name || site.webUrl)}</strong>.</p>
       <p class="notice notice-warn">The app will lose access to this site. If it relies on it, the workload may break.</p>`,
      'Remove'
    );
    if (!ok) return;

    const btn = document.getElementById('siteRemoveBtn');
    btn.disabled = true;
    try {
      await Graph.deleteSitePermission(site.id, existing.permissionId);
      Log.info(`Removed site access for ${sp.appId} on ${site.id}`);
      toast('Site access removed.', 'ok');
      sm.sitePerms = await Graph.getSitePermissions(site.id);
      renderSitePerms();
    } catch (e) {
      Log.err('Remove site access failed:', e);
      toast('Remove failed: ' + e.message, 'err');
    } finally {
      btn.disabled = false;
    }
  }

  // --------- HELPERS ---------
  function escHtml(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function escAttr(s) { return escHtml(s); }

  // --------- BOOT ---------
  async function boot() {
    initTheme();
    document.getElementById('darkToggle').addEventListener('click', toggleTheme);
    document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => activateTab(t.dataset.tab)));
    document.getElementById('signInBtn').addEventListener('click', onSignIn);
    document.getElementById('signOutBtn').addEventListener('click', onSignOut);
    document.getElementById('loadSpnsBtn').addEventListener('click', loadSpns);
    document.getElementById('spnFilter').addEventListener('input', applySpnFilter);
    document.getElementById('typeFilter').addEventListener('change', applySpnFilter);
    document.getElementById('hideMs').addEventListener('change', applySpnFilter);
    document.querySelector('#spnTable thead').addEventListener('click', onSpnSortClick);
    document.querySelector('#spnTable tbody').addEventListener('click', onSpnTableClick);
    document.querySelector('#appRolesTable tbody').addEventListener('click', onRevokeAppClick);
    document.querySelector('#oauth2Table tbody').addEventListener('click', onRevokeScopeClick);
    document.getElementById('revokeAllBtn').addEventListener('click', onRevokeAllClick);
    document.getElementById('refreshPermsBtn').addEventListener('click', loadPermissions);
    document.getElementById('manageSitesBtn').addEventListener('click', onManageSitesClick);
    document.getElementById('siteModalCloseBtn').addEventListener('click', closeSiteManager);
    document.querySelector('#sitePickerModal .modal-backdrop').addEventListener('click', closeSiteManager);
    document.getElementById('siteSearch').addEventListener('input', onSiteSearchInput);
    document.getElementById('siteResults').addEventListener('click', onSiteResultClick);
    document.getElementById('siteRolePicker').addEventListener('change', onSiteRoleToggle);
    document.getElementById('siteGrantBtn').addEventListener('click', onGrantSiteClick);
    document.getElementById('siteRemoveBtn').addEventListener('click', onRemoveSiteClick);
    const apiSearch = document.getElementById('apiSearch');
    apiSearch.addEventListener('change', onLoadApi);
    apiSearch.addEventListener('keydown', async (e) => {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      await onLoadApi();
    });
    document.querySelector('#addPermTable tbody').addEventListener('change', onAddPermTableChange);
    document.querySelectorAll('input[name="permKind"]').forEach(r => r.addEventListener('change', e => { state.addApiKind = e.target.value; state.addSelected.clear(); document.getElementById('addSelCount').textContent='0'; document.getElementById('addPermsBtn').disabled=true; renderAddPermTable(); }));
    document.getElementById('permFilter').addEventListener('input', renderAddPermTable);
    document.getElementById('addPermsBtn').addEventListener('click', onAddPermsClick);
    document.getElementById('clearLogBtn').addEventListener('click', () => Log.clear());
    document.getElementById('verboseLog').addEventListener('change', e => Log.setVerbose(e.target.checked));

    try {
      await Auth.init();
      Log.info('SPNRoleMgr ready.');
      await refreshAuthUI();
    } catch (e) {
      Log.err('Boot failed:', e);
      toast('Initialization failed: ' + e.message, 'err');
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
