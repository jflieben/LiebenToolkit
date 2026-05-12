// SPNRoleMgr - SPA wiring.
// Pattern follows DupedDevices: small global state, render functions per panel,
// confirm() before any mutation, dry-run by inspection (you only ever execute
// after picking individual rows + clicking the explicit grant/revoke button).
(() => {
  const state = {
    spns: [],            // all loaded service principals
    filteredSpns: [],
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
    state.spns = []; state.filteredSpns = []; state.selectedSpn = null;
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
      <tr data-id="${escAttr(sp.id)}">
        <td><strong>${escHtml(sp.displayName || '(no name)')}</strong></td>
        <td>${escHtml(Spns.classifySpn(sp))}</td>
        <td><code>${escHtml(sp.appId || '')}</code></td>
        <td>${sp.accountEnabled ? 'Yes' : 'No'}</td>
        <td><button class="ghost-btn" data-act="manage">Manage</button></td>
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
    const btn = e.target.closest('button[data-act="manage"]');
    if (!btn) return;
    const tr = btn.closest('tr');
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

  // --------- ADD PERMISSION PANEL ---------
  function renderAddPanel() {
    // Populate API suggestions from the loaded SPN list (only those that publish appRoles or oauth2Scopes — we can't tell that without fetching, so use full SPN list).
    const dl = document.getElementById('apiSuggestions');
    const seen = new Set();
    const opts = [];
    for (const sp of state.spns) {
      if (!sp.displayName) continue;
      if (seen.has(sp.displayName)) continue;
      seen.add(sp.displayName);
      opts.push(`<option value="${escAttr(sp.displayName)}"></option>`);
    }
    dl.innerHTML = opts.slice(0, 500).join('');
    state.addApiSpn = null;
    state.addSelected.clear();
    document.querySelector('#addPermTable tbody').innerHTML = '<tr><td colspan="4" class="muted">Pick an API and click "Load API permissions".</td></tr>';
    document.getElementById('addSelCount').textContent = '0';
    document.getElementById('addPermsBtn').disabled = true;
    document.getElementById('addResults').classList.add('hidden');
  }

  async function onLoadApi() {
    const q = document.getElementById('apiSearch').value.trim();
    const status = document.getElementById('apiLoadStatus');
    if (!q) { status.textContent = 'Enter an API name first.'; return; }
    status.textContent = 'Resolving...';
    try {
      // Try exact display-name match against loaded SPNs first.
      let resourceSpn = state.spns.find(s => (s.displayName || '').toLowerCase() === q.toLowerCase());
      if (!resourceSpn) {
        // Search by appId too.
        resourceSpn = state.spns.find(s => s.appId === q);
      }
      if (!resourceSpn) {
        // Last resort: case-insensitive contains
        resourceSpn = state.spns.find(s => (s.displayName || '').toLowerCase().includes(q.toLowerCase()));
      }
      if (!resourceSpn) { status.textContent = 'No matching SPN in the loaded list.'; return; }
      // Need full appRoles + scopes — re-fetch the SPN.
      const full = await Graph.getServicePrincipal(resourceSpn.id);
      state.addApiSpn = full;
      state.resourceCache.set(full.id, full);
      state.resourceIndexCache.set(full.id, Spns.buildPermissionIndex(full));
      const ar = (full.appRoles || []).length;
      const os = (full.oauth2PermissionScopes || []).length;
      status.textContent = `${full.displayName}: ${ar} app role(s), ${os} delegated scope(s).`;
      state.addSelected.clear();
      renderAddPermTable();
    } catch (e) {
      Log.err('Load API failed:', e);
      status.textContent = 'Failed: ' + e.message;
    }
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
    const resultsEl = document.getElementById('addResults');
    resultsEl.classList.remove('hidden');
    resultsEl.innerHTML = '';

    if (isAppRole) {
      for (const p of picks) {
        const li = document.createElement('li');
        li.innerHTML = `<span class="ok-mark">...</span> <code>${escHtml(p.value || p.id)}</code>`;
        resultsEl.appendChild(li);
        try {
          await Graph.addAppRoleAssignment(sp.id, { resourceId: api.id, appRoleId: p.id });
          li.innerHTML = `<span class="ok-mark">&#10003;</span> Granted <code>${escHtml(p.value || p.id)}</code>`;
        } catch (e) {
          li.innerHTML = `<span class="err-mark">&#10007;</span> <code>${escHtml(p.value || p.id)}</code> - ${escHtml(e.message)}`;
          Log.err('Grant failed:', e);
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
        } else {
          await Graph.createOAuth2Grant({ clientId: sp.id, resourceId: api.id, scope: Spns.joinScopeString(newScopes), consentType: 'AllPrincipals' });
          li.innerHTML = `<span class="ok-mark">&#10003;</span> Created grant with scopes: <code>${escHtml(newScopes.join(' '))}</code>`;
        }
      } catch (e) {
        li.innerHTML = `<span class="err-mark">&#10007;</span> ${escHtml(e.message)}`;
        Log.err('Grant failed:', e);
      }
    }

    state.addSelected.clear();
    document.getElementById('addSelCount').textContent = '0';
    btn.disabled = true;
    await loadPermissions();
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
    document.getElementById('refreshPermsBtn').addEventListener('click', loadPermissions);
    document.getElementById('loadApiBtn').addEventListener('click', onLoadApi);
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
