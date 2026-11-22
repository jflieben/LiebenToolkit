// ConditionalAccess UI.
(() => {
  const { esc, h, toast, confirm, modal, busy, setUser, iconHtml } = Toolkit;
  const P = () => window.PolicyUtils;

  window.Auth = ToolkitAuth.create({
    scopes: {
      read: ['User.Read', 'Policy.Read.All', 'Directory.Read.All'],
      write: ['Policy.ReadWrite.ConditionalAccess'],
    },
  });

  const STATE_LABEL = { disabled: 'Off', enabledForReportingButNotEnforced: 'Report-only', enabled: 'On' };
  const STATE_BADGE = { disabled: 'neutral', enabledForReportingButNotEnforced: 'info', enabled: 'ok' };
  const KIND_LABEL = { user: 'User', group: 'Group', role: 'Role', app: 'App', servicePrincipal: 'Workload identity', location: 'Named location' };
  const PREFS_KEY = 'ca-tool-prefs';

  const state = {
    tenantPolicies: [],
    rows: [],
    source: null,          // 'import' | 'starter'
    sourceLabel: '',
    names: new Map(),      // lower id -> { name, kind }
    missing: new Set(),    // lower ids of tenant-specific kinds that were looked up and not found
    resolverMap: {},       // lower missing id -> replacement id
    replacements: new Map(), // lower replacement id -> { name, kind } | null
    me: { userId: null, groupIds: [], roleIds: [] },
    roles: null,           // Map templateId -> name
    locations: null,       // Map id -> name
    lookupBusy: false,
  };
  const $ = (id) => document.getElementById(id);

  // ---------- prefs (prefix and suffix only; overwrite is never remembered) ----------
  function loadPrefs() {
    try {
      const prefs = JSON.parse(localStorage.getItem(PREFS_KEY) || '{}');
      $('namePrefix').value = typeof prefs.prefix === 'string' ? prefs.prefix : '';
      $('nameSuffix').value = typeof prefs.suffix === 'string' ? prefs.suffix : '';
    } catch { /* storage blocked */ }
  }
  function savePrefs() {
    try { localStorage.setItem(PREFS_KEY, JSON.stringify({ prefix: $('namePrefix').value, suffix: $('nameSuffix').value })); }
    catch { /* storage blocked */ }
  }

  // ---------- helpers ----------
  function fmtDate(v) {
    if (!v) return '';
    const d = new Date(v);
    return Number.isNaN(d.getTime()) ? '' : d.toISOString().replace('T', ' ').substring(0, 16);
  }
  function readTextFile(file) {
    if (file.size > 5 * 1024 * 1024) return Promise.reject(new Error('File is larger than 5 MB. That is not a policy export.'));
    return file.text();
  }
  const plural = (n, one, many) => `${n} ${n === 1 ? one : (many || `${one}s`)}`;
  const shortJson = (v) => {
    if (v === undefined || v === null) return '(not set)';
    const s = typeof v === 'string' ? v : JSON.stringify(v);
    return s.length > 400 ? `${s.substring(0, 400)}...` : s;
  };

  // ---------- auth ----------
  function refreshAuthUI() {
    const signedIn = !!Auth.getAccount();
    setUser(Auth.userInfo(onSignOut));
    $('connectCard').hidden = signedIn;
    $('policiesCard').hidden = !signedIn;
    $('importSignedOut').hidden = signedIn;
    $('importSources').hidden = !signedIn;
    if (!signedIn) {
      $('reviewCard').hidden = true;
      $('resolverCard').hidden = true;
      $('applyCard').hidden = true;
    }
  }
  async function onSignIn() {
    try { await Auth.signIn('read'); }
    catch (e) { Log.err('Sign-in failed:', e); toast(`Sign-in failed: ${e.message}`, 'err'); }
  }
  async function onSignOut() {
    state.rows = [];
    state.tenantPolicies = [];
    await Auth.signOut();
  }

  // Own id, groups and roles: the lockout check needs them. Best effort.
  async function loadMe() {
    const acc = Auth.getAccount();
    const claims = (acc && acc.idTokenClaims) || {};
    const oid = (acc && acc.localAccountId) || claims.oid || null;
    state.me.userId = Http.isGuid(oid) ? oid.toLowerCase() : null;
    if (state.me.userId) state.names.set(state.me.userId, { name: `${acc.name || acc.username} (you)`, upn: acc.username || '', kind: 'user' });
    try {
      const objs = await Graph.myMemberships();
      state.me.groupIds = objs.filter((o) => /group$/i.test(o['@odata.type'] || '')).map((o) => String(o.id).toLowerCase());
      state.me.roleIds = objs.filter((o) => /directoryRole$/i.test(o['@odata.type'] || '') && o.roleTemplateId).map((o) => String(o.roleTemplateId).toLowerCase());
      Log.info(`Lockout check: you are in ${plural(state.me.groupIds.length, 'group')} and ${plural(state.me.roleIds.length, 'active role')}.`);
    } catch (e) {
      Log.warn('Could not read your group and role memberships; the lockout check only looks at your user id.', e.message || e);
    }
    renderReview();
  }

  // ---------- tenant policies ----------
  function usersSummary(p) {
    const u = (p.conditions || {}).users || {};
    const parts = [];
    const inc = [...(u.includeUsers || []), ...(u.includeGroups || []).map(() => 'group'), ...(u.includeRoles || []).map(() => 'role')];
    if ((u.includeUsers || []).includes('All')) parts.push('All users');
    else if (inc.length) parts.push(`${plural((u.includeUsers || []).length, 'user')}, ${plural((u.includeGroups || []).length, 'group')}, ${plural((u.includeRoles || []).length, 'role')}`);
    const exc = (u.excludeUsers || []).length + (u.excludeGroups || []).length + (u.excludeRoles || []).length;
    if (exc) parts.push(`${exc} excluded`);
    return parts.join(', ') || 'None';
  }
  function grantSummary(p) {
    const g = p.grantControls || {};
    const bits = [...(g.builtInControls || [])];
    if (g.authenticationStrength) bits.push('auth strength');
    if ((g.termsOfUse || []).length) bits.push('terms of use');
    const session = p.sessionControls ? Object.keys(p.sessionControls).filter((k) => p.sessionControls[k] && !k.startsWith('@')) : [];
    return [bits.join(` ${(g.operator || 'OR').toLowerCase()} `), session.length ? `session: ${session.join(', ')}` : ''].filter(Boolean).join('; ') || 'None';
  }

  async function refreshTenantPolicies() {
    $('policiesStatus').textContent = 'Loading policies...';
    const list = await Graph.listConditionalAccessPolicies();
    state.tenantPolicies = list.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
    renderTenantPolicies();
    Log.info(`Loaded ${plural(state.tenantPolicies.length, 'policy', 'policies')}.`);
    if (state.rows.length) { applyNaming(); renderReview(); }
  }

  function renderTenantPolicies() {
    const list = state.tenantPolicies;
    const counts = { enabled: 0, enabledForReportingButNotEnforced: 0, disabled: 0 };
    list.forEach((p) => { if (counts[p.state] != null) counts[p.state]++; });
    $('policiesStatus').textContent = `${plural(list.length, 'policy', 'policies')}: ${counts.enabled} on, ${counts.enabledForReportingButNotEnforced} report-only, ${counts.disabled} off.`;
    $('policiesTableWrap').hidden = !list.length;
    $('policiesEmpty').hidden = !!list.length;
    $('policiesActions').hidden = !list.length;
    document.querySelector('#policiesTable tbody').innerHTML = list.map((p) => `
      <tr>
        <td><strong>${esc(p.displayName || '(no name)')}</strong></td>
        <td><span class="badge dot ${STATE_BADGE[p.state] || 'neutral'}">${esc(STATE_LABEL[p.state] || p.state || '')}</span></td>
        <td>${esc(usersSummary(p))}</td>
        <td>${esc(grantSummary(p))}</td>
        <td class="nowrap">${esc(fmtDate(p.createdDateTime))}</td>
        <td class="nowrap">${esc(fmtDate(p.modifiedDateTime))}</td>
      </tr>`).join('');
  }

  function exportJson() {
    const acc = Auth.getAccount();
    const payload = {
      tool: 'ConditionalAccess',
      schemaVersion: 1,
      exportedAtUtc: new Date().toISOString(),
      policyCount: state.tenantPolicies.length,
      metadata: {
        user: (acc && (acc.username || acc.name)) || '',
        tenantId: (acc && acc.tenantId) || '',
        source: 'Graph /identity/conditionalAccess/policies',
      },
      policies: state.tenantPolicies,
    };
    const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
    Export.save(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' }), `conditional-access-export-${stamp}.json`);
    toast(`Exported ${plural(state.tenantPolicies.length, 'policy', 'policies')}`, 'ok');
  }

  const EXPORT_COLUMNS = [
    { key: 'displayName', label: 'Name' },
    { key: 'state', label: 'State', format: (v) => STATE_LABEL[v] || v },
    { key: 'users', label: 'Users', format: (v, p) => usersSummary(p) },
    { key: 'grant', label: 'Controls', format: (v, p) => grantSummary(p) },
    { key: 'createdDateTime', label: 'Created', format: fmtDate },
    { key: 'modifiedDateTime', label: 'Modified', format: fmtDate },
    { key: 'id', label: 'Policy id' },
  ];

  // ---------- import rows ----------
  function toRows(validated) {
    return validated.map((v, idx) => ({
      rowId: `r${idx}`,
      policy: v.policy,
      fileState: v.fileState,
      dropped: v.dropped || [],
      selected: true,
      sourceName: v.policy.displayName,
      finalName: v.policy.displayName,
      targetState: P().DEFAULT_STATE,
      excludeMe: false,
      matches: [],
      open: false,
    }));
  }

  function applyNaming() {
    const prefix = $('namePrefix').value || '';
    const suffix = $('nameSuffix').value || '';
    for (const row of state.rows) {
      row.finalName = `${prefix}${row.sourceName}${suffix}`.trim();
      row.matches = state.tenantPolicies.filter((p) => (p.displayName || '').trim() === row.finalName);
    }
  }

  function mapping() {
    const out = {};
    for (const [k, v] of Object.entries(state.resolverMap)) {
      const hit = state.replacements.get(String(v).toLowerCase());
      if (hit) out[k] = String(v).trim().toLowerCase();
    }
    return out;
  }

  // Exactly what Graph will receive for this row.
  function bodyFor(row) {
    let p = P().applyPrincipalMapping(row.policy, mapping());
    if (row.excludeMe && state.me.userId) p = P().addUserExclusion(p, state.me.userId);
    return P().normalizePolicyForCreateOrUpdate(p, row.finalName, row.targetState);
  }

  function rowAction(row) {
    if (!row.matches.length) return 'create';
    if (!$('overwriteExisting').checked) return 'skip';
    if (row.matches.length > 1) return 'ambiguous';
    return 'overwrite';
  }

  function riskFor(row) {
    return P().lockoutRisk(bodyFor(row), state.me);
  }

  // ---------- name lookup ----------
  async function lookupRefs() {
    const refs = state.rows.flatMap((r) => P().extractRefs(r.policy)).filter((r) => r.isId);
    const byKind = (kinds) => Array.from(new Set(refs.filter((r) => kinds.includes(r.kind)).map((r) => r.value.toLowerCase())));
    const dirIds = byKind(['user', 'group', 'servicePrincipal']).filter((id) => !state.names.has(id));
    const appIds = byKind(['app']).filter((id) => !state.names.has(id));
    state.lookupBusy = true;
    renderReview();
    try {
      if (dirIds.length) {
        const found = await Graph.getDirectoryObjectsByIds(dirIds);
        for (const o of found) {
          const kind = /group$/i.test(o['@odata.type'] || '') ? 'group' : /servicePrincipal$/i.test(o['@odata.type'] || '') ? 'servicePrincipal' : 'user';
          state.names.set(String(o.id).toLowerCase(), { name: o.displayName || o.userPrincipalName || o.id, upn: o.userPrincipalName || '', kind });
        }
        dirIds.filter((id) => !state.names.has(id)).forEach((id) => state.missing.add(id));
      }
      if (byKind(['role']).length && !state.roles) {
        state.roles = new Map();
        for (const r of await Graph.listRoleDefinitions()) state.roles.set(String(r.templateId || r.id).toLowerCase(), r.displayName);
      }
      if (byKind(['location']).length && !state.locations) {
        state.locations = new Map();
        for (const l of await Graph.listNamedLocations()) state.locations.set(String(l.id).toLowerCase(), l.displayName);
      }
      byKind(['location']).filter((id) => !state.locations.has(id)).forEach((id) => state.missing.add(id));
      if (appIds.length) {
        const sps = await Graph.servicePrincipalsByAppIds(appIds);
        for (const sp of sps) state.names.set(String(sp.appId).toLowerCase(), { name: sp.displayName, kind: 'app' });
      }
    } catch (e) {
      Log.err('Looking up names failed:', e);
      toast(`Looking up names failed: ${e.message}`, 'err');
    } finally {
      state.lookupBusy = false;
    }
    renderReview();
  }

  function refName(ref) {
    const id = ref.value.toLowerCase();
    if (ref.kind === 'role') return state.roles ? state.roles.get(id) || null : undefined;
    if (ref.kind === 'location') return state.locations ? state.locations.get(id) || null : undefined;
    const hit = state.names.get(id);
    if (hit) return hit.upn && hit.upn !== hit.name ? `${hit.name} (${hit.upn})` : hit.name;
    if (state.missing.has(id)) return null;
    if (ref.kind === 'app') return state.lookupBusy ? undefined : null;
    return undefined;
  }

  // ---------- review rendering ----------
  function refHtml(ref, mapped) {
    if (!ref.isId) return `<span class="badge neutral">${esc(ref.value)}</span>`;
    const name = refName(ref);
    const repl = mapped[ref.value.toLowerCase()];
    if (repl) {
      const hit = state.replacements.get(repl);
      return `<span class="ref"><span class="badge warn">mapped</span> <code>${esc(ref.value)}</code> to <strong>${esc(hit ? hit.name : repl)}</strong> <code>${esc(repl)}</code></span>`;
    }
    if (name === undefined) return `<span class="ref"><span class="muted">looking up</span> <code>${esc(ref.value)}</code></span>`;
    if (name === null) {
      const tenantSpecific = P().MAPPABLE_KINDS.has(ref.kind);
      const label = tenantSpecific ? 'not in tenant' : ref.kind === 'app' ? 'no app in tenant' : 'unknown';
      return `<span class="ref"><span class="badge ${tenantSpecific ? 'danger' : 'warn'}">${label}</span> <code>${esc(ref.value)}</code></span>`;
    }
    return `<span class="ref"><strong>${esc(name)}</strong> <code>${esc(ref.value)}</code></span>`;
  }

  // From the source policy, so mapped ids show both sides of the mapping.
  function assignmentsHtml(row) {
    const policy = row.excludeMe && state.me.userId ? P().addUserExclusion(row.policy, state.me.userId) : row.policy;
    const refs = P().extractRefs(policy);
    const mapped = mapping();
    const groups = new Map();
    for (const r of refs) {
      if (!groups.has(r.label)) groups.set(r.label, []);
      groups.get(r.label).push(r);
    }
    const actions = (((policy.conditions || {}).applications || {}).includeUserActions || []);
    let html = [...groups.entries()].filter(([, list]) => list.length).map(([label, list]) =>
      `<div class="ref-group"><span class="ref-label">${esc(label)}</span>${list.map((r) => refHtml(r, mapped)).join('')}</div>`).join('');
    if (actions.length) html += `<div class="ref-group"><span class="ref-label">User actions</span>${actions.map((a) => `<span class="badge neutral">${esc(a)}</span>`).join('')}</div>`;
    return html || '<span class="muted">None</span>';
  }

  function controlsHtml(body) {
    const g = body.grantControls || {};
    const out = [];
    if ((g.builtInControls || []).length) out.push(`<div><span class="ref-label">Grant (${esc(g.operator || 'OR')})</span>${g.builtInControls.map((c) => `<span class="badge ${c === 'block' ? 'danger' : 'info'}">${esc(c)}</span>`).join('')}</div>`);
    if (g.authenticationStrength) out.push(`<div><span class="ref-label">Auth strength</span><code>${esc(g.authenticationStrength.id)}</code></div>`);
    if ((g.termsOfUse || []).length) out.push(`<div><span class="ref-label">Terms of use</span>${g.termsOfUse.map((t) => `<code>${esc(t)}</code>`).join(' ')}</div>`);
    const s = body.sessionControls;
    if (s) {
      const keys = Object.keys(s).filter((k) => s[k] != null && s[k] !== false);
      if (keys.length) out.push(`<div><span class="ref-label">Session</span>${keys.map((k) => `<span class="badge neutral">${esc(k)}</span>`).join('')}</div>`);
    }
    const c = body.conditions || {};
    const extra = ['clientAppTypes', 'platforms', 'locations', 'devices', 'signInRiskLevels', 'userRiskLevels', 'authenticationFlows', 'insiderRiskLevels', 'servicePrincipalRiskLevels']
      .filter((k) => c[k] != null && !(Array.isArray(c[k]) && !c[k].length));
    if (extra.length) out.push(`<div><span class="ref-label">Conditions</span>${extra.map((k) => `<span class="badge neutral">${esc(k)}</span>`).join('')}</div>`);
    return out.join('') || '<span class="muted">None</span>';
  }

  function policyCellHtml(row, risk) {
    const action = rowAction(row);
    const badges = [];
    if (action === 'create') badges.push('<span class="badge ok">New</span>');
    if (action === 'skip') badges.push('<span class="badge neutral">Exists, skipped</span>');
    if (action === 'ambiguous') badges.push(`<span class="badge danger">${row.matches.length} policies with this name, skipped</span>`);
    if (action === 'overwrite') badges.push('<span class="badge warn">Overwrites existing</span>');
    if (row.dropped.length) badges.push(`<span class="badge info">${plural(row.dropped.length, 'field')} dropped</span>`);
    if (row.excludeMe) badges.push('<span class="badge ok">You are excluded</span>');
    let html = `<strong>${esc(row.finalName)}</strong>`;
    if (row.finalName !== row.sourceName) html += `<br><span class="muted small">from ${esc(row.sourceName)}</span>`;
    html += `<div class="row tight badges">${badges.join('')}</div>`;
    if (risk) {
      html += `<div class="risk">${iconHtml('i-alert', 'sm')}<span><strong>Lockout risk.</strong> ${esc(risk.reason)}</span></div>`;
    }
    const buttons = [];
    if (risk && state.me.userId && !row.excludeMe) buttons.push(`<button type="button" class="btn btn-small" data-act="exclude-me">${iconHtml('i-user')}Exclude me</button>`);
    if (row.excludeMe) buttons.push(`<button type="button" class="btn btn-small btn-ghost" data-act="include-me">${iconHtml('i-undo')}Undo exclusion</button>`);
    buttons.push(`<button type="button" class="btn btn-small btn-ghost" data-act="details" aria-expanded="${row.open ? 'true' : 'false'}">${iconHtml('i-code')}${row.open ? 'Hide JSON' : 'Show JSON'}</button>`);
    html += `<div class="row tight">${buttons.join('')}</div>`;
    return html;
  }

  function stateCellHtml(row) {
    const opts = P().STATES.map((s) => `<option value="${s}" ${row.targetState === s ? 'selected' : ''}>${STATE_LABEL[s]}</option>`).join('');
    const note = row.fileState && row.fileState !== row.targetState ? `<br><span class="muted small">File says ${esc(STATE_LABEL[row.fileState])}</span>` : '';
    return `<select data-act="state" aria-label="State for ${esc(row.finalName)}">${opts}</select>${note}`;
  }

  function detailHtml(row, body) {
    const dropped = row.dropped.length
      ? `<p class="small"><strong>Dropped from the file:</strong> ${row.dropped.map((d) => `<code>${esc(d.path)}</code> <span class="muted">(${esc(d.reason)})</span>`).join(', ')}</p>`
      : '';
    return `${dropped}<p class="small muted">Body sent to Graph:</p><pre class="json-view">${esc(JSON.stringify(body, null, 2))}</pre>`;
  }

  function renderReview() {
    const has = state.rows.length > 0;
    $('reviewCard').hidden = !has;
    if (!has) { $('resolverCard').hidden = true; $('applyCard').hidden = true; return; }
    $('reviewTitle').textContent = state.source === 'starter' ? 'Review the starter set' : `Review ${state.sourceLabel}`;
    $('reviewHint').textContent = `${plural(state.rows.length, 'policy', 'policies')}. Every one starts as report-only. Nothing is written until you confirm.${state.lookupBusy ? ' Looking up names...' : ''}`;

    const tbody = document.querySelector('#reviewTable tbody');
    let risky = 0;
    tbody.innerHTML = state.rows.map((row) => {
      const body = bodyFor(row);
      const risk = riskFor(row);
      if (risk && row.selected) risky++;
      return `
      <tr data-row="${esc(row.rowId)}" class="${risk ? 'risk-row' : ''}">
        <td class="col-check"><input type="checkbox" data-act="pick" aria-label="Import ${esc(row.finalName)}" ${row.selected ? 'checked' : ''} /></td>
        <td class="col-policy">${policyCellHtml(row, risk)}</td>
        <td class="col-refs">${assignmentsHtml(row)}</td>
        <td class="col-controls">${controlsHtml(body)}</td>
        <td class="col-state">${stateCellHtml(row)}</td>
      </tr>
      <tr data-row="${esc(row.rowId)}" class="detail-row" ${row.open ? '' : 'hidden'}><td></td><td colspan="4">${row.open ? detailHtml(row, body) : ''}</td></tr>`;
    }).join('');
    $('selectAll').checked = state.rows.every((r) => r.selected);

    $('lockoutNotice').hidden = !risky;
    $('lockoutText').textContent = risky
      ? `${plural(risky, 'selected policy', 'selected policies')} can lock you out: all users, a block or no exclusions, and you are not excluded. Use Exclude me, or make sure an emergency account is excluded. Writing one of these as On needs an extra confirmation.`
      : '';
    renderResolver();
    updateApply();
  }

  // ---------- resolver ----------
  function missingRefs() {
    const out = new Map();
    for (const row of state.rows.filter((r) => r.selected)) {
      for (const ref of P().extractRefs(row.policy)) {
        if (!ref.isId || !P().MAPPABLE_KINDS.has(ref.kind)) continue;
        const id = ref.value.toLowerCase();
        if (!state.missing.has(id)) continue;
        if (!out.has(id)) out.set(id, { id, kind: ref.kind, policies: new Set() });
        out.get(id).policies.add(row.finalName);
      }
    }
    return [...out.values()];
  }

  // Tenant-specific ids that were never checked (lookup failed): writing them blind is not allowed.
  function uncheckedCount() {
    const ids = new Set();
    for (const row of state.rows.filter((r) => r.selected)) {
      for (const ref of P().extractRefs(row.policy)) {
        if (!ref.isId || !P().MAPPABLE_KINDS.has(ref.kind)) continue;
        const id = ref.value.toLowerCase();
        const known = ref.kind === 'location' ? !!state.locations : state.names.has(id);
        if (!known && !state.missing.has(id)) ids.add(id);
      }
    }
    return ids.size;
  }

  function unresolvedCount() {
    return missingRefs().filter((m) => !mapping()[m.id]).length;
  }

  function renderResolver() {
    const list = missingRefs();
    $('resolverCard').hidden = !list.length;
    document.querySelector('#resolverTable tbody').innerHTML = list.map((m) => {
      const val = state.resolverMap[m.id] || '';
      const hit = val ? state.replacements.get(val.toLowerCase()) : undefined;
      let status = '<span class="badge danger">Open</span>';
      if (val && !Http.isGuid(val)) status = '<span class="badge danger">Not a GUID</span>';
      else if (val && hit === undefined) status = '<span class="muted small">Looking up...</span>';
      else if (val && hit === null) status = `<span class="badge danger">No ${esc(KIND_LABEL[m.kind].toLowerCase())} with this id</span>`;
      else if (hit) status = `<span class="badge ok">${esc(hit.name)}</span>`;
      return `<tr data-src="${esc(m.id)}">
        <td>${esc(KIND_LABEL[m.kind])}</td>
        <td><code>${esc(m.id)}</code></td>
        <td class="small">${[...m.policies].map(esc).join('<br>')}</td>
        <td><input type="text" data-act="resolve" value="${esc(val)}" placeholder="Object id in this tenant" aria-label="Replacement for ${esc(m.id)}" /></td>
        <td>${status}</td>
      </tr>`;
    }).join('');
  }

  async function onResolveChange(e) {
    const input = e.target.closest('input[data-act="resolve"]');
    if (!input) return;
    const src = input.closest('tr').dataset.src;
    const val = input.value.trim().toLowerCase();
    if (!val) delete state.resolverMap[src];
    else state.resolverMap[src] = val;
    const kind = (missingRefs().find((m) => m.id === src) || {}).kind;
    if (val && Http.isGuid(val) && !state.replacements.has(val)) {
      renderResolver();
      try {
        let hit = null;
        if (kind === 'location') {
          const name = state.locations && state.locations.get(val);
          hit = name ? { name, kind } : null;
        } else {
          const found = await Graph.getDirectoryObjectsByIds([val]);
          const o = found[0];
          const foundKind = o ? (/group$/i.test(o['@odata.type'] || '') ? 'group' : /servicePrincipal$/i.test(o['@odata.type'] || '') ? 'servicePrincipal' : 'user') : null;
          hit = o && foundKind === kind ? { name: o.displayName || o.userPrincipalName, kind } : null;
        }
        state.replacements.set(val, hit);
      } catch (err) {
        Log.err('Lookup of replacement failed:', err);
        toast(`Lookup failed: ${err.message}`, 'err');
      }
    }
    renderReview();
  }

  // ---------- apply ----------
  function updateApply() {
    const selected = state.rows.filter((r) => r.selected);
    $('applyCard').hidden = !selected.length;
    if (!selected.length) return;
    const counts = { create: 0, overwrite: 0, skip: 0, ambiguous: 0 };
    selected.forEach((r) => { counts[rowAction(r)]++; });
    const unresolved = unresolvedCount();
    const parts = [`${plural(selected.length, 'policy', 'policies')} selected: ${counts.create} new, ${counts.overwrite} overwrite`];
    if (counts.skip + counts.ambiguous) parts.push(`${counts.skip + counts.ambiguous} skipped (name exists)`);
    if (unresolved) parts.push(`${plural(unresolved, 'missing object')} to resolve first`);
    const unchecked = state.lookupBusy ? 0 : uncheckedCount();
    if (unchecked) parts.push(`${plural(unchecked, 'id')} could not be checked, load the policies again`);
    $('applyHint').textContent = parts.join(', ');
    $('applyImportBtn').disabled = unresolved > 0 || unchecked > 0 || state.lookupBusy || counts.create + counts.overwrite === 0;
  }

  function diffTable(fields, [beforeLabel, afterLabel] = ['Existing', 'New']) {
    if (!fields.length) return h('p.muted', { text: 'No field differences, apart from the name.' });
    return h('div.table-wrap.diff-wrap', null, h('table.data-table.diff-table', null, [
      h('thead', null, h('tr', null, [h('th', { text: 'Field' }), h('th', { text: beforeLabel }), h('th', { text: afterLabel })])),
      h('tbody', null, fields.map((f) => h('tr', null, [
        h('td', null, h('code', { text: f.path })),
        h('td.diff-old', { text: shortJson(f.before) }),
        h('td.diff-new', { text: shortJson(f.after) }),
      ]))),
    ]));
  }

  // One dialog per overwrite with the field-level diff. Returns the rows the admin approved.
  async function confirmOverwrites(rows) {
    const approved = new Set();
    for (const [i, row] of rows.entries()) {
      const existing = row.matches[0];
      // Same cleaning as the import, so read-only noise does not show up as a change.
      const clean = P().sanitizePolicy(existing).policy || P().stripNoiseForDiff(existing);
      const before = P().normalizePolicyForCreateOrUpdate(clean, existing.displayName, existing.state);
      const fields = P().fieldDiff(before, bodyFor(row));
      const body = h('div', null, [
        h('p', { text: `"${existing.displayName}" (${existing.id}) exists in the tenant. This is what changes when you overwrite it.` }),
        diffTable(fields),
      ]);
      const choice = await modal({
        title: `Overwrite ${i + 1} of ${rows.length}: ${row.finalName}`,
        body, wide: true, danger: true,
        actions: [
          { label: 'Cancel all', value: 'cancel' },
          { label: 'Skip this one', value: 'skip' },
          { label: 'Overwrite', value: 'overwrite', variant: 'danger' },
        ],
      });
      if (choice === 'cancel' || choice === null) return null;
      if (choice === 'overwrite') approved.add(row);
    }
    return approved;
  }

  async function confirmLockouts(rows) {
    if (!rows.length) return true;
    const list = h('ul.tk-modal-list', null, rows.map((r) => h('li', null, [h('strong', { text: r.finalName }), ` ${riskFor(r).reason}`])));
    const body = h('div', null, [
      h('p', { text: 'These policies are set to On, apply to all users and do not exclude you. Written like this they can lock you and everyone else out of the tenant.' }),
      list,
      h('p.muted.small', { text: 'Exclude me adds your own account to the excluded users of these policies. Your emergency accounts should be excluded too.' }),
    ]);
    const choice = await modal({
      title: 'Lockout risk', body, danger: true, wide: true,
      actions: [
        { label: 'Cancel', value: null },
        ...(state.me.userId ? [{ label: 'Exclude me and continue', value: 'exclude', variant: 'primary', icon: 'i-user' }] : []),
        { label: 'Write without exclusion', value: 'force', variant: 'danger' },
      ],
    });
    if (choice === 'exclude') { rows.forEach((r) => { r.excludeMe = true; }); renderReview(); return true; }
    return choice === 'force';
  }

  async function applyImport() {
    if (!state.rows.some((r) => r.selected)) return;
    // Match against the tenant as it is now, not as it was at load time.
    await refreshTenantPolicies();
    applyNaming();
    renderReview();
    if (unresolvedCount() > 0 || uncheckedCount() > 0) { toast('Resolve the missing objects first.', 'warn'); return; }

    const selected = state.rows.filter((r) => r.selected);
    const writable = selected.filter((r) => ['create', 'overwrite'].includes(rowAction(r)));
    if (!writable.length) { toast('Nothing to write: every selected policy already exists.', 'warn'); return; }

    const items = selected.map((r) => {
      const a = rowAction(r);
      const risk = riskFor(r) ? ' [LOCKOUT RISK]' : '';
      if (a === 'create') return `Create "${r.finalName}" as ${STATE_LABEL[r.targetState]}${risk}`;
      if (a === 'overwrite') return `Overwrite "${r.finalName}" as ${STATE_LABEL[r.targetState]}${risk}`;
      return `Skip "${r.finalName}": a policy with this name exists`;
    });
    const anyOn = writable.some((r) => r.targetState === 'enabled');
    const ok = await confirm({
      title: `Write ${plural(writable.length, 'policy', 'policies')} to the tenant?`,
      body: anyOn ? 'Policies set to On are enforced right away.' : 'Report-only and off policies change nothing for users yet.',
      items,
      okLabel: 'Continue',
      danger: anyOn,
      wide: true,
    });
    if (!ok) return;

    if (!await confirmLockouts(writable.filter((r) => r.targetState === 'enabled' && riskFor(r)))) return;

    const overwrites = writable.filter((r) => rowAction(r) === 'overwrite');
    let approved = new Set();
    if (overwrites.length) {
      approved = await confirmOverwrites(overwrites);
      if (!approved) return;
    }
    const todo = writable.filter((r) => rowAction(r) === 'create' || approved.has(r));
    if (!todo.length) { toast('Nothing left to write.', 'info'); return; }

    // Consent before the first write, by popup, so this page and its state stay put.
    try { await Auth.getWriteToken(); }
    catch (e) { toast(`No write permission: ${e.message}`, 'err'); return; }

    const list = $('applyResults');
    list.hidden = false;
    list.replaceChildren();
    let okCount = 0;
    for (const row of todo) {
      const action = rowAction(row);
      const li = h('li', null, [h('span.warn-mark', { text: '...' }), ` ${row.finalName}`]);
      list.appendChild(li);
      try {
        const body = bodyFor(row);
        if (action === 'overwrite') await Graph.updateConditionalAccessPolicy(row.matches[0].id, body);
        else await Graph.createConditionalAccessPolicy(body);
        okCount++;
        li.replaceChildren(h('span.ok-mark', { text: '✓' }), ` ${action === 'overwrite' ? 'Overwritten' : 'Created'}: ${row.finalName} (${STATE_LABEL[row.targetState]})`);
        Log.info(`${action === 'overwrite' ? 'Overwrote' : 'Created'} "${row.finalName}" as ${row.targetState}`);
      } catch (e) {
        li.replaceChildren(h('span.err-mark', { text: '✗' }), ` ${row.finalName}: ${e.message}`);
        Log.err(`Writing "${row.finalName}" failed:`, e);
      }
    }
    const failed = todo.length - okCount;
    toast(`Wrote ${okCount} of ${todo.length}${failed ? `, ${failed} failed (see the list and the debug log)` : ''}`, failed ? 'warn' : 'ok');
    await refreshTenantPolicies();
  }

  // ---------- load import / starter ----------
  async function startReview(validated, source, label) {
    state.rows = toRows(validated);
    state.source = source;
    state.sourceLabel = label;
    state.resolverMap = {};
    $('overwriteExisting').checked = false;
    $('applyResults').hidden = true;
    applyNaming();
    renderReview();
    await lookupRefs();
    Log.info(`Loaded ${plural(state.rows.length, 'policy', 'policies')} from ${label}.`);
  }

  async function loadImportFile() {
    const file = $('importFile').files && $('importFile').files[0];
    if (!file) throw new Error('Pick a file first.');
    const { policies } = P().validateImport(await readTextFile(file));
    const dropped = policies.reduce((n, p) => n + p.dropped.length, 0);
    if (dropped) Log.info(`Dropped ${plural(dropped, 'field')} from ${file.name}:`, policies.flatMap((p) => p.dropped.map((d) => `${p.policy.displayName}: ${d.path} (${d.reason})`)));
    await startReview(policies, 'import', file.name);
  }

  async function loadStarterSet() {
    const validated = P().buildStarterPolicies().map((p, i) => P().sanitizePolicy(p, i));
    await startReview(validated, 'starter', 'the starter set');
  }

  function onReviewClick(e) {
    const btn = e.target.closest('button[data-act]');
    if (!btn) return;
    const row = state.rows.find((r) => r.rowId === btn.closest('tr').dataset.row);
    if (!row) return;
    if (btn.dataset.act === 'details') row.open = !row.open;
    if (btn.dataset.act === 'exclude-me') row.excludeMe = true;
    if (btn.dataset.act === 'include-me') row.excludeMe = false;
    renderReview();
  }
  function onReviewChange(e) {
    const el = e.target.closest('[data-act]');
    if (!el) return;
    const row = state.rows.find((r) => r.rowId === el.closest('tr').dataset.row);
    if (!row) return;
    if (el.dataset.act === 'pick') row.selected = el.checked;
    if (el.dataset.act === 'state' && P().STATES.includes(el.value)) row.targetState = el.value;
    renderReview();
  }

  // ---------- diff ----------
  async function runDiff() {
    const fOld = $('diffFileOld').files[0];
    const fNew = $('diffFileNew').files[0];
    if (!fOld || !fNew) return;
    const [oldText, newText] = await Promise.all([readTextFile(fOld), readTextFile(fNew)]);
    const diff = P().comparePolicySets(P().parseImportPayload(oldText).policies, P().parseImportPayload(newText).policies);

    const sum = $('diffSummary');
    sum.hidden = false;
    sum.innerHTML = [
      `<span class="badge ok">${diff.added.length} added</span>`,
      `<span class="badge danger">${diff.removed.length} removed</span>`,
      `<span class="badge warn">${diff.changed.length} changed</span>`,
      `<span class="badge neutral">${diff.unchanged.length} unchanged</span>`,
    ].join('');

    const wrap = $('diffWrap');
    wrap.replaceChildren();
    const section = (title, names, kind) => {
      if (!names.length) return;
      wrap.appendChild(h('h3', { text: title }));
      wrap.appendChild(h('ul.action-result-list', null, names.map((n) => h('li', null, [h('span', { class: `badge ${kind}`, text: title.toLowerCase() }), ` ${n}`]))));
    };
    section('Added', diff.added, 'ok');
    section('Removed', diff.removed, 'danger');
    if (diff.changed.length) {
      wrap.appendChild(h('h3', { text: 'Changed' }));
      for (const c of diff.changed) {
        wrap.appendChild(h('details', null, [
          h('summary', null, [c.name, h('span.badge.warn', { text: plural(c.fields.length, 'field') })]),
          diffTable(c.fields, ['Old export', 'New export']),
        ]));
      }
    }
    if (!diff.added.length && !diff.removed.length && !diff.changed.length) wrap.appendChild(h('p.muted', { text: 'No differences.' }));
    Log.info(`Diff: ${diff.added.length} added, ${diff.removed.length} removed, ${diff.changed.length} changed, ${diff.unchanged.length} unchanged.`);
  }

  // ---------- boot ----------
  function guarded(label, fn) {
    return async (e) => {
      const btn = e && e.currentTarget instanceof HTMLButtonElement ? e.currentTarget : null;
      await busy(btn, async () => {
        try { await fn(); }
        catch (err) { Log.err(`${label} failed:`, err); toast(`${label} failed: ${err.message}`, 'err'); }
      });
    };
  }

  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    Toolkit.tabs('#tabs');
    loadPrefs();

    $('signInBtn').addEventListener('click', onSignIn);
    $('refreshPoliciesBtn').addEventListener('click', guarded('Loading policies', refreshTenantPolicies));
    $('exportAllBtn').addEventListener('click', guarded('Export', async () => {
      if (!state.tenantPolicies.length) await refreshTenantPolicies();
      exportJson();
    }));
    $('exportXlsxBtn').addEventListener('click', guarded('Excel export', () => Export.xlsx([{ name: 'CA policies', rows: state.tenantPolicies, columns: EXPORT_COLUMNS }], `conditional-access-${Export.stamp()}.xlsx`)));
    $('exportCsvBtn').addEventListener('click', () => Export.csv(state.tenantPolicies, { columns: EXPORT_COLUMNS, filename: `conditional-access-${Export.stamp()}.csv` }));
    $('copyNamesBtn').addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(state.tenantPolicies.map((p) => `${p.displayName || ''}\t${STATE_LABEL[p.state] || p.state || ''}`).join('\n'));
        toast('Names copied', 'ok');
      } catch { toast('Copy failed', 'warn'); }
    });

    $('importFile').addEventListener('change', () => { $('loadImportBtn').disabled = !$('importFile').files[0]; });
    $('loadImportBtn').addEventListener('click', guarded('Loading the file', loadImportFile));
    $('loadStarterBtn').addEventListener('click', guarded('Loading the starter set', loadStarterSet));
    ['namePrefix', 'nameSuffix'].forEach((id) => $(id).addEventListener('input', () => { savePrefs(); applyNaming(); renderReview(); }));
    $('overwriteExisting').addEventListener('change', renderReview);
    $('bulkState').addEventListener('change', () => {
      const v = $('bulkState').value;
      if (P().STATES.includes(v)) state.rows.filter((r) => r.selected).forEach((r) => { r.targetState = v; });
      $('bulkState').value = '';
      renderReview();
    });
    $('selectAll').addEventListener('change', (e) => { state.rows.forEach((r) => { r.selected = e.target.checked; }); renderReview(); });
    const reviewBody = document.querySelector('#reviewTable tbody');
    reviewBody.addEventListener('click', onReviewClick);
    reviewBody.addEventListener('change', onReviewChange);
    document.querySelector('#resolverTable tbody').addEventListener('change', onResolveChange);
    $('applyImportBtn').addEventListener('click', guarded('Writing policies', applyImport));

    const diffReady = () => { $('runDiffBtn').disabled = !($('diffFileOld').files[0] && $('diffFileNew').files[0]); };
    $('diffFileOld').addEventListener('change', diffReady);
    $('diffFileNew').addEventListener('change', diffReady);
    $('runDiffBtn').addEventListener('click', guarded('Diff', runDiff));

    try {
      await Auth.init();
      refreshAuthUI();
      if (Auth.getAccount()) {
        guarded('Loading policies', refreshTenantPolicies)();
        loadMe();
      }
      Log.info('ConditionalAccess ready.');
    } catch (e) {
      Log.err('Start-up failed:', e);
      toast(`Start-up failed: ${e.message}`, 'err');
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
