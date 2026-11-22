// Pure helpers used by SPNRoleMgr - safe to unit-test.
(() => {
  /**
   * Classify a service principal into a friendly category. The Graph
   * `servicePrincipalType` only distinguishes Application vs ManagedIdentity
   * vs Legacy/SocialIdp, so we additionally use `tags` and `appOwnerOrganizationId`
   * to tell apart Microsoft first-party apps, third-party gallery apps and
   * customer-owned app registrations.
   *
   * Microsoft owns tenant id `f8cdef31-a31e-4b4a-93e4-5f571e91255a`.
   */
  const MS_FIRST_PARTY_TENANT = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a';

  // Well-known resource appIds used to spot Sites.Selected assignments.
  const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';
  const SPO_APP_ID = '00000003-0000-0ff1-ce00-000000000000';

  // Roles accepted by POST/PATCH /sites/{id}/permissions. read/write are fully
  // documented; owner/fullcontrol/manage work in most tenants but can be rejected
  // depending on SharePoint configuration - the GUI warns about that.
  const SITE_ROLE_OPTIONS = [
    { value: 'read', label: 'Read', hint: 'View site content' },
    { value: 'write', label: 'Write', hint: 'View and edit site content' },
    { value: 'owner', label: 'Owner', hint: 'Site owner rights', extended: true },
    { value: 'fullcontrol', label: 'Full control', hint: 'Full control of the site', extended: true },
    { value: 'manage', label: 'Manage', hint: 'Manage site settings', extended: true },
  ];

  function classifySpn(sp) {
    const t = sp.servicePrincipalType || '';
    if (t === 'ManagedIdentity') {
      const tags = (sp.tags || []).map(x => String(x).toLowerCase());
      if (tags.some(x => x.includes('userassigned'))) return 'Managed identity (user-assigned)';
      return 'Managed identity (system-assigned)';
    }
    if (t === 'SocialIdp') return 'Social identity provider';
    if (t === 'Legacy') return 'Legacy';
    // Application
    if (sp.appOwnerOrganizationId === MS_FIRST_PARTY_TENANT) return 'Microsoft first-party';
    const tags = (sp.tags || []).map(x => String(x).toLowerCase());
    if (tags.includes('windowsazureactivedirectoryintegratedapp') || tags.includes('webapp') || tags.includes('hideapp')) {
      // Could be tenant-owned or gallery; gallery apps usually have a publisherName.
      if (sp.publisherName && sp.appOwnerOrganizationId && sp.appOwnerOrganizationId !== MS_FIRST_PARTY_TENANT) return 'Gallery / third-party';
    }
    return 'Application';
  }

  function isMicrosoftFirstParty(sp) {
    return !!(sp && sp.appOwnerOrganizationId === MS_FIRST_PARTY_TENANT);
  }

  /** True when the SPN exposes at least one app role assignable to applications. */
  function hasAssignableAppRoles(sp) {
    const roles = (sp && sp.appRoles) || [];
    return roles.some(r => r && r.isEnabled !== false && Array.isArray(r.allowedMemberTypes) && r.allowedMemberTypes.includes('Application'));
  }

  /** True when the SPN exposes at least one enabled delegated scope. */
  function hasAssignableDelegatedScopes(sp) {
    const scopes = (sp && sp.oauth2PermissionScopes) || [];
    return scopes.some(s => s && s.isEnabled !== false);
  }

  /** True when the SPN exposes at least one assignable API permission. */
  function hasAssignablePermissions(sp) {
    return hasAssignableAppRoles(sp) || hasAssignableDelegatedScopes(sp);
  }

  /**
   * Build a lookup map { roleId -> { name, description, type } } from a resource SPN.
   * `type` is 'AppRole' for application permissions, 'OAuth2' for delegated.
   */
  function buildPermissionIndex(resourceSpn) {
    const idx = new Map();
    if (!resourceSpn) return idx;
    for (const r of (resourceSpn.appRoles || [])) {
      idx.set(r.id, {
        id: r.id,
        name: r.value || r.displayName || r.id,
        displayName: r.displayName || r.value || r.id,
        description: r.description || '',
        type: 'AppRole',
      });
    }
    for (const s of (resourceSpn.oauth2PermissionScopes || [])) {
      idx.set(s.id, {
        id: s.id,
        name: s.value || s.adminConsentDisplayName || s.id,
        displayName: s.adminConsentDisplayName || s.value || s.id,
        description: s.adminConsentDescription || s.userConsentDescription || '',
        type: 'OAuth2',
      });
    }
    return idx;
  }

  /** Parse a space-separated oauth2 grant.scope string into trimmed names. */
  function parseScopeString(scope) {
    if (!scope) return [];
    return String(scope).split(/\s+/).filter(Boolean);
  }
  function joinScopeString(scopes) {
    return [...new Set(scopes.filter(Boolean).map(s => s.trim()))].join(' ');
  }

  /** Summarize how many permission entries would be removed for a selected SPN. */
  function summarizePermissionRemoval(appRoleAssignments, oauth2Grants) {
    const appRoleCount = Array.isArray(appRoleAssignments) ? appRoleAssignments.length : 0;
    const grantCount = Array.isArray(oauth2Grants) ? oauth2Grants.length : 0;
    const delegatedScopeCount = (Array.isArray(oauth2Grants) ? oauth2Grants : []).reduce(
      (total, grant) => total + parseScopeString(grant && grant.scope).length,
      0
    );
    return {
      appRoleCount,
      grantCount,
      delegatedScopeCount,
      totalPermissionEntries: appRoleCount + grantCount,
    };
  }

  /**
   * True when the SPN holds a `Sites.Selected` application permission on Microsoft
   * Graph or SharePoint Online. Reuses the already-resolved resource + index caches
   * so no extra Graph calls and no hardcoded role GUIDs are needed.
   *   appRoleAssignments : Array of appRoleAssignment ({ resourceId, appRoleId })
   *   resourceCache      : Map(resourceObjectId -> resource SPN with .appId)
   *   resourceIndexCache : Map(resourceObjectId -> Map(roleId -> { name, ... }))
   */
  function hasSitesSelected(appRoleAssignments, resourceCache, resourceIndexCache) {
    if (!Array.isArray(appRoleAssignments)) return false;
    const rc = resourceCache || new Map();
    const ric = resourceIndexCache || new Map();
    for (const a of appRoleAssignments) {
      if (!a || !a.resourceId) continue;
      const resource = rc.get ? rc.get(a.resourceId) : null;
      const appId = resource && resource.appId ? String(resource.appId).toLowerCase() : '';
      if (appId !== GRAPH_APP_ID && appId !== SPO_APP_ID) continue;
      const idx = ric.get ? ric.get(a.resourceId) : null;
      const meta = idx && idx.get ? idx.get(a.appRoleId) : null;
      const name = meta && meta.name ? String(meta.name).toLowerCase() : '';
      if (name === 'sites.selected') return true;
    }
    return false;
  }

  /** Normalize a site's /permissions response into flat app-permission rows. */
  function extractSiteAppPermissions(sitePermissions) {
    const rows = [];
    for (const p of (Array.isArray(sitePermissions) ? sitePermissions : [])) {
      if (!p) continue;
      const identities = p.grantedToIdentitiesV2 || p.grantedToIdentities || [];
      for (const idn of identities) {
        const app = idn && idn.application;
        if (!app || !app.id) continue;
        rows.push({
          permissionId: p.id,
          appId: String(app.id),
          displayName: app.displayName || '',
          roles: Array.isArray(p.roles) ? p.roles.slice() : [],
        });
      }
    }
    return rows;
  }

  /** Return the site permission row that targets the given application appId, or null. */
  function findAppPermissionOnSite(sitePermissions, appId) {
    if (!appId) return null;
    const needle = String(appId).toLowerCase();
    return extractSiteAppPermissions(sitePermissions).find(r => r.appId.toLowerCase() === needle) || null;
  }

  // ---------- API picker ----------
  const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

  /** Datalist value for an API: the name alone is not unique, the appId is. */
  function apiOptionLabel(sp) {
    return `${(sp && sp.displayName) || '(no name)'} (${(sp && sp.appId) || '?'})`;
  }

  /**
   * Resolve what the user typed in the API box. Accepts "Name (appId)", a bare
   * appId, or an exact display name that only one API carries. No substring
   * matching: picking the wrong resource means granting roles on the wrong API.
   * Returns { sp } or { error, matches }.
   */
  function resolveApiQuery(apiSpns, query) {
    const q = String(query == null ? '' : query).trim();
    if (!q) return { error: 'Enter an API name or appId.', matches: [] };
    const list = Array.isArray(apiSpns) ? apiSpns : [];
    const m = /\(([0-9a-f-]{36})\)\s*$/i.exec(q);
    const appId = m ? m[1] : (GUID_RE.test(q) ? q : null);
    if (appId) {
      const hits = list.filter((x) => (x.appId || '').toLowerCase() === appId.toLowerCase());
      if (hits.length === 1) return { sp: hits[0] };
      return { error: `No API with appId ${appId} that exposes assignable permissions.`, matches: hits };
    }
    const lower = q.toLowerCase();
    const hits = list.filter((x) => (x.displayName || '').toLowerCase() === lower);
    if (hits.length === 1) return { sp: hits[0] };
    if (hits.length > 1) return { error: `${hits.length} APIs are called "${q}". Pick one from the list, it shows the appId.`, matches: hits };
    return { error: 'No API with that name. Pick one from the list.', matches: [] };
  }

  /** Who owns an API, for status lines and confirm dialogs. */
  function ownerLabel(sp, tenantId) {
    if (!sp) return 'unknown owner';
    const owner = String(sp.appOwnerOrganizationId || '').toLowerCase();
    if (owner === MS_FIRST_PARTY_TENANT) return 'Microsoft';
    if (owner && tenantId && owner === String(tenantId).toLowerCase()) return 'your tenant';
    if (sp.publisherName) return `${sp.publisherName}${owner ? `, tenant ${owner}` : ''}`;
    return owner ? `tenant ${owner}` : 'unknown owner';
  }

  // ---------- tenant guard ----------
  /** tid claim of an access token, or null when it is not a readable JWT. */
  function tokenTenantId(token) {
    try {
      const part = String(token).split('.')[1];
      if (!part) return null;
      const b64 = part.replace(/-/g, '+').replace(/_/g, '/');
      const json = atob(b64 + '='.repeat((4 - (b64.length % 4)) % 4));
      const tid = JSON.parse(json).tid;
      return tid ? String(tid) : null;
    } catch { return null; }
  }

  /** Throws when a token belongs to another tenant than the one this page loaded. */
  function assertTokenTenant(token, expectedTenantId) {
    if (!expectedTenantId) return;
    const tid = tokenTenantId(token);
    if (tid && tid.toLowerCase() !== String(expectedTenantId).toLowerCase()) {
      throw new Error(`The token is for tenant ${tid}, this page works in ${expectedTenantId}. Nothing was changed. Reload the page and sign in with the right account.`);
    }
  }

  // ---------- SharePoint URLs ----------
  function isSharePointHost(host) {
    const h = String(host || '').toLowerCase();
    return /^[a-z0-9-]+(\.[a-z0-9-]+)*\.sharepoint\.com$/.test(h);
  }

  /** https URL on *.sharepoint.com, or null. */
  function parseSharePointUrl(text) {
    let u;
    try { u = new URL(String(text || '').trim()); } catch { return null; }
    if (u.protocol !== 'https:' || u.username || u.password || u.port || !isSharePointHost(u.hostname)) return null;
    let segments;
    try { segments = u.pathname.split('/').filter(Boolean).map(decodeURIComponent); } catch { return null; }
    return { host: u.hostname.toLowerCase(), segments };
  }

  /**
   * Graph path (relative to /v1.0) of the site collection a URL belongs to, e.g.
   * /sites/contoso.sharepoint.com:/sites/Marketing. Sites.Selected grants live on
   * the collection, so subsites and library URLs resolve to their collection.
   * Returns null for anything that is not a SharePoint Online URL.
   */
  function siteCollectionPath(url) {
    const p = parseSharePointUrl(url);
    if (!p) return null;
    const host = encodeURIComponent(p.host);
    if (p.segments.length < 2) return `/sites/${host}`;
    return `/sites/${host}:/${p.segments.slice(0, 2).map(encodeURIComponent).join('/')}`;
  }

  /** Graph site id: "<host>,<site guid>,<web guid>" on a SharePoint Online host. */
  function isValidSiteId(id) {
    const parts = String(id || '').split(',');
    return parts.length === 3 && isSharePointHost(parts[0]) && GUID_RE.test(parts[1]) && GUID_RE.test(parts[2]);
  }

  // ---------- confirm dialog text ----------
  /** [label, value] rows that head every change dialog. */
  function confirmFacts({ tenant, spn, api, site }) {
    const rows = [];
    rows.push(['Tenant', tenant ? `${tenant.domain || '(unknown domain)'} (${tenant.tenantId || 'unknown id'})` : '(not signed in)']);
    if (spn) rows.push(['Service principal', `${spn.displayName || '(no name)'} (appId ${spn.appId || '?'})`]);
    if (api) rows.push(['API', `${api.displayName || '(no name)'} (appId ${api.appId || '?'}), owned by ${ownerLabel(api, tenant && tenant.tenantId)}`]);
    if (site) rows.push(['Site', `${site.displayName || site.name || '(no name)'} ${site.webUrl ? `(${site.webUrl})` : ''}`.trim()]);
    return rows;
  }

  /** One readable line per application permission, for confirm lists. */
  function appRoleLabel(assignment, resourceCache, resourceIndexCache) {
    const res = resourceCache && resourceCache.get ? resourceCache.get(assignment.resourceId) : null;
    const idx = resourceIndexCache && resourceIndexCache.get ? resourceIndexCache.get(assignment.resourceId) : null;
    const meta = idx && idx.get ? idx.get(assignment.appRoleId) : null;
    const api = (res && res.displayName) || assignment.resourceDisplayName || assignment.resourceId;
    return `${api}: ${meta ? meta.name : assignment.appRoleId}`;
  }

  /** One line per delegated grant, scopes included. */
  function grantLabel(grant, resourceCache) {
    const res = resourceCache && resourceCache.get ? resourceCache.get(grant.resourceId) : null;
    const api = (res && res.displayName) || grant.resourceId;
    const scopes = parseScopeString(grant.scope);
    const who = grant.consentType === 'AllPrincipals' ? 'all users' : `user ${grant.principalId || '?'}`;
    return `${api}: ${scopes.length ? scopes.join(' ') : '(no scopes)'} (${who})`;
  }

  /** Every line "Revoke all" removes. */
  function revokeAllItems(appRoleAssignments, oauth2Grants, resourceCache, resourceIndexCache) {
    return [
      ...(appRoleAssignments || []).map((a) => `Remove application permission ${appRoleLabel(a, resourceCache, resourceIndexCache)}`),
      ...(oauth2Grants || []).map((g) => `Remove delegated grant ${grantLabel(g, resourceCache)}`),
    ];
  }

  // ---------- bulk revoke of delegated scopes ----------
  /** Selection key of one scope row. A grant without scopes has scope ''. */
  function scopeKey(grantId, scope) { return `${grantId}|${scope || ''}`; }

  /**
   * Per grant: which selected scopes go and which stay. A grant left without scopes is
   * deleted, otherwise it is patched. Returns [{ grant, remove, keep, action }].
   */
  function planScopeRevocation(oauth2Grants, selectedKeys) {
    const sel = selectedKeys instanceof Set ? selectedKeys : new Set(selectedKeys || []);
    const plan = [];
    for (const grant of (oauth2Grants || [])) {
      if (!grant) continue;
      const scopes = parseScopeString(grant.scope);
      if (!scopes.length) {
        if (sel.has(scopeKey(grant.id, ''))) plan.push({ grant, remove: [], keep: [], action: 'delete' });
        continue;
      }
      const remove = scopes.filter((s) => sel.has(scopeKey(grant.id, s)));
      if (!remove.length) continue;
      const keep = scopes.filter((s) => !remove.includes(s));
      plan.push({ grant, remove, keep, action: keep.length ? 'update' : 'delete' });
    }
    return plan;
  }

  // ---------- Exchange hint ----------
  const EXO_APP_ID = '00000002-0000-0ff1-ce00-000000000000';
  const MAILBOX_ROLE_RE = /^(mail|calendars|contacts|mailboxsettings|mailboxfolder|mailboxitem|full_access_as_app)/i;

  /** True when the SPN holds an application permission that reaches mailboxes. */
  function hasMailboxPermissions(appRoleAssignments, resourceCache, resourceIndexCache) {
    for (const a of (appRoleAssignments || [])) {
      const res = resourceCache && resourceCache.get ? resourceCache.get(a.resourceId) : null;
      const appId = String((res && res.appId) || '').toLowerCase();
      if (appId === EXO_APP_ID) return true;
      if (appId !== GRAPH_APP_ID) continue;
      const idx = resourceIndexCache && resourceIndexCache.get ? resourceIndexCache.get(a.resourceId) : null;
      const meta = idx && idx.get ? idx.get(a.appRoleId) : null;
      if (meta && MAILBOX_ROLE_RE.test(meta.name || '')) return true;
    }
    return false;
  }

  // ---------- site list ----------
  /** Lowercase site collection URL, the key sites are matched and cached on. */
  function siteKey(url) {
    const p = parseSharePointUrl(url);
    if (!p) return null;
    return `https://${p.host}${p.segments.length >= 2 ? `/${p.segments.slice(0, 2).join('/')}` : ''}`.toLowerCase();
  }

  /**
   * Substring filter on name and URL, every word must match. A pasted URL matches the
   * site collection it belongs to.
   */
  function filterSites(sites, query) {
    const list = Array.isArray(sites) ? sites : [];
    const q = String(query == null ? '' : query).trim().toLowerCase();
    if (!q || q === '*') return list.slice();
    const asUrl = /^https?:\/\//.test(q) ? siteKey(q) : null;
    if (asUrl) return list.filter((s) => siteKey(s.webUrl) === asUrl);
    const words = q.split(/\s+/).filter(Boolean);
    return list.filter((s) => {
      const hay = `${s.displayName || ''}\n${s.webUrl || ''}`.toLowerCase();
      return words.every((w) => hay.includes(w));
    });
  }

  /** Sites the SPN can reach first, in their original order otherwise. access: Map(siteKey -> { status }). */
  function sortSitesByAccess(sites, access) {
    const rank = (s) => {
      const a = access && access.get ? access.get(siteKey(s.webUrl)) : null;
      return a && a.status === 'yes' ? 0 : 1;
    };
    return (sites || []).map((s, i) => ({ s, i, r: rank(s) })).sort((x, y) => x.r - y.r || x.i - y.i).map((x) => x.s);
  }

  /** Graph refused because of the user's own rights on the site. */
  function isAccessDenied(e) {
    if (!e) return false;
    return e.status === 401 || e.status === 403 || /accessDenied/i.test(e.code || '');
  }

  /** Site permission rows picked for removal, one per permission object. */
  function sitePermsForRemoval(sitePermissions, permissionIds) {
    const ids = permissionIds instanceof Set ? permissionIds : new Set(permissionIds || []);
    const byId = new Map();
    for (const r of extractSiteAppPermissions(sitePermissions)) {
      if (!ids.has(r.permissionId)) continue;
      if (!byId.has(r.permissionId)) byId.set(r.permissionId, { permissionId: r.permissionId, apps: [], roles: r.roles });
      byId.get(r.permissionId).apps.push({ appId: r.appId, displayName: r.displayName });
    }
    return [...byId.values()];
  }

  // ---------- SharePoint admin ----------
  /** contoso.sharepoint.com (or -my) -> contoso-admin.sharepoint.com, else null. */
  function adminHostOf(spoHost) {
    const m = /^([a-z0-9]+)(?:-my|-admin)?\.sharepoint\.com$/i.exec(String(spoHost || ''));
    return m ? `${m[1].toLowerCase()}-admin.sharepoint.com` : null;
  }

  function claimLoginName(upn) {
    const u = String(upn || '').trim();
    if (!u) throw new Error('No account name to make site admin.');
    return /^i:0/i.test(u) ? u : `i:0#.f|membership|${u}`;
  }

  // ---------- temporary site admin records ----------
  // A record is written before SetSiteAdmin runs, so a closed tab never hides one.
  function sameAccount(rec, tenantId, upn) {
    return !!rec && String(rec.tenantId || '').toLowerCase() === String(tenantId || '').toLowerCase() &&
      String(rec.upn || '').toLowerCase() === String(upn || '').toLowerCase();
  }
  function addElevation(list, rec) {
    const key = siteKey(rec.siteUrl);
    const rest = (list || []).filter((r) => !(siteKey(r.siteUrl) === key && sameAccount(r, rec.tenantId, rec.upn)));
    return [...rest, { ...rec, siteUrl: key }];
  }
  function dropElevation(list, siteUrl, tenantId, upn) {
    const key = siteKey(siteUrl);
    return (list || []).filter((r) => !(siteKey(r.siteUrl) === key && sameAccount(r, tenantId, upn)));
  }
  function elevationsFor(list, tenantId, upn) {
    return (list || []).filter((r) => sameAccount(r, tenantId, upn));
  }

  window.Spns = {
    MS_FIRST_PARTY_TENANT,
    GRAPH_APP_ID, SPO_APP_ID, SITE_ROLE_OPTIONS,
    classifySpn, isMicrosoftFirstParty,
    hasAssignableAppRoles, hasAssignableDelegatedScopes, hasAssignablePermissions,
    buildPermissionIndex,
    parseScopeString, joinScopeString,
    summarizePermissionRemoval,
    hasSitesSelected, extractSiteAppPermissions, findAppPermissionOnSite,
    apiOptionLabel, resolveApiQuery, ownerLabel,
    tokenTenantId, assertTokenTenant,
    isSharePointHost, parseSharePointUrl, siteCollectionPath, isValidSiteId,
    confirmFacts, appRoleLabel, grantLabel, revokeAllItems,
    scopeKey, planScopeRevocation,
    EXO_APP_ID, hasMailboxPermissions,
    siteKey, filterSites, sortSitesByAccess, isAccessDenied, sitePermsForRemoval,
    adminHostOf, claimLoginName,
    addElevation, dropElevation, elevationsFor,
  };
})();
