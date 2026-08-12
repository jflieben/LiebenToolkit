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

  window.Spns = {
    MS_FIRST_PARTY_TENANT,
    GRAPH_APP_ID, SPO_APP_ID, SITE_ROLE_OPTIONS,
    classifySpn, isMicrosoftFirstParty,
    hasAssignableAppRoles, hasAssignableDelegatedScopes, hasAssignablePermissions,
    buildPermissionIndex,
    parseScopeString, joinScopeString,
    summarizePermissionRemoval,
    hasSitesSelected, extractSiteAppPermissions, findAppPermissionOnSite,
  };
})();
