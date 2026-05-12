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
    return sp && sp.appOwnerOrganizationId === MS_FIRST_PARTY_TENANT;
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

  window.Spns = {
    MS_FIRST_PARTY_TENANT,
    classifySpn, isMicrosoftFirstParty,
    buildPermissionIndex,
    parseScopeString, joinScopeString,
  };
})();
