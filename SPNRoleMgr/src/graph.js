// Microsoft Graph calls for SPNRoleMgr: service principals, app role assignments,
// delegated grants and Sites.Selected site permissions.
(() => {
  const V1 = 'https://graph.microsoft.com/v1.0';
  const R = Http.raw(V1);

  // Every token is checked against the tenant this page signed in to, so a consent
  // popup that ends up in another account cannot write to another tenant.
  let pinnedTenant = null;
  function pinTenant(tenantId) { pinnedTenant = tenantId || null; }
  function guarded(get) {
    return async () => {
      const token = await get();
      Spns.assertTokenTenant(token, pinnedTenant);
      return token;
    };
  }
  const TOKENS = {
    read: guarded(() => Auth.getReadToken()),
    write: guarded(() => Auth.getWriteToken()),
    sites: guarded(() => Auth.getToken('sites')),
  };

  function call(url, { kind = 'read', ...opts } = {}) {
    return Http.json(url, { ...opts, label: 'Graph', token: TOKENS[kind] });
  }
  function pageAll(url, { kind = 'read', ...opts } = {}, onPage) {
    return Http.pageAll(url, { ...opts, label: 'Graph', token: TOKENS[kind] }, onPage);
  }

  const SPN_FIELDS = [
    'id', 'appId', 'displayName', 'servicePrincipalType', 'accountEnabled', 'tags',
    'appOwnerOrganizationId', 'homepage', 'publisherName', 'signInAudience',
    'appDisplayName', 'createdDateTime', 'appRoles', 'oauth2PermissionScopes',
  ].join(',');

  // onProgress(loadedCount, pageNumber)
  function listAllServicePrincipals(onProgress, signal) {
    return pageAll(`${V1}/servicePrincipals?$top=999&$select=${SPN_FIELDS}`,
      { headers: { ConsistencyLevel: 'eventual' }, signal },
      (items, page, total) => onProgress && onProgress(total, page));
  }

  const getServicePrincipal = (id) => call(Http.url`${R}/servicePrincipals/${id}`);

  // ---------- application permissions ----------
  const listAppRoleAssignments = (spId) => pageAll(Http.url`${R}/servicePrincipals/${spId}/appRoleAssignments`);
  const addAppRoleAssignment = (spId, { resourceId, appRoleId }) => call(Http.url`${R}/servicePrincipals/${spId}/appRoleAssignments`, {
    method: 'POST', body: { principalId: spId, resourceId, appRoleId }, kind: 'write',
  });
  const removeAppRoleAssignment = (spId, assignmentId) => call(Http.url`${R}/servicePrincipals/${spId}/appRoleAssignments/${assignmentId}`, {
    method: 'DELETE', kind: 'write',
  });

  // ---------- delegated permissions ----------
  const listOAuth2Grants = (spId) => pageAll(Http.url`${R}/servicePrincipals/${spId}/oauth2PermissionGrants`);
  const createOAuth2Grant = ({ clientId, resourceId, scope, consentType = 'AllPrincipals', principalId = null }) => call(`${V1}/oauth2PermissionGrants`, {
    method: 'POST', body: { clientId, consentType, principalId, resourceId, scope }, kind: 'write',
  });
  const updateOAuth2Grant = (grantId, { scope }) => call(Http.url`${R}/oauth2PermissionGrants/${grantId}`, {
    method: 'PATCH', body: { scope }, kind: 'write',
  });
  const deleteOAuth2Grant = (grantId) => call(Http.url`${R}/oauth2PermissionGrants/${grantId}`, { method: 'DELETE', kind: 'write' });

  // ---------- SharePoint site permissions (Sites.Selected) ----------
  // Need delegated Sites.FullControl.All and a SharePoint Administrator.
  function siteId(id) {
    if (!Spns.isValidSiteId(id)) throw new Error(`Not a SharePoint Online site id: ${id}`);
    return id;
  }

  // First page only: search-as-you-type, a full crawl would take ages in big tenants.
  async function searchSites(query, top = 100) {
    const q = String(query == null ? '' : query).trim() || '*';
    const data = await call(`${V1}/sites?search=${encodeURIComponent(q)}&$top=${top}&$select=id,displayName,name,webUrl`, { kind: 'sites' });
    return (data && Array.isArray(data.value) ? data.value : []).filter((s) => Spns.isValidSiteId(s.id));
  }

  // Site collection of a SharePoint Online URL (subsites and library URLs included).
  function getSiteCollectionByUrl(url) {
    const path = Spns.siteCollectionPath(url);
    if (!path) return Promise.reject(new Error('Only https URLs on *.sharepoint.com are accepted.'));
    return call(`${V1}${path}?$select=id,displayName,name,webUrl`, { kind: 'sites' });
  }

  const getSitePermissions = (id) => pageAll(Http.url`${R}/sites/${siteId(id)}/permissions`, { kind: 'sites' });
  const addSitePermission = (id, { appId, displayName, roles }) => call(Http.url`${R}/sites/${siteId(id)}/permissions`, {
    method: 'POST', body: { roles, grantedToIdentities: [{ application: { id: appId, displayName } }] }, kind: 'sites',
  });
  const updateSitePermission = (id, permissionId, { roles }) => call(Http.url`${R}/sites/${siteId(id)}/permissions/${permissionId}`, {
    method: 'PATCH', body: { roles }, kind: 'sites',
  });
  const deleteSitePermission = (id, permissionId) => call(Http.url`${R}/sites/${siteId(id)}/permissions/${permissionId}`, { method: 'DELETE', kind: 'sites' });

  // Graph site search matches whole words. No hits: try again as a prefix.
  async function searchSitesLoose(query) {
    const q = String(query == null ? '' : query).trim();
    const hits = await searchSites(q);
    if (hits.length || !q || q === '*' || /[*"]/.test(q)) return hits;
    return searchSites(`${q}*`);
  }

  window.Graph = {
    pinTenant, guarded, call, pageAll, tokens: TOKENS, searchSitesLoose,
    listAllServicePrincipals, getServicePrincipal,
    listAppRoleAssignments, addAppRoleAssignment, removeAppRoleAssignment,
    listOAuth2Grants, createOAuth2Grant, updateOAuth2Grant, deleteOAuth2Grant,
    searchSites, getSiteCollectionByUrl, getSitePermissions, addSitePermission, updateSitePermission, deleteSitePermission,
  };
})();
