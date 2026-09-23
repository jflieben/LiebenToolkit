// Microsoft Graph client for SPNRoleMgr.
// Handles paging, retry/backoff and the specific endpoints we need to inspect
// and mutate service principal permissions (app role assignments + delegated
// oauth2 permission grants).
(() => {
  const BASE = 'https://graph.microsoft.com/v1.0';

  async function _tokenFor({ write = false, sites = false } = {}) {
    if (sites) return Auth.getSitesToken();
    if (write) return Auth.getWriteToken();
    return Auth.getReadToken();
  }

  async function callRaw(url, { method = 'GET', body, headers = {}, attempt = 0, write = false, sites = false } = {}) {
    const token = await _tokenFor({ write, sites });
    const opts = {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        ...(body ? { 'Content-Type': 'application/json' } : {}),
        ...headers,
      },
    };
    if (body) opts.body = typeof body === 'string' ? body : JSON.stringify(body);
    Log.dbg(`GRAPH ${method} ${url}`);
    const res = await fetch(url, opts);
    if (res.status === 429 || res.status === 503 || res.status === 504) {
      const ra = parseInt(res.headers.get('Retry-After') || '0', 10);
      const wait = (ra > 0 ? ra : Math.min(60, Math.pow(2, attempt))) * 1000;
      if (attempt >= 5) throw new Error(`Graph ${res.status} after retries: ${url}`);
      Log.warn(`Graph ${res.status}, backing off ${wait}ms (attempt ${attempt + 1})`);
      await Concurrency.sleep(wait);
      return callRaw(url, { method, body, headers, attempt: attempt + 1, write, sites });
    }
    return res;
  }

  async function call(path, opts = {}) {
    const url = path.startsWith('http') ? path : BASE + path;
    const res = await callRaw(url, opts);
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      const e = new Error(`Graph ${res.status} ${res.statusText}: ${txt.substring(0, 400)}`);
      e.status = res.status;
      throw e;
    }
    if (res.status === 204) return null;
    const ct = res.headers.get('Content-Type') || '';
    if (ct.includes('json')) return res.json();
    return res.text();
  }

  async function pageAll(path, opts = {}) {
    const out = [];
    let next = path.startsWith('http') ? path : BASE + path;
    while (next) {
      const data = await call(next, opts);
      if (data && Array.isArray(data.value)) out.push(...data.value);
      next = data && data['@odata.nextLink'];
    }
    return out;
  }

  // ---------- Service principals ----------

  /**
   * List all service principals in the tenant.
   * Uses ConsistencyLevel=eventual + $top=999 for max throughput.
   */
  async function listAllServicePrincipals(onProgress) {
    const select = [
      'id','appId','displayName','servicePrincipalType','accountEnabled','tags',
      'appOwnerOrganizationId','homepage','publisherName','signInAudience',
      'appDisplayName','createdDateTime','appRoles','oauth2PermissionScopes',
    ].join(',');
    const path = `/servicePrincipals?$top=999&$select=${select}`;
    const out = [];
    let pageNo = 0;
    let next = BASE + path;
    while (next) {
      const data = await call(next, { headers: { ConsistencyLevel: 'eventual' } });
      if (data && Array.isArray(data.value)) out.push(...data.value);
      pageNo++;
      if (onProgress) onProgress(out.length, pageNo);
      next = data && data['@odata.nextLink'];
    }
    return out;
  }

  /** Get a single service principal by object id (rich detail). */
  async function getServicePrincipal(spObjectId) {
    return call(`/servicePrincipals/${spObjectId}`);
  }

  /** Find the resource SPN for a given appId (e.g. Microsoft Graph). */
  async function getServicePrincipalByAppId(appId) {
    const data = await call(`/servicePrincipals?$filter=appId eq '${appId}'&$top=1`);
    return data && data.value && data.value[0] ? data.value[0] : null;
  }

  // ---------- Application permissions (appRoleAssignments) ----------

  /**
   * List the application permissions GRANTED TO the given SPN.
   * Each entry has resourceId (the API SPN) and appRoleId (the role guid).
   */
  function listAppRoleAssignments(spObjectId) {
    return pageAll(`/servicePrincipals/${spObjectId}/appRoleAssignments`);
  }

  /** Add an application permission to an SPN. */
  function addAppRoleAssignment(spObjectId, { resourceId, appRoleId }) {
    return call(`/servicePrincipals/${spObjectId}/appRoleAssignments`, {
      method: 'POST',
      body: { principalId: spObjectId, resourceId, appRoleId },
      write: true,
    });
  }

  /** Remove an application permission. */
  function removeAppRoleAssignment(spObjectId, assignmentId) {
    return call(`/servicePrincipals/${spObjectId}/appRoleAssignments/${assignmentId}`, {
      method: 'DELETE',
      write: true,
    });
  }

  // ---------- Delegated permissions (oauth2PermissionGrants) ----------

  /**
   * List delegated permission grants where this SPN is the CLIENT.
   * Each grant has resourceId (the API SPN), scope (space-separated scope names),
   * and consentType ('AllPrincipals' = tenant-wide, 'Principal' = per-user).
   */
  function listOAuth2Grants(spObjectId) {
    return pageAll(`/servicePrincipals/${spObjectId}/oauth2PermissionGrants`);
  }

  /** Add a delegated permission grant (or replace an existing one's scopes). */
  function createOAuth2Grant({ clientId, resourceId, scope, consentType = 'AllPrincipals', principalId = null }) {
    return call(`/oauth2PermissionGrants`, {
      method: 'POST',
      body: { clientId, consentType, principalId, resourceId, scope },
      write: true,
    });
  }
  function updateOAuth2Grant(grantId, { scope }) {
    return call(`/oauth2PermissionGrants/${grantId}`, {
      method: 'PATCH',
      body: { scope },
      write: true,
    });
  }
  function deleteOAuth2Grant(grantId) {
    return call(`/oauth2PermissionGrants/${grantId}`, { method: 'DELETE', write: true });
  }

  // ---------- SharePoint site permissions (Sites.Selected) ----------
  // All of these need the delegated Sites.FullControl.All scope (sites: true) and a
  // signed-in user who is SharePoint Administrator or higher.

  /**
   * Find sites in the tenant. `query` is a free-text search; pass '*' to match everything.
   * Returns the first page (search-as-you-type); large tenants would choke on a full crawl.
   */
  async function searchSites(query, top = 100) {
    const q = String(query == null ? '' : query).trim() || '*';
    const select = ['id', 'displayName', 'name', 'webUrl'].join(',');
    const path = `/sites?search=${encodeURIComponent(q)}&$top=${top}&$select=${select}`;
    const data = await call(path, { sites: true });
    return (data && Array.isArray(data.value)) ? data.value : [];
  }

  /** List the application permissions currently granted on a site. */
  function getSitePermissions(siteId) {
    return pageAll(`/sites/${siteId}/permissions`, { sites: true });
  }

  /** Grant an application (by appId / client id) access to a site with the given roles. */
  function addSitePermission(siteId, { appId, displayName, roles }) {
    return call(`/sites/${siteId}/permissions`, {
      method: 'POST',
      body: { roles, grantedToIdentities: [{ application: { id: appId, displayName } }] },
      sites: true,
    });
  }

  /** Change the roles of an existing site permission. */
  function updateSitePermission(siteId, permissionId, { roles }) {
    return call(`/sites/${siteId}/permissions/${permissionId}`, {
      method: 'PATCH',
      body: { roles },
      sites: true,
    });
  }

  /** Remove a site permission entirely. */
  function deleteSitePermission(siteId, permissionId) {
    return call(`/sites/${siteId}/permissions/${permissionId}`, { method: 'DELETE', sites: true });
  }

  window.Graph = {
    call, pageAll,
    listAllServicePrincipals, getServicePrincipal, getServicePrincipalByAppId,
    listAppRoleAssignments, addAppRoleAssignment, removeAppRoleAssignment,
    listOAuth2Grants, createOAuth2Grant, updateOAuth2Grant, deleteOAuth2Grant,
    searchSites, getSitePermissions, addSitePermission, updateSitePermission, deleteSitePermission,
  };
})();
