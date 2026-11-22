// Microsoft Graph calls per connection: site search, user search for the OneDrive
// picker and a site list fallback when the SharePoint admin API is not available.
(() => {
  const BASE = 'https://graph.microsoft.com/v1.0';
  const USER_SCOPES = ['https://graph.microsoft.com/User.ReadBasic.All'];

  function opts(role, scopes) {
    return {
      label: `Graph[${role}]`,
      token: () => (scopes ? Auth.getGraphTokenScoped(role, scopes) : Auth.getGraphToken(role)),
    };
  }

  function call(role, path, { scopes, ...rest } = {}) {
    if (/^https?:/i.test(path) && !path.startsWith(`${BASE}/`)) throw new Error(`Refusing a Graph call to ${path}`);
    const url = path.startsWith('http') ? path : BASE + path;
    return Http.json(url, { ...rest, ...opts(role, scopes) });
  }

  // Graph /sites?search matches per term on a prefix; a trailing * per term finds partial words.
  async function searchSites(role, query) {
    const q = String(query || '').trim();
    const kql = q ? q.split(/\s+/).filter(Boolean).map((t) => (t.endsWith('*') ? t : `${t}*`)).join(' ') : '*';
    const data = await call(role, `/sites?search=${encodeURIComponent(kql)}&$select=id,name,displayName,webUrl&$top=50`);
    return ((data && data.value) || []).map((s) => ({ id: s.id, name: s.name, displayName: s.displayName || s.name || s.webUrl, webUrl: s.webUrl }));
  }

  async function searchUsers(role, query) {
    const q = String(query || '').trim();
    if (!q) return [];
    const lit = Http.odataString(q);
    const filter = ['displayName', 'userPrincipalName', 'mail', 'givenName', 'surname'].map((f) => `startswith(${f},${lit})`).join(' or ');
    const select = '$select=id,displayName,userPrincipalName,mail&$top=25';
    let data;
    try {
      data = await call(role, `/users?$filter=${encodeURIComponent(filter)}&${select}&$orderby=displayName`, { scopes: USER_SCOPES });
    } catch (e) {
      // $orderby with this $filter needs ConsistencyLevel on some tenants.
      if (e.status !== 400) throw e;
      data = await call(role, `/users?$filter=${encodeURIComponent(filter)}&${select}`, { scopes: USER_SCOPES });
    }
    return ((data && data.value) || []).map((u) => ({ id: u.id, displayName: u.displayName || u.userPrincipalName, userPrincipalName: u.userPrincipalName, mail: u.mail }));
  }

  // The drive webUrl is authoritative, but reading another user's /drive needs Files.Read.All
  // or Sites.Read.All, which a tenant often has not consented. Fallback: build the personal
  // site URL from the UPN and the tenant's -my host.
  async function getUserOneDriveSiteUrl(role, user) {
    const userId = user && typeof user === 'object' ? user.id : user;
    const upn = user && typeof user === 'object' ? (user.userPrincipalName || user.mail) : null;
    try {
      const data = await call(role, Http.url`/users/${userId}/drive?$select=webUrl`, { scopes: USER_SCOPES });
      const site = oneDriveSiteUrlFromWebUrl(data && data.webUrl);
      if (site) return site;
    } catch (e) {
      if (e.status !== 404 && e.status !== 403) throw e;
      Log.dbg(`Graph /drive unavailable (${e.status}); building the OneDrive URL from the UPN.`);
    }
    const built = buildOneDriveSiteUrl(role, upn);
    if (built) return built;
    throw new Error('Could not resolve this user\'s OneDrive. Paste its URL instead.');
  }

  function buildOneDriveSiteUrl(role, upn) {
    const conn = Auth.getConnection(role);
    const host = conn && conn.tenantHost;
    if (!upn || !host) return null;
    const myHost = host.replace(/\.sharepoint\.com$/i, '-my.sharepoint.com');
    return `https://${myHost}/personal/${String(upn).toLowerCase().replace(/[^a-z0-9]/g, '_')}`;
  }

  // https://t-my.sharepoint.com/personal/user_domain_com/Documents -> .../personal/user_domain_com
  function oneDriveSiteUrlFromWebUrl(webUrl) {
    if (!webUrl) return null;
    try {
      const u = new URL(webUrl);
      const m = u.pathname.match(/\/personal\/[^/]+/i);
      if (m) return u.origin + m[0];
      return webUrl.replace(/\/Documents\/?$/i, '');
    } catch {
      return String(webUrl).replace(/\/Documents\/?$/i, '');
    }
  }

  // Sites visible to the signed-in account (not every site in the tenant).
  async function listAllSites(role, onProgress) {
    const rows = await Http.pageAll(`${BASE}/sites?search=*&$select=id,name,displayName,webUrl&$top=100`, opts(role),
      (items, page, total) => onProgress && onProgress(total));
    const seen = new Set();
    const out = [];
    for (const s of rows) {
      if (!s.webUrl) continue;
      const k = s.webUrl.toLowerCase();
      if (seen.has(k)) continue;
      seen.add(k);
      out.push({ url: s.webUrl, title: s.displayName || s.name || s.webUrl });
    }
    out.sort((a, b) => String(a.title).localeCompare(String(b.title)));
    return out;
  }

  window.Graph = { searchSites, searchUsers, getUserOneDriveSiteUrl, oneDriveSiteUrlFromWebUrl, listAllSites };
})();
