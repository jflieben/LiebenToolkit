// Microsoft Graph client (per-connection). Used for site search by name so the
// wizard can offer a picker instead of forcing the user to paste a URL.
(() => {
  const BASE = 'https://graph.microsoft.com/v1.0';
  const USER_SCOPES = ['https://graph.microsoft.com/User.ReadBasic.All'];

  async function callRaw(role, url, { method = 'GET', body, headers = {}, attempt = 0, scopes } = {}) {
    const token = scopes && scopes.length ? await Auth.getGraphTokenScoped(role, scopes) : await Auth.getGraphToken(role);
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
    Log.dbg(`GRAPH[${role}] ${method} ${url}`);
    let res;
    try {
      res = await fetch(url, opts);
    } catch (e) {
      // A dropped connection ("Failed to fetch") is usually throttling/overload; retry.
      if (attempt >= 5) throw new Error(`Graph throttled (connection dropped) after retries: ${e.message}`);
      const wait = Math.min(20, Math.pow(2, attempt)) * 1000;
      Log.warn(`Graph connection dropped (throttling) → retry in ${Math.round(wait / 1000)}s (attempt ${attempt + 1})`);
      await Concurrency.sleep(wait);
      return callRaw(role, url, { method, body, headers, attempt: attempt + 1, scopes });
    }
    if (res.status === 429 || res.status === 503 || res.status === 504) {
      const ra = parseInt(res.headers.get('Retry-After') || '0', 10);
      const wait = (ra > 0 ? ra : Math.min(60, Math.pow(2, attempt))) * 1000;
      if (attempt >= 5) throw new Error(`Graph ${res.status} after retries: ${url}`);
      Log.warn(`Graph ${res.status} → backing off ${wait}ms (attempt ${attempt + 1})`);
      await Concurrency.sleep(wait);
      return callRaw(role, url, { method, body, headers, attempt: attempt + 1, scopes });
    }
    return res;
  }

  async function call(role, path, opts = {}) {
    const url = path.startsWith('http') ? path : BASE + path;
    const res = await callRaw(role, url, opts);
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      const e = new Error(`Graph ${res.status} ${res.statusText}: ${txt.substring(0, 300)}`);
      e.status = res.status;
      throw e;
    }
    if (res.status === 204) return null;
    return res.json();
  }

  // Search sites by free text. Returns [{ id, name, displayName, webUrl }].
  async function searchSites(role, query) {
    const q = (query || '').trim();
    // Graph /sites?search does prefix matching per term, not substring. Append a
    // trailing wildcard to each term so partial words match (e.g. "rag" -> "RAGdemo").
    const kql = q
      ? q.split(/\s+/).filter(Boolean).map(t => (t.endsWith('*') ? t : `${t}*`)).join(' ')
      : '*';
    const data = await call(role, `/sites?search=${encodeURIComponent(kql)}&$select=id,name,displayName,webUrl&$top=50`);
    return (data.value || []).map(s => ({
      id: s.id, name: s.name, displayName: s.displayName || s.name || s.webUrl, webUrl: s.webUrl,
    }));
  }

  // Search users by name / UPN / mail (prefix match). Returns
  // [{ id, displayName, userPrincipalName, mail }]. Used to locate OneDrive sites.
  async function searchUsers(role, query) {
    const q = (query || '').trim();
    if (!q) return [];
    const lit = q.replace(/'/g, "''");
    const filter = `startswith(displayName,'${lit}') or startswith(userPrincipalName,'${lit}') or startswith(mail,'${lit}') or startswith(givenName,'${lit}') or startswith(surname,'${lit}')`;
    const url = `/users?$filter=${encodeURIComponent(filter)}&$select=id,displayName,userPrincipalName,mail&$top=25&$orderby=displayName`;
    let data;
    try {
      data = await call(role, url, { scopes: USER_SCOPES });
    } catch (e) {
      // $orderby + $filter can require ConsistencyLevel on some tenants; retry without orderby.
      if (e.status === 400) {
        data = await call(role, `/users?$filter=${encodeURIComponent(filter)}&$select=id,displayName,userPrincipalName,mail&$top=25`, { scopes: USER_SCOPES });
      } else { throw e; }
    }
    return (data.value || []).map(u => ({
      id: u.id, displayName: u.displayName || u.userPrincipalName, userPrincipalName: u.userPrincipalName, mail: u.mail,
    }));
  }

  // Resolve a user's OneDrive personal SharePoint site URL. Prefer the authoritative
  // Graph drive webUrl, but reading another user's /drive needs Files.Read.All /
  // Sites.Read.All — which a target tenant often hasn't consented (Graph then 404s).
  // Fall back to constructing the personal site URL from the UPN + tenant -my host,
  // which needs no extra Graph permission and works cross-tenant.
  async function getUserOneDriveSiteUrl(role, user) {
    const userId = (user && typeof user === 'object') ? user.id : user;
    const upn = (user && typeof user === 'object') ? (user.userPrincipalName || user.mail) : null;
    try {
      const data = await call(role, `/users/${encodeURIComponent(userId)}/drive?$select=webUrl`, { scopes: USER_SCOPES });
      const site = oneDriveSiteUrlFromWebUrl(data?.webUrl);
      if (site) return site;
    } catch (e) {
      if (e.status !== 404 && e.status !== 403) throw e;
      Log.dbg(`Graph /drive unavailable (${e.status}); building OneDrive URL from UPN.`);
    }
    const built = buildOneDriveSiteUrl(role, upn);
    if (built) return built;
    throw new Error('Could not resolve this user’s OneDrive. Paste its URL directly instead.');
  }

  // Construct a user's personal site URL from their UPN and the tenant's -my host,
  // e.g. john.doe@contoso.com on contoso.sharepoint.com ->
  // https://contoso-my.sharepoint.com/personal/john_doe_contoso_com.
  function buildOneDriveSiteUrl(role, upn) {
    const host = Auth.getConnection(role)?.tenantHost;
    if (!upn || !host) return null;
    const myHost = host.replace(/\.sharepoint\.com$/i, '-my.sharepoint.com');
    const seg = String(upn).toLowerCase().replace(/[^a-z0-9]/g, '_');
    return `https://${myHost}/personal/${seg}`;
  }

  // Derive the personal site collection URL from a OneDrive drive webUrl, e.g.
  // https://t-my.sharepoint.com/personal/user_domain_com/Documents ->
  // https://t-my.sharepoint.com/personal/user_domain_com
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

  // Fallback site enumeration via Graph when the SharePoint tenant admin API is unavailable
  // (it needs the SharePoint Administrator role + app consent). Returns sites VISIBLE to the
  // signed-in account: [{ url, title }]. Not every tenant site — just what search can see.
  async function listAllSites(role, onProgress) {
    const out = [];
    const seen = new Set();
    let url = `/sites?search=*&$select=id,name,displayName,webUrl&$top=100`;
    let guard = 0;
    while (url && guard++ < 500) {
      const data = await call(role, url);
      for (const s of (data.value || [])) {
        const w = s.webUrl; if (!w) continue;
        const k = w.toLowerCase(); if (seen.has(k)) continue; seen.add(k);
        out.push({ url: w, title: s.displayName || s.name || w });
      }
      if (onProgress) onProgress(out.length);
      url = data['@odata.nextLink'] || null;
    }
    out.sort((a, b) => String(a.title).localeCompare(String(b.title)));
    return out;
  }

  window.Graph = { searchSites, searchUsers, getUserOneDriveSiteUrl, oneDriveSiteUrlFromWebUrl, listAllSites };
})();
