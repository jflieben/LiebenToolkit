// MSAL.js PKCE auth wrapper supporting TWO independent tenant connections
// (a "source" connection and a "target" connection). Each connection tracks
// its own signed-in account and its own SharePoint tenant host so we can move
// files within one tenant or between two different tenants.
//
// Sign-in uses the popup flow so the wizard state is never lost when a second
// tenant is authenticated. A single MSAL PublicClientApplication (authority
// 'organizations') holds all accounts; per-connection tokens are acquired for
// the tenant-specific SharePoint resource (https://{host}/.default).
(() => {
  const GRAPH_SCOPES = ['https://graph.microsoft.com/User.Read'];

  let CLIENT_ID = null;
  let REDIRECT_URI = null;
  let authConfig = null;
  let msal = null;

  // A connection = one signed-in account + its resolved SharePoint hosts.
  // roles: 'source' and 'target'. Same-tenant mode reuses the source connection
  // for both ends.
  const connections = {
    source: emptyConnection('source'),
    target: emptyConnection('target'),
  };
  const _tokenInflight = new Map();

  function emptyConnection(role) {
    return { role, account: null, tenantHost: null, adminHost: null, tenantId: null, tenantName: null, upn: null, spoConsent: false };
  }

  // ---------- .auth loading ----------
  async function loadAuthConfig() {
    const res = await fetch('./.auth', { cache: 'no-store' });
    if (!res.ok) throw new Error(`Could not load .auth file (HTTP ${res.status}). Serve the tool over HTTP and make sure .auth exists in the tool folder.`);
    const text = await res.text();
    const cfg = _parseAuthFile(text);
    if (!cfg.clientId) throw new Error('.auth is missing clientId');
    if (!cfg['redirect-url-local'] || !cfg['redirect-url-web']) {
      throw new Error('.auth is missing redirect-url-local or redirect-url-web');
    }
    return cfg;
  }

  function _parseAuthFile(text) {
    const cfg = {};
    for (const raw of text.split(/\r?\n/)) {
      const line = raw.trim();
      if (!line || line.startsWith('#')) continue;
      const idx = line.indexOf(':');
      if (idx < 0) continue;
      cfg[line.substring(0, idx).trim()] = line.substring(idx + 1).trim();
    }
    return cfg;
  }

  function pickRedirectUri(cfg) {
    const host = window.location.hostname;
    const isLocal = host === 'localhost' || host === '127.0.0.1' || host === '::1' || host === '';
    const chosen = isLocal ? cfg['redirect-url-local'] : cfg['redirect-url-web'];
    Log.info(`Auth environment: ${isLocal ? 'local' : 'production'} → redirect_uri=${chosen}`);
    try {
      const u = new URL(chosen);
      if (u.origin !== window.location.origin) {
        Log.warn(`Configured redirect origin (${u.origin}) != page origin (${window.location.origin}). MSAL will reject this.`);
      }
    } catch { Log.warn(`Configured redirect URI is not a valid URL: ${chosen}`); }
    return chosen;
  }

  // Detect the popup callback: MSAL's popup client reads the response from the
  // opened window, so when index.html loads inside that popup we must NOT boot
  // the full app.
  function isAuthPopup() {
    if (!window.opener || window.opener === window) return false;
    const bits = `${window.location.hash || ''}&${window.location.search || ''}`;
    return /(?:^|[?&#])(code|error|state|id_token)=/i.test(bits);
  }

  async function init() {
    authConfig = await loadAuthConfig();
    CLIENT_ID = authConfig.clientId;
    REDIRECT_URI = pickRedirectUri(authConfig);

    if (window.__msalLoadFailed || !window.msal || !window.msal.PublicClientApplication) {
      throw new Error('MSAL.js failed to load. Check that ../vendor/msal-browser.3.27.0.min.js is present and allowed by CSP.');
    }

    msal = new window.msal.PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/organizations', redirectUri: REDIRECT_URI },
      cache: { cacheLocation: 'localStorage', temporaryCacheLocation: 'sessionStorage' },
      system: { allowNativeBroker: false },
    });
    await msal.initialize();
    // Drain any stray redirect response (we use popups, but be safe).
    await msal.handleRedirectPromise().catch(e => { Log.dbg('handleRedirectPromise:', e.message); return null; });
    return true;
  }

  // ---------- Sign-in / out per connection ----------
  // opts.loginHint pre-fills the account; opts.expectTenantId / expectTenantHost pin the
  // sign-in to a specific tenant (used when reconnecting for an existing job) and reject a
  // wrong pick so files can't be transferred to/from the wrong tenant.
  async function signIn(role, { loginHint, expectTenantId, expectTenantHost } = {}) {
    const conn = connections[role];
    if (!conn) throw new Error(`Unknown connection role: ${role}`);
    if (!msal) throw new Error('Auth not initialised');
    Log.info(`Opening sign-in popup for the ${role} tenant...`);
    const req = { scopes: GRAPH_SCOPES };
    if (loginHint) req.loginHint = loginHint; else req.prompt = 'select_account';
    let resp;
    try {
      resp = await msal.loginPopup(req);
    } catch (e) {
      if (/popup_window_error|popup.*block|BrowserAuthError/i.test(e.message || '')) {
        throw new Error('The sign-in popup was blocked. Allow popups for this site and try again.');
      }
      throw e;
    }
    conn.account = resp.account;
    conn.upn = resp.account?.username || null;
    await detectTenant(conn);
    if ((expectTenantHost && conn.tenantHost && conn.tenantHost !== expectTenantHost) ||
        (expectTenantId && conn.tenantId && conn.tenantId !== expectTenantId)) {
      const got = conn.tenantName || conn.tenantHost || conn.upn;
      connections[role] = emptyConnection(role);
      throw new Error(`Signed in to ${got}, but this job's ${role} tenant is ${expectTenantHost || expectTenantId}. Use an account in the correct tenant.`);
    }
    Log.info(`${role} tenant connected: ${conn.tenantName || conn.tenantHost} (${conn.upn})`);
    // Detect (silently, no popup) whether the SharePoint API is consented for this tenant, so
    // the wizard can prompt for it right after login — before the user tries to browse files.
    conn.spoConsent = await hasSharePointConsent(role);
    if (!conn.spoConsent) Log.warn(`SharePoint access not yet approved for the ${role} tenant — the wizard will prompt to grant it.`);
    return conn;
  }

  async function signOut(role) {
    const conn = connections[role];
    if (!conn || !conn.account) return;
    // Only clear our local reference; do not logoutRedirect (would blow away the
    // other connection and reload the page). Removing the account from the cache
    // is enough for our purposes.
    try { await msal.clearCache({ account: conn.account }); } catch {}
    connections[role] = emptyConnection(role);
  }

  async function signOutAll() {
    try { await msal.logoutPopup(); } catch (e) { Log.dbg('logoutPopup:', e.message); }
    connections.source = emptyConnection('source');
    connections.target = emptyConnection('target');
  }

  // Resolve tenant id, primary domain and SharePoint hosts for a connection.
  async function detectTenant(conn, { silent = false } = {}) {
    try {
      const tok = await getToken(conn.role, GRAPH_SCOPES, { silent });
      const orgRes = await fetch('https://graph.microsoft.com/v1.0/organization?$select=id,displayName,verifiedDomains', {
        headers: { Authorization: `Bearer ${tok}` },
      });
      if (!orgRes.ok) throw new Error(`organization endpoint ${orgRes.status}`);
      const org = await orgRes.json();
      const o = org.value?.[0] || {};
      conn.tenantId = o.id || conn.account?.tenantId || null;
      conn.tenantName = o.displayName || null;
      const verified = o.verifiedDomains || [];
      const initial = verified.find(d => d.isInitial)?.name || verified[0]?.name || '';
      const prefix = initial.split('.')[0];
      if (prefix) {
        conn.tenantHost = `${prefix}.sharepoint.com`;
        conn.adminHost = `${prefix}-admin.sharepoint.com`;
      }
      Log.info(`Detected ${conn.role} SPO host: ${conn.tenantHost}`);
    } catch (e) {
      Log.warn(`Failed to detect ${conn.role} tenant host:`, e.message);
      throw new Error(`Could not resolve the SharePoint host for the ${conn.role} tenant: ${e.message}`);
    }
  }

  // Re-attach a connection from MSAL's cached accounts WITHOUT any popup. Used when a job is
  // resumed/re-run after a reload: the account is usually still in the localStorage cache even
  // though our in-memory connection was reset. Returns true if a matching account was restored.
  async function restoreConnection(role, loc) {
    if (!msal) return false;
    const accs = msal.getAllAccounts();
    if (!accs.length) return false;
    let candidates = [];
    if (loc?.tenantId) candidates = accs.filter(a => a.tenantId === loc.tenantId);
    if (!candidates.length && loc?.upn) candidates = accs.filter(a => (a.username || '').toLowerCase() === loc.upn.toLowerCase());
    if (!candidates.length) candidates = accs;
    for (const acc of candidates) {
      const conn = emptyConnection(role);
      conn.account = acc;
      conn.upn = acc.username || null;
      connections[role] = conn; // so getToken can see the account during silent detection
      try {
        await detectTenant(conn, { silent: true });
      } catch (e) {
        Log.dbg(`Silent restore of ${role} via ${acc.username} failed: ${e.message}`);
        connections[role] = emptyConnection(role);
        continue;
      }
      if ((loc?.tenantHost && conn.tenantHost && conn.tenantHost !== loc.tenantHost) ||
          (loc?.tenantId && conn.tenantId && conn.tenantId !== loc.tenantId)) {
        connections[role] = emptyConnection(role);
        continue;
      }
      conn.spoConsent = await hasSharePointConsent(role);
      Log.info(`Reconnected ${role} tenant from cache: ${conn.tenantName || conn.tenantHost} (${conn.upn})`);
      return true;
    }
    connections[role] = emptyConnection(role);
    return false;
  }

  // Is the current connection already for the tenant a job needs?
  function connectionMatches(role, loc) {
    const c = connections[role];
    if (!c?.account) return false;
    if (loc?.tenantId && c.tenantId) return c.tenantId === loc.tenantId;
    if (loc?.tenantHost && c.tenantHost) return c.tenantHost === loc.tenantHost;
    return true;
  }

  // ---------- Token acquisition ----------
  function _isCacheQuota(e) {
    return !!e && (e.errorCode === 'cache_quota_exceeded' || e.name === 'CacheError' || /cache_quota_exceeded|Exceeded cache storage/i.test(e.message || ''));
  }
  // MSAL persists tokens in localStorage; when the origin's storage is full it throws a
  // CacheError and the token is lost. Access tokens are re-acquirable from the (retained)
  // refresh tokens, so dropping them frees the most space without signing the user out.
  function _freeTokenCacheSpace() {
    let removed = 0;
    try {
      for (let i = localStorage.length - 1; i >= 0; i--) {
        const k = localStorage.key(i);
        const v = k && localStorage.getItem(k);
        if (!v || v[0] !== '{') continue;
        try { if (JSON.parse(v).credentialType === 'AccessToken') { localStorage.removeItem(k); removed++; } } catch {}
      }
    } catch {}
    return removed;
  }

  async function getToken(role, scopes, { silent = false } = {}) {
    const conn = connections[role];
    if (!conn || !conn.account) throw new Error(`Not signed in for the ${role} tenant`);
    const key = `${role}|${silent ? 's' : 'i'}|${scopes.join('|')}`;
    if (_tokenInflight.has(key)) return _tokenInflight.get(key);
    const acquire = async (mode) => {
      const req = { scopes, account: conn.account };
      const once = () => (mode === 'popup' ? msal.acquireTokenPopup(req) : msal.acquireTokenSilent(req)).then(r => r.accessToken);
      try {
        return await once();
      } catch (e) {
        if (!_isCacheQuota(e)) throw e;
        const n = _freeTokenCacheSpace();
        Log.warn(`Token cache storage was full; cleared ${n} cached access token(s) and retrying.`);
        try {
          return await once();
        } catch (e2) {
          if (_isCacheQuota(e2)) throw new Error('Your browser storage for this site is full. Clear this site’s data in the browser and reconnect.');
          throw e2;
        }
      }
    };
    const p = (async () => {
      try {
        return await acquire('silent');
      } catch (e) {
        if (silent) throw e;
        if (/storage for this site is full/i.test(e.message || '')) throw e;
        Log.warn(`Silent token failed for ${role} [${scopes.join(',')}], using popup`);
        return await acquire('popup');
      }
    })();
    _tokenInflight.set(key, p);
    try { return await p; }
    finally { _tokenInflight.delete(key); }
  }

  function getGraphToken(role)      { return getToken(role, GRAPH_SCOPES); }
  // Incremental consent: acquire a Graph token with extra scopes (e.g.
  // User.ReadBasic.All for OneDrive-by-user lookup) only when the feature is used.
  function getGraphTokenScoped(role, scopes) {
    const wanted = Array.from(new Set([...GRAPH_SCOPES, ...(scopes || [])]));
    return getToken(role, wanted);
  }
  function getSpoToken(role, host)  {
    const conn = connections[role];
    const h = host || conn?.tenantHost;
    if (!h) throw new Error(`SharePoint host unknown for the ${role} tenant`);
    return getToken(role, [`https://${h}/.default`]);
  }
  function getSpoAdminToken(role)   {
    const conn = connections[role];
    if (!conn?.adminHost) throw new Error(`SharePoint admin host unknown for the ${role} tenant`);
    return getToken(role, [`https://${conn.adminHost}/.default`]);
  }

  // Silent check (no popup) of whether the SharePoint API is already consented for a tenant.
  async function hasSharePointConsent(role) {
    const conn = connections[role];
    if (!conn?.account || !conn?.tenantHost) return false;
    try {
      await msal.acquireTokenSilent({ scopes: [`https://${conn.tenantHost}/.default`], account: conn.account });
      return true;
    } catch { return false; }
  }

  // Consent to SharePoint for a tenant. All SharePoint hosts (sites, -my OneDrive, -admin) are
  // the SAME first-party resource (SharePoint Online), so a single .default consent covers them
  // all. .default scopes for DIFFERENT resource URLs must NOT be combined in one request — doing
  // so fails with AADSTS70011 "static scope limit exceeded". We never force prompt:'consent'
  // (that re-shows the approval screen every call); a silent-first check + an in-flight guard
  // make repeat calls no-ops once consent exists.
  async function grantSharePointConsent(role) {
    const conn = connections[role];
    if (!conn?.account) throw new Error(`Not signed in for the ${role} tenant`);
    if (!conn.tenantHost) throw new Error(`SharePoint host unknown for the ${role} tenant`);
    if (conn._spoConsentInflight) return conn._spoConsentInflight;
    conn._spoConsentInflight = (async () => {
      const primary = `https://${conn.tenantHost}/.default`;
      // Already consented? Acquire silently and stop — no popup, no repeat prompt.
      try {
        await msal.acquireTokenSilent({ scopes: [primary], account: conn.account });
        conn.spoConsent = true;
        return;
      } catch { /* need interactive consent */ }
      Log.info(`Requesting SharePoint consent for the ${role} tenant...`);
      await msal.acquireTokenPopup({ scopes: [primary], account: conn.account });
      conn.spoConsent = true;
      // Pre-warm the OneDrive (-my) and admin audiences silently so later use needs no extra
      // popup; failures are fine — getToken falls back to a single-scope popup on demand.
      for (const h of [conn.tenantHost.replace(/\.sharepoint\.com$/i, '-my.sharepoint.com'), conn.adminHost]) {
        if (!h) continue;
        try { await msal.acquireTokenSilent({ scopes: [`https://${h}/.default`], account: conn.account }); }
        catch (e) { Log.dbg(`Could not pre-warm ${h} token: ${e.message}`); }
      }
    })();
    try { return await conn._spoConsentInflight; }
    finally { conn._spoConsentInflight = null; }
  }

  function getConnection(role)  { return connections[role]; }
  function isConnected(role)    { return !!connections[role]?.account; }
  function getRedirectUri()     { return REDIRECT_URI; }

  window.Auth = {
    init, signIn, signOut, signOutAll, isAuthPopup, restoreConnection, connectionMatches,
    getGraphToken, getGraphTokenScoped, getSpoToken, getSpoAdminToken, grantSharePointConsent, hasSharePointConsent,
    getConnection, isConnected, getRedirectUri,
    _parseAuthFile,
  };
})();
