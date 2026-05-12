// MSAL.js PKCE auth wrapper
(() => {
  let CLIENT_ID = null;
  let REDIRECT_URI = null;
  let authConfig = null;

  let msal = null;
  let account = null;
  let tenantHost = null; // e.g. contoso.sharepoint.com
  let adminHost  = null; // e.g. contoso-admin.sharepoint.com
  const _tokenInflight = new Map();

  const SCOPES = {
    graph: ['https://graph.microsoft.com/Sites.Read.All',
            'https://graph.microsoft.com/Sites.FullControl.All',
            'https://graph.microsoft.com/User.Read'],
    sharepoint: () => [`https://${tenantHost}/.default`],
    sharepointAdmin: () => [`https://${adminHost}/.default`],
  };

  /**
   * Loads the .auth file (sibling of index.html) and parses key:value lines.
   * Required keys: clientId, redirect-url-local, redirect-url-web.
   */
  async function loadAuthConfig() {
    const res = await fetch('./.auth', { cache: 'no-store' });
    if (!res.ok) throw new Error(`Could not load .auth file (HTTP ${res.status}). Make sure the tool is being served over HTTP and the .auth file exists in the tool folder.`);
    const text = await res.text();
    const cfg = {};
    for (const raw of text.split(/\r?\n/)) {
      const line = raw.trim();
      if (!line || line.startsWith('#')) continue;
      const idx = line.indexOf(':');
      if (idx < 0) continue;
      const key = line.substring(0, idx).trim();
      const val = line.substring(idx + 1).trim();
      cfg[key] = val;
    }
    if (!cfg.clientId) throw new Error('.auth is missing clientId');
    if (!cfg['redirect-url-local'] || !cfg['redirect-url-web']) {
      throw new Error('.auth is missing redirect-url-local or redirect-url-web');
    }
    return cfg;
  }

  /**
   * Picks the correct redirect URI based on the current origin.
   * - localhost / 127.0.0.1 / file → redirect-url-local
   * - any other origin → redirect-url-web
   * The chosen URI must match what is registered on the app registration.
   */
  function pickRedirectUri(cfg) {
    const host = window.location.hostname;
    const isLocal = host === 'localhost' || host === '127.0.0.1' || host === '::1' || host === '';
    const chosen = isLocal ? cfg['redirect-url-local'] : cfg['redirect-url-web'];
    Log.info(`Auth environment detected: ${isLocal ? 'local' : 'production'} → redirect_uri=${chosen}`);
    // Sanity-check that the chosen URI's origin matches the page origin.
    try {
      const u = new URL(chosen);
      if (u.origin !== window.location.origin) {
        Log.warn(`Configured redirect URI origin (${u.origin}) does not match page origin (${window.location.origin}). MSAL will reject this. Update .auth or serve from the matching origin.`);
      }
    } catch (e) {
      Log.warn(`Configured redirect URI is not a valid URL: ${chosen}`);
    }
    return chosen;
  }

  async function init() {
    authConfig = await loadAuthConfig();
    CLIENT_ID = authConfig.clientId;
    REDIRECT_URI = pickRedirectUri(authConfig);

    if (window.__msalLoadFailed || !window.msal || !window.msal.PublicClientApplication) {
      throw new Error('MSAL.js failed to load from CDN. Check your network connection or CSP - the SPA needs https://cdn.jsdelivr.net to be reachable.');
    }

    msal = new window.msal.PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/organizations', redirectUri: REDIRECT_URI },
      cache: { cacheLocation: 'localStorage' },
    });
    await msal.initialize();
    const resp = await msal.handleRedirectPromise().catch(e => { Log.err('Redirect handler:', e); return null; });
    if (resp && resp.account) {
      account = resp.account;
    } else {
      const accs = msal.getAllAccounts();
      if (accs.length) account = accs[0];
    }
    if (account) {
      await detectTenantHosts();
    }
    return account;
  }

  async function detectTenantHosts() {
    try {
      const tok = await getToken(SCOPES.graph);
      const orgRes = await fetch('https://graph.microsoft.com/v1.0/organization', {
        headers: { Authorization: `Bearer ${tok}` },
      });
      if (!orgRes.ok) throw new Error(`organization endpoint ${orgRes.status}`);
      const org = await orgRes.json();
      const verified = org.value?.[0]?.verifiedDomains || [];
      let initial = verified.find(d => d.isInitial)?.name || verified[0]?.name || '';
      const tenantPrefix = initial.split('.')[0];
      tenantHost = `${tenantPrefix}.sharepoint.com`;
      adminHost  = `${tenantPrefix}-admin.sharepoint.com`;
      Log.info(`Detected tenant: ${initial} (SPO host: ${tenantHost})`);
    } catch (e) {
      Log.warn('Failed to detect tenant host:', e.message);
    }
  }

  async function signIn() {
    const req = { scopes: SCOPES.graph, prompt: 'select_account' };
    try {
      const r = await msal.loginPopup(req);
      account = r.account;
    } catch (e) {
      Log.warn('Popup sign-in failed, falling back to redirect:', e.message);
      await msal.loginRedirect(req);
      return null;
    }
    await detectTenantHosts();
    return account;
  }

  async function signOut() {
    if (!account) return;
    try { await msal.logoutPopup({ account }); } catch (e) { Log.warn('Logout popup failed:', e.message); }
    account = null;
  }

  async function getToken(scopes) {
    if (!account) throw new Error('Not signed in');
    const key = scopes.join('|');
    if (_tokenInflight.has(key)) return _tokenInflight.get(key);
    const p = (async () => {
      try {
        const r = await msal.acquireTokenSilent({ scopes, account });
        return r.accessToken;
      } catch (e) {
        Log.warn(`Silent token acquisition failed for [${scopes.join(',')}], using popup`);
        const r = await msal.acquireTokenPopup({ scopes, account });
        return r.accessToken;
      } finally {
        // Brief debounce: keep the inflight entry only until first await chain settles
      }
    })();
    _tokenInflight.set(key, p);
    try { return await p; }
    finally { _tokenInflight.delete(key); }
  }

  function getGraphToken()           { return getToken(SCOPES.graph); }
  function getSpoToken()             { return getToken(SCOPES.sharepoint()); }
  function getSpoAdminToken()        { return getToken(SCOPES.sharepointAdmin()); }
  function getAccount()              { return account; }
  function getTenantHost()           { return tenantHost; }
  function getAdminHost()            { return adminHost; }
  function getAuthConfig()           { return authConfig; }
  function getRedirectUri()          { return REDIRECT_URI; }

  // Exported only for tests
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
  function _pickRedirectUriForOrigin(cfg, origin) {
    const host = new URL(origin).hostname;
    const isLocal = host === 'localhost' || host === '127.0.0.1' || host === '::1';
    return isLocal ? cfg['redirect-url-local'] : cfg['redirect-url-web'];
  }

  window.Auth = {
    init, signIn, signOut,
    getGraphToken, getSpoToken, getSpoAdminToken,
    getAccount, getTenantHost, getAdminHost, getAuthConfig, getRedirectUri,
    _parseAuthFile, _pickRedirectUriForOrigin,
  };
})();
