// MSAL.js PKCE auth wrapper for DupedDevices.
// Only Microsoft Graph is needed - no SharePoint scopes.
(() => {
  let CLIENT_ID = null;
  let REDIRECT_URI = null;
  let authConfig = null;
  let msal = null;
  let account = null;
  const _tokenInflight = new Map();
  const MSAL_VENDOR_FILE = 'msal-browser.3.27.0.min.js';

  // Read scopes:  Device.Read.All  - list all devices
  //               Directory.Read.All - resolve owners / extension attribute schema
  // Write scopes: Directory.AccessAsUser.All - covers PATCH/DELETE on /devices including extensionAttributes
  //               (Device.ReadWrite.All works for accountEnabled+delete but does NOT cover extensionAttributes)
  const SCOPES = {
    read: [
      'https://graph.microsoft.com/Device.Read.All',
      'https://graph.microsoft.com/Directory.Read.All',
      'https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All',
      'https://graph.microsoft.com/User.Read',
    ],
    write: [
      'https://graph.microsoft.com/Directory.AccessAsUser.All',
    ],
  };

  function _shouldClosePopupAfterAuth(resp) {
    if (!window.opener || window.opener.closed) return false;
    if (resp && resp.account) return true;
    const authBits = `${window.location.hash || ''}&${window.location.search || ''}`;
    return /(?:^|[?&#])(code|error|id_token|access_token|state)=/i.test(authBits);
  }

  function _loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.async = false;
      s.onload = () => resolve(true);
      s.onerror = () => reject(new Error(`Failed to load script: ${src}`));
      document.head.appendChild(s);
    });
  }

  async function _ensureMsalLoaded() {
    if (window.msal && window.msal.PublicClientApplication) return true;
    const candidates = [
      `../vendor/${MSAL_VENDOR_FILE}`,
      `./vendor/${MSAL_VENDOR_FILE}`,
      `/tools/vendor/${MSAL_VENDOR_FILE}`,
    ];
    for (const src of candidates) {
      try {
        await _loadScript(src);
        if (window.msal && window.msal.PublicClientApplication) {
          window.__msalLoadFailed = false;
          return true;
        }
      } catch (e) {
        Log.warn(`MSAL probe failed for ${src}: ${e.message}`);
      }
    }
    return false;
  }

  async function loadAuthConfig() {
    const res = await fetch('./.auth', { cache: 'no-store' });
    if (!res.ok) throw new Error(`Could not load .auth file (HTTP ${res.status}). Make sure the tool is being served over HTTP and the .auth file exists in the tool folder.`);
    const text = await res.text();
    const cfg = _parseAuthFile(text);
    if (!cfg.clientId) throw new Error('.auth is missing clientId');
    if (!cfg['redirect-url-local'] || !cfg['redirect-url-web']) throw new Error('.auth is missing redirect-url-local or redirect-url-web');
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
  function _pickRedirectUriForOrigin(cfg, origin) {
    const host = new URL(origin).hostname;
    const isLocal = host === 'localhost' || host === '127.0.0.1' || host === '::1';
    return isLocal ? cfg['redirect-url-local'] : cfg['redirect-url-web'];
  }

  function pickRedirectUri(cfg) {
    const chosen = _pickRedirectUriForOrigin(cfg, window.location.origin);
    Log.info(`Auth: redirect_uri=${chosen}`);
    try {
      const u = new URL(chosen);
      if (u.origin !== window.location.origin) {
        Log.warn(`Configured redirect URI origin (${u.origin}) does not match page origin (${window.location.origin}). MSAL will reject this.`);
      }
    } catch { Log.warn(`Configured redirect URI is not a valid URL: ${chosen}`); }
    return chosen;
  }

  async function init() {
    authConfig = await loadAuthConfig();
    CLIENT_ID = authConfig.clientId;
    REDIRECT_URI = pickRedirectUri(authConfig);

    await _ensureMsalLoaded();
    if (!window.msal || !window.msal.PublicClientApplication) {
      throw new Error('MSAL.js failed to load from local vendor bundle. Check script path or CSP.');
    }

    msal = new window.msal.PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/organizations', redirectUri: REDIRECT_URI },
      cache: { cacheLocation: 'localStorage' },
    });
    await msal.initialize();
    const resp = await msal.handleRedirectPromise().catch(e => { Log.err('Redirect handler:', e); return null; });
    if (_shouldClosePopupAfterAuth(resp)) {
      Log.info('Auth popup callback completed. Closing popup window.');
      try { window.close(); } catch {}
      if (!window.closed) window.location.replace('about:blank');
      return null;
    }
    if (resp && resp.account) account = resp.account;
    else { const accs = msal.getAllAccounts(); if (accs.length) account = accs[0]; }
    return account;
  }

  async function signIn(includeWriteScopes = false) {
    const scopes = includeWriteScopes ? [...SCOPES.read, ...SCOPES.write] : SCOPES.read;
    const req = { scopes, prompt: 'select_account' };
    await msal.loginRedirect(req);
    return null;
  }

  async function signOut() {
    if (!account) return;
    await msal.logoutRedirect({ account });
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
        Log.warn(`Silent token acquisition failed for [${scopes.join(',')}], using redirect`);
        await msal.acquireTokenRedirect({ scopes, account });
        return new Promise(() => {});
      }
    })();
    _tokenInflight.set(key, p);
    try { return await p; } finally { _tokenInflight.delete(key); }
  }

  function getReadToken()  { return getToken(SCOPES.read); }
  function getWriteToken() { return getToken(SCOPES.write); }
  function getAccount()    { return account; }
  function getAuthConfig() { return authConfig; }

  window.Auth = {
    init, signIn, signOut,
    getReadToken, getWriteToken,
    getAccount, getAuthConfig,
    _parseAuthFile, _pickRedirectUriForOrigin,
  };
})();
