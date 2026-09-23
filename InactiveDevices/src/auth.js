// MSAL.js PKCE auth wrapper for InactiveDevices.
(() => {
  let CLIENT_ID = null;
  let REDIRECT_URI = null;
  let authConfig = null;
  let msal = null;
  let account = null;
  const tokenInflight = new Map();

  const SCOPES = {
    read: [
      'https://graph.microsoft.com/Device.Read.All',
      'https://graph.microsoft.com/AuditLog.Read.All',
      'https://graph.microsoft.com/Directory.Read.All',
      'https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All',
      'https://graph.microsoft.com/User.Read',
    ],
    write: [
      'https://graph.microsoft.com/Device.ReadWrite.All',
      'https://graph.microsoft.com/Directory.AccessAsUser.All',
    ],
  };

  function shouldClosePopupAfterAuth(resp) {
    if (!window.opener || window.opener.closed) return false;
    if (resp && resp.account) return true;
    const authBits = `${window.location.hash || ''}&${window.location.search || ''}`;
    return /(?:^|[?&#])(code|error|id_token|access_token|state)=/i.test(authBits);
  }

  async function loadAuthConfig() {
    const res = await fetch('./.auth', { cache: 'no-store' });
    if (!res.ok) {
      throw new Error(`Could not load .auth file (HTTP ${res.status}). Make sure the tool is served over HTTP and the .auth file exists in this folder.`);
    }
    const text = await res.text();
    const cfg = parseAuthFile(text);
    if (!cfg.clientId) throw new Error('.auth is missing clientId');
    if (!cfg['redirect-url-local'] || !cfg['redirect-url-web']) {
      throw new Error('.auth is missing redirect-url-local or redirect-url-web');
    }
    return cfg;
  }

  function parseAuthFile(text) {
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

  function pickRedirectUriForOrigin(cfg, origin) {
    const host = new URL(origin).hostname;
    const isLocal = host === 'localhost' || host === '127.0.0.1' || host === '::1';
    return isLocal ? cfg['redirect-url-local'] : cfg['redirect-url-web'];
  }

  function pickRedirectUri(cfg) {
    const chosen = pickRedirectUriForOrigin(cfg, window.location.origin);
    Log.info(`Auth: redirect_uri=${chosen}`);
    try {
      const u = new URL(chosen);
      if (u.origin !== window.location.origin) {
        Log.warn(`Configured redirect URI origin (${u.origin}) does not match page origin (${window.location.origin}). MSAL will reject this.`);
      }
    } catch {
      Log.warn(`Configured redirect URI is not a valid URL: ${chosen}`);
    }
    return chosen;
  }

  async function init() {
    authConfig = await loadAuthConfig();
    CLIENT_ID = authConfig.clientId;
    REDIRECT_URI = pickRedirectUri(authConfig);

    if (window.__msalLoadFailed || !window.msal || !window.msal.PublicClientApplication) {
      throw new Error('MSAL.js failed to load from CDN. Check your network or CSP.');
    }

    msal = new window.msal.PublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: 'https://login.microsoftonline.com/organizations',
        redirectUri: REDIRECT_URI,
      },
      cache: { cacheLocation: 'localStorage' },
    });

    await msal.initialize();
    const resp = await msal.handleRedirectPromise().catch((e) => {
      Log.err('Redirect handler:', e);
      return null;
    });
    if (shouldClosePopupAfterAuth(resp)) {
      Log.info('Auth popup callback completed. Closing popup window.');
      try { window.close(); } catch {}
      if (!window.closed) window.location.replace('about:blank');
      return null;
    }

    if (resp && resp.account) account = resp.account;
    else {
      const accs = msal.getAllAccounts();
      if (accs.length) account = accs[0];
    }
    return account;
  }

  async function signIn() {
    const req = { scopes: SCOPES.read, prompt: 'select_account' };
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
    if (tokenInflight.has(key)) return tokenInflight.get(key);
    const p = (async () => {
      try {
        const r = await msal.acquireTokenSilent({ scopes, account });
        return r.accessToken;
      } catch (e) {
        Log.warn(`Silent token acquisition failed for [${scopes.join(', ')}], using redirect`);
        await msal.acquireTokenRedirect({ scopes, account });
        return new Promise(() => {});
      }
    })();

    tokenInflight.set(key, p);
    try { return await p; }
    finally { tokenInflight.delete(key); }
  }

  function getReadToken() { return getToken(SCOPES.read); }
  function getWriteToken() { return getToken([...SCOPES.read, ...SCOPES.write]); }
  function getAccount() { return account; }

  window.Auth = {
    init,
    signIn,
    signOut,
    getReadToken,
    getWriteToken,
    getAccount,
    parseAuthFile,
    pickRedirectUriForOrigin,
  };
})();
