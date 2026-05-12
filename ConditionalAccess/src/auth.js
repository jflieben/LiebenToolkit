(() => {
  let authConfig = null;
  let msal = null;
  let account = null;
  const tokenInflight = new Map();

  const SCOPES = {
    read: [
      'https://graph.microsoft.com/Policy.Read.All',
      'https://graph.microsoft.com/Directory.Read.All',
      'https://graph.microsoft.com/User.Read',
    ],
    write: [
      'https://graph.microsoft.com/Policy.ReadWrite.ConditionalAccess',
      'https://graph.microsoft.com/Directory.Read.All',
      'https://graph.microsoft.com/User.Read',
    ],
  };

  async function loadAuthConfig() {
    const res = await fetch('./.auth', { cache: 'no-store' });
    if (!res.ok) {
      throw new Error(`Could not load .auth file (HTTP ${res.status}). Serve this folder over HTTP.`);
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
      const k = line.substring(0, idx).trim();
      const v = line.substring(idx + 1).trim();
      cfg[k] = v;
    }
    return cfg;
  }

  function pickRedirectUri(cfg) {
    const host = new URL(window.location.origin).hostname;
    const isLocal = host === 'localhost' || host === '127.0.0.1' || host === '::1';
    const chosen = isLocal ? cfg['redirect-url-local'] : cfg['redirect-url-web'];
    Log.info(`Auth redirect URI: ${chosen}`);
    return chosen;
  }

  async function init() {
    authConfig = await loadAuthConfig();
    if (window.__msalLoadFailed || !window.msal || !window.msal.PublicClientApplication) {
      throw new Error('MSAL.js failed to load from CDN.');
    }

    msal = new window.msal.PublicClientApplication({
      auth: {
        clientId: authConfig.clientId,
        authority: 'https://login.microsoftonline.com/organizations',
        redirectUri: pickRedirectUri(authConfig),
      },
      cache: { cacheLocation: 'localStorage' },
    });
    await msal.initialize();

    const response = await msal.handleRedirectPromise().catch((e) => {
      Log.err('Redirect handler failed:', e.message);
      return null;
    });
    if (response && response.account) {
      account = response.account;
    } else {
      const accounts = msal.getAllAccounts();
      if (accounts.length) account = accounts[0];
    }
    return account;
  }

  async function signIn(includeWriteScopes = false) {
    const scopes = includeWriteScopes ? [...SCOPES.write] : [...SCOPES.read];
    const req = { scopes, prompt: 'select_account' };
    try {
      const res = await msal.loginPopup(req);
      account = res.account;
      return account;
    } catch (e) {
      Log.warn('Popup sign-in failed, trying redirect:', e.message);
      await msal.loginRedirect(req);
      return null;
    }
  }

  async function signOut() {
    if (!account) return;
    try {
      await msal.logoutPopup({ account });
    } catch (e) {
      Log.warn('Logout popup failed:', e.message);
    }
    account = null;
  }

  async function getToken(scopes) {
    if (!account) throw new Error('Not signed in.');
    const key = scopes.join('|');
    if (tokenInflight.has(key)) return tokenInflight.get(key);
    const promise = (async () => {
      try {
        const r = await msal.acquireTokenSilent({ scopes, account });
        return r.accessToken;
      } catch {
        const r = await msal.acquireTokenPopup({ scopes, account });
        return r.accessToken;
      }
    })();
    tokenInflight.set(key, promise);
    try {
      return await promise;
    } finally {
      tokenInflight.delete(key);
    }
  }

  window.Auth = {
    init,
    signIn,
    signOut,
    getReadToken: () => getToken(SCOPES.read),
    getWriteToken: () => getToken(SCOPES.write),
    getAccount: () => account,
  };
})();