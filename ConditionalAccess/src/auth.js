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

  function _shouldClosePopupAfterAuth(resp) {
    if (!window.opener || window.opener.closed) return false;
    if (resp && resp.account) return true;
    const authBits = `${window.location.hash || ''}&${window.location.search || ''}`;
    return /(?:^|[?&#])(code|error|id_token|access_token|state)=/i.test(authBits);
  }

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
    if (_shouldClosePopupAfterAuth(response)) {
      Log.info('Auth popup callback completed. Closing popup window.');
      try { window.close(); } catch {}
      if (!window.closed) window.location.replace('about:blank');
      return null;
    }
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
    await msal.loginRedirect(req);
    return null;
  }

  async function signOut() {
    if (!account) return;
    await msal.logoutRedirect({ account });
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
      } catch (e) {
        Log.warn(`Silent token failed for [${scopes.join(', ')}], using redirect:`, e.message || e);
        await msal.acquireTokenRedirect({ scopes, account });
        return new Promise(() => {});
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