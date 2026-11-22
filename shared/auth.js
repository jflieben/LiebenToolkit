// Shared MSAL (PKCE) wrapper. Each tool creates its own instance:
//
//   window.Auth = ToolkitAuth.create({
//     scopes: { read: [...], write: [...] },   // named sets, 'read' is used to sign in
//     login: 'redirect',                         // or 'popup' when state must survive sign-in
//   });
//
// Security choices that apply to every tool:
// - Token cache in sessionStorage: tokens die with the tab and are not shared with
//   other tabs or other pages on this origin later on.
// - One explicit active account per tab, never "the first cached account", so a tool
//   cannot silently act in a tenant you used elsewhere.
// - Extra consent (write scopes) by popup, so the page and its state stay put. Falls back
//   to a redirect only when the popup is blocked.
(() => {
  const AUTHORITY = 'https://login.microsoftonline.com/organizations';
  const MSAL_FILE = 'msal-browser.3.27.0.min.js';

  function parseAuthFile(text) {
    const cfg = {};
    for (const raw of String(text).split(/\r?\n/)) {
      const line = raw.trim();
      if (!line || line.startsWith('#')) continue;
      const idx = line.indexOf(':');
      if (idx < 0) continue;
      cfg[line.substring(0, idx).trim()] = line.substring(idx + 1).trim();
    }
    return cfg;
  }

  function isLocalHost(hostname) {
    return hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '::1' || hostname === '[::1]';
  }

  function pickRedirectUriForOrigin(cfg, origin) {
    return isLocalHost(new URL(origin).hostname) ? cfg['redirect-url-local'] : cfg['redirect-url-web'];
  }

  async function loadAuthConfig() {
    const res = await fetch('./.auth', { cache: 'no-store', credentials: 'same-origin' });
    if (!res.ok) throw new Error(`Could not load .auth (HTTP ${res.status}). Serve the tool over HTTP with .auth in the tool folder.`);
    const cfg = parseAuthFile(await res.text());
    if (!/^[0-9a-f-]{36}$/i.test(cfg.clientId || '')) throw new Error('.auth has no valid clientId');
    if (!cfg['redirect-url-local'] || !cfg['redirect-url-web']) throw new Error('.auth is missing redirect-url-local or redirect-url-web');
    return cfg;
  }

  // This page was opened by MSAL as a sign-in popup and now holds the response. The
  // opener reads it; this page must not boot the tool or consume the response itself.
  function isAuthPopup() {
    try {
      if (!window.opener || window.opener === window) return false;
      if (!String(window.name || '').startsWith('msal.')) return false;
      const bits = `${location.hash || ''}&${location.search || ''}`;
      return /(?:^|[?&#])(code|error|state)=/i.test(bits);
    } catch { return false; }
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error(`Failed to load ${src}`));
      document.head.appendChild(s);
    });
  }

  async function ensureMsal() {
    if (window.msal && window.msal.PublicClientApplication) return;
    const base = window.Toolkit && Toolkit.paths ? Toolkit.paths.root : '../';
    await loadScript(`${base}vendor/${MSAL_FILE}`);
    if (!window.msal || !window.msal.PublicClientApplication) throw new Error('MSAL.js did not load. Check the vendor folder and the CSP.');
  }

  function isCacheQuota(e) {
    return !!e && (e.errorCode === 'cache_quota_exceeded' || /cache_quota_exceeded|Exceeded cache storage/i.test(e.message || ''));
  }
  function isPopupBlocked(e) {
    return !!e && /popup_window_error|empty_window_error|popup.*block/i.test(`${e.errorCode || ''} ${e.message || ''}`);
  }
  function isUserCancel(e) {
    return !!e && /user_cancelled|user_canceled/i.test(`${e.errorCode || ''} ${e.message || ''}`);
  }
  function needsInteraction(e) {
    if (!e) return false;
    if (window.msal && e instanceof window.msal.InteractionRequiredAuthError) return true;
    return /interaction_required|consent_required|login_required|no_account|no_tokens_found|invalid_grant|AADSTS65001|AADSTS50076|AADSTS50079|AADSTS70044/i.test(`${e.errorCode || ''} ${e.message || ''}`);
  }

  // Drop cached access tokens (not refresh tokens) when storage is full.
  function freeAccessTokens() {
    let removed = 0;
    try {
      for (let i = sessionStorage.length - 1; i >= 0; i--) {
        const k = sessionStorage.key(i);
        const v = k && sessionStorage.getItem(k);
        if (!v || v[0] !== '{') continue;
        try { if (JSON.parse(v).credentialType === 'AccessToken') { sessionStorage.removeItem(k); removed++; } } catch { /* not ours */ }
      }
    } catch { /* storage blocked */ }
    return removed;
  }

  // Older builds cached MSAL tokens, refresh tokens included, in localStorage. The toolkit
  // now uses sessionStorage only, so anything MSAL left in localStorage is removed.
  function purgeLegacyCache() {
    let removed = 0;
    try {
      for (let i = localStorage.length - 1; i >= 0; i--) {
        const k = localStorage.key(i);
        if (!k) continue;
        let msalEntry = k.startsWith('msal.') || /-(login\.windows\.net|login\.microsoftonline\.com)-/i.test(k);
        if (!msalEntry) {
          const v = localStorage.getItem(k);
          if (v && v[0] === '{') {
            try { const j = JSON.parse(v); msalEntry = !!(j && (j.credentialType || (j.homeAccountId && j.environment))); } catch { /* not JSON */ }
          }
        }
        if (msalEntry) { localStorage.removeItem(k); removed++; }
      }
    } catch { /* storage blocked */ }
    if (removed && window.Log) Log.info(`Removed ${removed} sign-in cache entries an older version left in local storage.`);
    return removed;
  }

  function create({ scopes = {}, login = 'redirect', cacheLocation = 'sessionStorage' } = {}) {
    const SETS = { read: ['User.Read'], ...scopes };
    let msal = null;
    let account = null;
    let authConfig = null;
    let redirectUri = null;
    const inflight = new Map();

    // Bare names like 'User.Read.All' are Graph; anything with a resource part
    // (https://..., api://..., <appId>/scope) is passed through untouched.
    const full = (s) => (/[/:]/.test(s) || /^(openid|profile|offline_access|email)$/i.test(s) ? s : `https://graph.microsoft.com/${s}`);
    const resolve = (set) => [...new Set((Array.isArray(set) ? set : (SETS[set] || [])).map(full))];

    async function init() {
      if (isAuthPopup()) return null;
      authConfig = await loadAuthConfig();
      redirectUri = pickRedirectUriForOrigin(authConfig, location.origin);
      try {
        if (new URL(redirectUri).origin !== location.origin) Log.warn(`Redirect URI ${redirectUri} is not on this origin (${location.origin}); sign-in will fail.`);
      } catch { Log.warn(`Redirect URI in .auth is not a valid URL: ${redirectUri}`); }
      await ensureMsal();
      if (cacheLocation === 'sessionStorage') purgeLegacyCache();
      msal = new window.msal.PublicClientApplication({
        auth: { clientId: authConfig.clientId, authority: AUTHORITY, redirectUri, navigateToLoginRequestUrl: true },
        cache: { cacheLocation, temporaryCacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
        system: { allowNativeBroker: false },
      });
      await msal.initialize();
      const resp = await msal.handleRedirectPromise().catch((e) => {
        if (!isUserCancel(e)) Log.err('Sign-in response could not be processed:', e.message || e);
        return null;
      });
      if (resp && resp.account) account = resp.account;
      else account = msal.getActiveAccount() || null;
      // A tab can only have one account it signed in with itself. Several cached
      // accounts and no active one means we do not guess.
      if (!account) {
        const all = msal.getAllAccounts();
        if (all.length === 1) account = all[0];
      }
      if (account) msal.setActiveAccount(account);
      Log.info(account ? `Signed in as ${account.username} (tenant ${account.tenantId})` : 'Not signed in');
      return account;
    }

    async function signIn(set = 'read', { loginHint, prompt = 'select_account' } = {}) {
      if (!msal) throw new Error('Auth not initialised');
      const req = { scopes: resolve(set), prompt: loginHint ? undefined : prompt, loginHint };
      if (login === 'popup') {
        let resp;
        try { resp = await msal.loginPopup(req); }
        catch (e) {
          if (isPopupBlocked(e)) throw new Error('The sign-in popup was blocked. Allow popups for this site and try again.');
          throw e;
        }
        account = resp.account;
        msal.setActiveAccount(account);
        return account;
      }
      await msal.loginRedirect(req);
      return null;
    }

    async function signOut() {
      const acc = account;
      account = null;
      if (!msal || !acc) return;
      // Clear this tab's cache and end the session at Microsoft as well.
      try { await msal.logoutRedirect({ account: acc, postLogoutRedirectUri: redirectUri }); }
      catch (e) { Log.warn('Sign-out redirect failed, clearing local cache:', e.message || e); await msal.clearCache().catch(() => {}); }
    }

    // noRedirect: when the popup is blocked, fail instead of navigating away (use during runs).
    // extraConsent: scopes of other resources to consent to in the same popup.
    async function acquire(scopesList, { interactive = true, noRedirect = false, extraConsent } = {}) {
      const req = { scopes: scopesList, account };
      if (extraConsent && extraConsent.length) req.extraScopesToConsent = extraConsent;
      const silent = async () => {
        try { return (await msal.acquireTokenSilent(req)).accessToken; }
        catch (e) {
          if (!isCacheQuota(e)) throw e;
          Log.warn(`Token cache full, cleared ${freeAccessTokens()} cached access token(s).`);
          return (await msal.acquireTokenSilent(req)).accessToken;
        }
      };
      try {
        return await silent();
      } catch (e) {
        if (!interactive || !needsInteraction(e)) throw e;
        Log.info(`Asking for consent or sign-in for: ${scopesList.join(', ')}`);
        try {
          return (await msal.acquireTokenPopup(req)).accessToken;
        } catch (pe) {
          if (isUserCancel(pe)) throw new Error('Sign-in or consent was cancelled.');
          if (!isPopupBlocked(pe)) throw pe;
          if (noRedirect) throw new Error('The consent popup was blocked. Allow popups for this site and try again.');
          Log.warn('Popup blocked, continuing with a full-page redirect.');
          await msal.acquireTokenRedirect(req);
          return new Promise(() => {});
        }
      }
    }

    async function getToken(set, opts = {}) {
      if (!account) throw new Error('Not signed in');
      const list = resolve(set);
      if (opts.extraConsent) opts = { ...opts, extraConsent: opts.extraConsent.flatMap((s) => resolve(s)) };
      const key = `${opts.interactive === false ? 's' : 'i'}|${list.join(' ')}`;
      if (inflight.has(key)) return inflight.get(key);
      const p = acquire(list, opts);
      inflight.set(key, p);
      try { return await p; } finally { inflight.delete(key); }
    }

    // Consent to several sets (different resources) with one popup before a run starts,
    // then warm each token silently.
    async function consent(sets, { noRedirect = true } = {}) {
      const missing = [];
      for (const s of sets) if (!(await hasConsent(s))) missing.push(s);
      if (!missing.length) return;
      await getToken(missing[0], { noRedirect, extraConsent: missing.slice(1) });
      for (const s of missing.slice(1)) await getToken(s, { noRedirect });
    }

    // Silent check whether a scope set is already consented.
    async function hasConsent(set) {
      try { await getToken(set, { interactive: false }); return true; } catch { return false; }
    }

    function tenantInfo() {
      if (!account) return null;
      const claims = account.idTokenClaims || {};
      const upn = account.username || '';
      return { tenantId: account.tenantId || claims.tid || null, domain: upn.includes('@') ? upn.split('@')[1] : null, upn, name: account.name || claims.name || upn };
    }

    // Header chip data for Toolkit.setUser.
    function userInfo(onSignOut) {
      const t = tenantInfo();
      return t ? { name: t.name, upn: t.upn, tenant: t.domain || t.tenantId, onSignOut } : null;
    }

    return {
      init, signIn, signOut, getToken, hasConsent, consent,
      getReadToken: () => getToken('read'),
      getWriteToken: () => getToken(SETS.write ? [...SETS.read, ...SETS.write] : 'read'),
      getAccount: () => account,
      getAuthConfig: () => authConfig,
      getRedirectUri: () => redirectUri,
      getMsal: () => msal,
      tenantInfo, userInfo, scopes: SETS, resolve,
    };
  }

  window.ToolkitAuth = {
    create, isAuthPopup, parseAuthFile, pickRedirectUriForOrigin, loadAuthConfig, ensureMsal,
    isCacheQuota, isPopupBlocked, isUserCancel, needsInteraction, freeAccessTokens, purgeLegacyCache, AUTHORITY,
  };
})();
