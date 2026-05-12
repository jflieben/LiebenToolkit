// MSAL.js PKCE auth wrapper. SimpleMaester only talks to Microsoft Graph
// (v1.0 and beta), so we keep auth simple: one consent, refresh silently after.
(() => {
  let CLIENT_ID = null;
  let REDIRECT_URI = null;
  let authConfig = null;
  let msal = null;
  let account = null;
  let initPromise = null;
  const _tokenInflight = new Map();

  // Scopes are cumulative. We ask for everything our test catalogue might need
  // up front, so the user only sees one consent dialog. Keeping the list small
  // and readonly avoids scaring users.
  const CORE_GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Directory.Read.All',
    'https://graph.microsoft.com/Policy.Read.All',
    'https://graph.microsoft.com/RoleManagement.Read.Directory',
    'https://graph.microsoft.com/Application.Read.All',
    'https://graph.microsoft.com/UserAuthenticationMethod.Read.All',
    'https://graph.microsoft.com/AuditLog.Read.All',
    'https://graph.microsoft.com/Reports.Read.All',
  ];

  const SCOPES = {
    // Core Graph scopes consented at sign-in. All tests use this token.
    graphFull: CORE_GRAPH_SCOPES,
    // Org settings APIs require dedicated OrgSettings-* scopes.
    graphForms: [
      'https://graph.microsoft.com/OrgSettings-Forms.Read.All',
      ...CORE_GRAPH_SCOPES,
    ],
    graphAppsAndServices: [
      'https://graph.microsoft.com/OrgSettings-AppsAndServices.Read.All',
      ...CORE_GRAPH_SCOPES,
    ],
    graphSecurityEvents: [
      'https://graph.microsoft.com/SecurityEvents.Read.All',
      ...CORE_GRAPH_SCOPES,
    ],
    graphGroupRead:       null,
    // Exchange Online admin REST. We hit /adminapi/beta/{org}/InvokeCommand on
    // outlook.office365.com (resource 00000002-0000-0ff1-ce00-000000000000).
    // The only delegated scope Exchange exposes for this is `.default`, which
    // requires a tenant-specific authority (NOT /organizations) - otherwise
    // AAD can't resolve which SPN's scopes to consent. Authority is set
    // per-request in getExchangeToken() once we know the tenant id.
    exchange: ['https://outlook.office365.com/.default'],
  };

  async function loadAuthConfig() {
    const res = await fetch('./.auth', { cache: 'no-store' });
    if (!res.ok) throw new Error(`Could not load .auth (HTTP ${res.status}). Make sure SimpleMaester is being served over HTTP and the .auth file exists.`);
    const text = await res.text();
    const cfg = _parseAuthFile(text);
    if (!cfg.clientId) throw new Error('.auth is missing clientId');
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
    const host = new URL(window.location.origin).hostname;
    const isLocal = host === 'localhost' || host === '127.0.0.1' || host === '::1';
    const chosen = isLocal ? cfg['redirect-url-local'] : cfg['redirect-url-web'];
    Log.info(`Auth: redirect_uri=${chosen}`);
    return chosen;
  }

  async function init() {
    if (initPromise) return initPromise;
    initPromise = (async () => {
      if (!msal) {
        authConfig = await loadAuthConfig();
        CLIENT_ID = authConfig.clientId;
        REDIRECT_URI = pickRedirectUri(authConfig);
        if (window.__msalLoadFailed || !window.msal || !window.msal.PublicClientApplication) {
          throw new Error('MSAL.js failed to load from CDN. Check your network or CSP.');
        }
        // Authority: prefer tenantId from .auth (so SPA token redemption works
        // reliably - some AAD edge cases reject /organizations for SPA cross-origin
        // token redemption). Falls back to /organizations for true multi-tenant.
        const tenantForAuthority = authConfig.tenantId || authConfig['tenant-id'] || 'organizations';
        msal = new window.msal.PublicClientApplication({
          auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${tenantForAuthority}`, redirectUri: REDIRECT_URI },
          cache: { cacheLocation: 'localStorage' },
        });
        await msal.initialize();
      }
      const resp = await msal.handleRedirectPromise().catch(e => { Log.err('Redirect handler:', e); return null; });
      if (resp && resp.account) account = resp.account;
      else { const accs = msal.getAllAccounts(); if (accs.length) account = accs[0]; }
      return account;
    })();

    try {
      return await initPromise;
    } catch (e) {
      initPromise = null;
      throw e;
    }
  }

  async function signIn() {
    if (!msal) await init();
    if (!msal) throw new Error('Authentication is not initialized. Refresh and try again.');
    // Always use redirect for sign-in. login.microsoftonline.com sets
    // Cross-Origin-Opener-Policy: same-origin which breaks MSAL's popup flow
    // (window.closed polling is blocked). Redirect is reliable everywhere.
    const req = { scopes: SCOPES.graphFull, prompt: 'select_account' };
    await msal.loginRedirect(req);
    return null; // page navigates away; init() picks up the result via handleRedirectPromise
  }

  async function signOut() {
    if (!account) return;
    try { await msal.logoutPopup({ account }); } catch (e) { Log.warn('Logout popup failed:', e.message); }
    account = null;
  }

  async function getToken(scopes) {
    if (!account) throw new Error('Not signed in');
    // All token requests are normalised to graphFull. If a caller passes one of
    // the legacy alias arrays (graphForms, graphAppsAndServices, graphGroupRead)
    // those are null, so fall back to graphFull.
    const effectiveScopes = (scopes && scopes.length) ? scopes : SCOPES.graphFull;
    const key = effectiveScopes.join('|');
    if (_tokenInflight.has(key)) return _tokenInflight.get(key);
    const p = (async () => {
      try {
        const r = await msal.acquireTokenSilent({ scopes: effectiveScopes, account });
        return r.accessToken;
      } catch (e) {
        Log.warn(`Silent token failed for [${effectiveScopes.join(',')}], escalating to interactive`);
        // Never call acquireTokenRedirect mid-scan: it navigates away from the
        // page and kills the entire scan in progress. Instead throw an error so
        // the test can catch it and return Skipped.
        try {
          const r = await msal.acquireTokenPopup({ scopes: effectiveScopes, account });
          return r.accessToken;
        } catch (popupErr) {
          // Popup was blocked or user dismissed. Throw a descriptive error that
          // errRow() will turn into a Skipped result for this one test.
          const msg = popupErr.message || String(popupErr);
          Log.warn('Interactive token failed (popup blocked or dismissed):', msg);
          const err = new Error(`Could not get token (consent needed or popup blocked): ${msg}`);
          err.status = 401;
          throw err;
        }
      }
    })();
    _tokenInflight.set(key, p);
    try { return await p; } finally { _tokenInflight.delete(key); }
  }

  function _isDefaultScope(scope) {
    return typeof scope === 'string' && /\/\.default$/i.test(scope);
  }

  function _scopeResource(scope) {
    if (!scope || typeof scope !== 'string') return '';
    if (scope === 'openid' || scope === 'profile' || scope === 'offline_access') return 'oidc';
    if (scope.startsWith('api://')) return scope.split('/')[2] || scope;
    const m = scope.match(/^https:\/\/([^/]+)/i);
    return (m && m[1]) ? m[1].toLowerCase() : scope;
  }

  function _buildConsentGroups(scopes) {
    const uniq = Array.from(new Set((scopes || []).filter(s => !!s)));
    if (!uniq.length) return [SCOPES.graphFull.slice()];

    // AAD rejects requests mixing resource-specific delegated scopes with
    // resource .default scopes. Preflight by resource and isolate every
    // .default into its own request.
    const groups = [];
    const delegatedByResource = new Map();

    for (const s of uniq) {
      if (_isDefaultScope(s)) {
        groups.push([s]);
        continue;
      }
      const key = _scopeResource(s);
      if (!delegatedByResource.has(key)) delegatedByResource.set(key, []);
      delegatedByResource.get(key).push(s);
    }

    for (const arr of delegatedByResource.values()) groups.push(arr);
    return groups;
  }

  function _buildConsentRequest(scopes) {
    const req = { scopes, account };
    const needsExchangeAuthority = scopes.some(s => /^https:\/\/outlook\.office365\.com\/\.default$/i.test(s));
    if (needsExchangeAuthority && account?.tenantId) {
      req.authority = `https://login.microsoftonline.com/${account.tenantId}`;
    }
    return req;
  }

  // Preflight consent for a set of scopes BEFORE starting a scan. This prevents
  // mid-scan interactive prompts that would otherwise interrupt execution.
  async function ensureScopes(scopes) {
    if (!account) throw new Error('Not signed in');
    const groups = _buildConsentGroups(scopes);

    for (const groupScopes of groups) {
      const req = _buildConsentRequest(groupScopes);
      try {
        await msal.acquireTokenSilent(req);
      } catch (e) {
        Log.warn(`Preflight consent required for [${groupScopes.join(',')}]`);
        await msal.acquireTokenRedirect({ ...req, prompt: 'consent' });
        return false;
      }
    }
    return true;
  }

  function getGraphToken() { return getToken(SCOPES.graphFull); }

  // Exchange needs a tenant-specific authority for `.default` to resolve.
  // We override the authority per-request once we know the user's tenant id.
  async function getExchangeToken() {
    if (!account) throw new Error('Not signed in');
    const tenantId = account.tenantId;
    if (!tenantId) throw new Error('Cannot determine tenant id from account');
    const authority = `https://login.microsoftonline.com/${tenantId}`;
    const scopes = SCOPES.exchange;
    const key = 'EXO|' + scopes.join('|');
    if (_tokenInflight.has(key)) return _tokenInflight.get(key);
    const p = (async () => {
      try {
        const r = await msal.acquireTokenSilent({ scopes, account, authority });
        return r.accessToken;
      } catch (e) {
        Log.warn(`Silent Exchange token failed, escalating to interactive: ${e.message}`);
        const useRedirect = !!(navigator.webdriver) || /HeadlessChrome|Playwright/i.test(navigator.userAgent);
        const req = { scopes, account, authority };
        if (useRedirect) {
          await msal.acquireTokenRedirect(req);
          return new Promise(() => {});
        }
        try {
          const r = await msal.acquireTokenPopup(req);
          return r.accessToken;
        } catch (popupErr) {
          Log.warn('Exchange popup blocked, falling back to redirect:', popupErr.message);
          await msal.acquireTokenRedirect(req);
          return new Promise(() => {});
        }
      }
    })();
    _tokenInflight.set(key, p);
    try { return await p; } finally { _tokenInflight.delete(key); }
  }

  // Forced-interactive Exchange consent. Use this BEFORE the first test that
  // needs Exchange tokens. Bypasses silent cache, shows AAD consent dialog.
  async function consentExchange() {
    if (!account) throw new Error('Not signed in');
    const tenantId = account.tenantId;
    const authority = `https://login.microsoftonline.com/${tenantId}`;
    const req = { scopes: SCOPES.exchange, account, authority, prompt: 'consent' };
    const useRedirect = !!(navigator.webdriver) || /HeadlessChrome|Playwright/i.test(navigator.userAgent);
    if (useRedirect) {
      await msal.acquireTokenRedirect(req);
      return new Promise(() => {});
    }
    try {
      const r = await msal.acquireTokenPopup(req);
      return r.accessToken;
    } catch (e) {
      Log.warn('Exchange consent popup blocked, falling back to redirect:', e.message);
      await msal.acquireTokenRedirect(req);
      return new Promise(() => {});
    }
  }
  function getAccount()    { return account; }
  function getUserPrincipalName() { return account?.username || null; }

  window.Auth = { init, signIn, signOut, getGraphToken, getExchangeToken, consentExchange, getToken, ensureScopes, getAccount, getUserPrincipalName, SCOPES };
})();
