(() => {
  // NOTE: MSAL 3.27.0 iframe sandbox warnings
  // ============================================
  // When MSAL performs redirect-based auth, it creates a sandboxed iframe with the sandbox attributes:
  // - allow-scripts + allow-same-origin (for token requests)
  // - Missing allow-top-navigation (for navigating the parent window after auth completes)
  // 
  // This produces non-blocking browser warnings:
  // 1. "An iframe which has both allow-scripts and allow-same-origin for its sandbox attribute can escape its sandboxing"
  // 2. "Unsafe attempt to initiate navigation for frame with origin ... The frame attempting navigation is sandboxed,
  //     but the flag of 'allow-top-navigation' or 'allow-top-navigation-by-user-activation' is not set"
  //
  // These warnings are cosmetic and do NOT prevent the auth flow from completing. The redirect returns and the app
  // resumes normally. They are suppressed in app.js via console.warn filter. Consider upgrading MSAL in a future
  // release to address this upstream (e.g., MSAL 4.x or later may have improved iframe handling).
  //
  // Reference: https://github.com/AzureAD/microsoft-authentication-library-for-js/issues/

  let msalApp = null;
  let account = null;
  let cfg = null;
  let lastRedirectError = null;

  const CORE_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Directory.Read.All',
    'https://graph.microsoft.com/Policy.Read.All',
    'https://graph.microsoft.com/Application.Read.All',
    'https://graph.microsoft.com/RoleManagement.Read.Directory',
    'https://graph.microsoft.com/UserAuthenticationMethod.Read.All',
    'https://graph.microsoft.com/AuditLog.Read.All',
    'https://graph.microsoft.com/Reports.Read.All',
    'https://graph.microsoft.com/AccessReview.Read.All',
    'https://graph.microsoft.com/DirectoryRecommendations.Read.All',
    'https://graph.microsoft.com/DeviceManagementConfiguration.Read.All',
    'https://graph.microsoft.com/DeviceManagementServiceConfig.Read.All',
    'https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All',
    'https://graph.microsoft.com/DeviceManagementApps.Read.All',
    'https://graph.microsoft.com/EntitlementManagement.Read.All',
    'https://graph.microsoft.com/DeviceManagementRBAC.Read.All',
    'https://graph.microsoft.com/IdentityRiskEvent.Read.All',
    'https://graph.microsoft.com/IdentityRiskyUser.Read.All',
  ];

  function parseAuth(text) {
    const map = {};
    for (const raw of text.split(/\r?\n/)) {
      const line = raw.trim();
      if (!line || line.startsWith('#')) continue;
      const idx = line.indexOf(':');
      if (idx < 0) continue;
      map[line.slice(0, idx).trim()] = line.slice(idx + 1).trim();
    }
    return map;
  }

  function pickRedirect(authCfg) {
    const host = new URL(window.location.origin).hostname;
    const local = host === 'localhost' || host === '127.0.0.1' || host === '::1';
    return local ? authCfg['redirect-url-local'] : authCfg['redirect-url-web'];
  }

  function isConsentRequiredError(err) {
    const text = `${err?.errorCode || ''} ${err?.subError || ''} ${err?.message || ''}`.toLowerCase();
    return text.includes('consent') || text.includes('65001') || text.includes('permission');
  }

  function isInteractionRequiredError(err) {
    const text = `${err?.errorCode || ''} ${err?.subError || ''} ${err?.message || ''}`.toLowerCase();
    return text.includes('interaction_required') || text.includes('login_required') || text.includes('consent_required');
  }

  function resourceKey(scope) {
    const s = `${scope || ''}`.trim().toLowerCase();
    if (s.startsWith('https://management.azure.com/')) return 'arm';
    if (s.startsWith('https://graph.microsoft.com/')) return 'graph';
    if (s === 'openid' || s === 'profile' || s === 'offline_access') return 'oidc';
    const i = s.indexOf('/');
    return i > 0 ? s.slice(0, i) : s;
  }

  function groupScopesByResource(scopes) {
    const groups = new Map();
    for (const scope of (Array.isArray(scopes) ? scopes : [])) {
      const key = resourceKey(scope);
      if (!groups.has(key)) groups.set(key, []);
      const list = groups.get(key);
      if (!list.includes(scope)) list.push(scope);
    }
    // OIDC-only scopes should not trigger standalone token requests.
    groups.delete('oidc');
    return Array.from(groups.values());
  }

  async function init() {
    if (window.__msalLoadFailed || !window.msal || !window.msal.PublicClientApplication) {
      throw new Error('MSAL failed to load.');
    }
    const res = await fetch('./.auth', { cache: 'no-store' });
    if (!res.ok) throw new Error(`Cannot load .auth (${res.status})`);
    cfg = parseAuth(await res.text());
    if (!cfg.clientId) throw new Error('Missing clientId in .auth');

    const authorityTenant = cfg.tenantId || cfg['tenant-id'] || 'organizations';
    msalApp = new window.msal.PublicClientApplication({
      auth: {
        clientId: cfg.clientId,
        authority: `https://login.microsoftonline.com/${authorityTenant}`,
        redirectUri: pickRedirect(cfg),
      },
      cache: { cacheLocation: 'localStorage' },
    });

    await msalApp.initialize();
    const resp = await msalApp.handleRedirectPromise().catch((err) => {
      lastRedirectError = err;
      console.error('MSAL redirect handling failed:', err);
      return null;
    });
    if (resp?.account) account = resp.account;
    if (!account) {
      const all = msalApp.getAllAccounts();
      if (all.length) account = all[0];
    }

    return account;
  }

  async function signIn() {
    await msalApp.loginRedirect({ scopes: CORE_SCOPES, prompt: 'select_account' });
  }

  async function signOut() {
    if (!account) return;
    await msalApp.logoutRedirect({ account });
  }

  async function getToken(scopes) {
    if (!account) throw new Error('Not signed in');
    const wanted = (Array.isArray(scopes) && scopes.length) ? scopes : CORE_SCOPES;
    try {
      const s = await msalApp.acquireTokenSilent({ scopes: wanted, account });
      return s.accessToken;
    } catch (err) {
      // Fall back to redirect so this also works in popup-restricted browser modes.
      // Do not force consent each time; let Entra decide whether consent is needed.
      await msalApp.acquireTokenRedirect({
        scopes: wanted,
        account,
        loginHint: account?.username,
        ...(isConsentRequiredError(err) ? { prompt: 'consent' } : {}),
      });
      return null;
    }
  }

  async function getTokenSilent(scopes) {
    if (!account) throw new Error('Not signed in');
    const wanted = (Array.isArray(scopes) && scopes.length) ? scopes : CORE_SCOPES;
    try {
      const s = await msalApp.acquireTokenSilent({ scopes: wanted, account });
      return s.accessToken;
    } catch {
      return null;
    }
  }

  async function ensureScopes(scopes, options) {
    const opts = options || {};
    if (!account) throw new Error('Not signed in');
    const wanted = (Array.isArray(scopes) && scopes.length) ? scopes : CORE_SCOPES;
    const groups = groupScopesByResource(wanted);

    for (const group of groups) {
      try {
        await msalApp.acquireTokenSilent({ scopes: group, account });
      } catch (err) {
        const shouldPromptConsent = !!opts.forceConsent || isConsentRequiredError(err);
        if (!isInteractionRequiredError(err) && !shouldPromptConsent) {
          throw err;
        }
        await msalApp.acquireTokenRedirect({
          scopes: group,
          account,
          loginHint: account?.username,
          ...(shouldPromptConsent ? { prompt: 'consent' } : { prompt: 'select_account' }),
        });
        return false;
      }
    }

    return true;
  }

  window.Auth = {
    init,
    signIn,
    signOut,
    getToken,
    getTokenSilent,
    ensureScopes,
    getAccount: () => account,
    getLastRedirectError: () => lastRedirectError,
    coreScopes: CORE_SCOPES,
  };
})();
