// MSAL.js PKCE auth wrapper for PermView.
// PermView talks to multiple Microsoft APIs (Graph, ARM, PowerBI, DevOps,
// PowerPlatform, SharePoint REST, Exchange via Graph). Each API needs its own
// access token with its own resource. We acquire tokens lazily, per resource,
// and ask for additional scopes via consent only when the user actually opens
// the corresponding workload tab.
(() => {
  let CLIENT_ID = null;
  let REDIRECT_URI = null;
  let authConfig = null;
  let msal = null;
  let account = null;
  const _tokenInflight = new Map();

  // Resource -> scopes. We keep these minimal and read-only - PermView is
  // a viewer, not a manager. Sign-in initially asks only for the cheap Graph
  // scopes; everything else is requested incrementally.
  const SCOPES = {
    graphBase: [
      'https://graph.microsoft.com/User.Read',
      'https://graph.microsoft.com/Directory.Read.All',
    ],
    graphFull: [
      'https://graph.microsoft.com/User.Read',
      'https://graph.microsoft.com/Directory.Read.All',
      'https://graph.microsoft.com/Application.Read.All',
      'https://graph.microsoft.com/RoleManagement.Read.Directory',
      'https://graph.microsoft.com/Sites.Read.All',
      'https://graph.microsoft.com/Files.Read.All',
      'https://graph.microsoft.com/Calendars.Read.Shared',
    ],
    arm:      ['https://management.azure.com/user_impersonation'],
    powerbi:  ['https://analysis.windows.net/powerbi/api/Tenant.Read.All',
               'https://analysis.windows.net/powerbi/api/Workspace.Read.All'],
    devops:   ['499b84ac-1321-427f-aa17-267ca6975798/user_impersonation'],
    // PowerApps audience works for BOTH BAP (api.bap.microsoft.com) and Flow
    // (api.flow.microsoft.com) admin endpoints. The custom app registration
    // needs the "PowerApps Service" / "Power Platform API" delegated permission
    // (User) granted. This mirrors how M365PermissionsV2 handles it.
    flow:     ['https://service.powerapps.com/User'],
  };

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

    if (window.__msalLoadFailed || !window.msal || !window.msal.PublicClientApplication) {
      throw new Error('MSAL.js failed to load from CDN. Check your network or CSP.');
    }

    msal = new window.msal.PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/organizations', redirectUri: REDIRECT_URI },
      cache: { cacheLocation: 'localStorage' },
    });
    await msal.initialize();
    const resp = await msal.handleRedirectPromise().catch(e => { Log.err('Redirect handler:', e); return null; });
    if (resp && resp.account) account = resp.account;
    else { const accs = msal.getAllAccounts(); if (accs.length) account = accs[0]; }
    return account;
  }

  async function signIn() {
    const req = { scopes: SCOPES.graphBase, prompt: 'select_account' };
    try {
      const r = await msal.loginPopup(req);
      account = r.account;
    } catch (e) {
      Log.warn('Popup sign-in failed, falling back to redirect:', e.message);
      await msal.loginRedirect(req);
      return null;
    }
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
      }
    })();
    _tokenInflight.set(key, p);
    try { return await p; } finally { _tokenInflight.delete(key); }
  }

  function getGraphToken()    { return getToken(SCOPES.graphFull); }
  function getArmToken()      { return getToken(SCOPES.arm); }
  function getPowerBIToken()  { return getToken(SCOPES.powerbi); }
  function getDevOpsToken()   { return getToken(SCOPES.devops); }
  function getFlowToken()     { return getToken(SCOPES.flow); }
  // Dataverse uses a per-instance audience (https://{org}.crm{region}.dynamics.com).
  function getDataverseToken(instanceUrl) {
    if (!instanceUrl) throw new Error('Dataverse token requires an instance URL');
    const u = new URL(instanceUrl);
    return getToken([`https://${u.host}/user_impersonation`]);
  }
  function getAccount()       { return account; }
  function getAuthConfig()    { return authConfig; }

  // For SharePoint REST we'd need a per-tenant resource scope (https://{tenant}.sharepoint.com/.default).
  // That requires us to know the tenant first. PermView keeps things simple by routing
  // SharePoint queries through Microsoft Graph, so we do not expose a SharePoint REST token here.

  window.Auth = {
    init, signIn, signOut,
    getGraphToken, getArmToken, getPowerBIToken, getDevOpsToken, getFlowToken, getDataverseToken,
    getAccount, getAuthConfig,
    _parseAuthFile,
    SCOPES,
  };
})();
