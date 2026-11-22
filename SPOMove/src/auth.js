// Two tenant connections (source and target) on one MSAL instance, built on the shared
// ToolkitAuth helpers. Sign-in is popup only, so the wizard state survives a second login.
// Tokens live in sessionStorage. A role only ever uses an account this tab signed in for
// that role; the role-to-account map (no tokens) is kept in sessionStorage for reloads.
(() => {
  const GRAPH_SCOPES = ['https://graph.microsoft.com/User.Read'];
  const MAP_KEY = 'SPOMove-connections';
  const ROLES = ['source', 'target'];

  let msal = null;
  let authConfig = null;
  let redirectUri = null;
  const connections = { source: empty('source'), target: empty('target') };
  const inflight = new Map();

  function empty(role) {
    return { role, account: null, tenantHost: null, adminHost: null, tenantId: null, tenantName: null, upn: null, spoConsent: false };
  }
  const sameUpn = (a, b) => !!a && !!b && String(a).toLowerCase() === String(b).toLowerCase();

  function saveMap() {
    const map = {};
    for (const role of ROLES) {
      const c = connections[role];
      if (c.account) map[role] = { homeAccountId: c.account.homeAccountId, upn: c.upn, tenantId: c.tenantId };
    }
    try { sessionStorage.setItem(MAP_KEY, JSON.stringify(map)); } catch { /* storage blocked */ }
  }
  function loadMap() {
    try { return JSON.parse(sessionStorage.getItem(MAP_KEY) || '{}') || {}; } catch { return {}; }
  }

  async function init() {
    if (ToolkitAuth.isAuthPopup()) return false;
    authConfig = await ToolkitAuth.loadAuthConfig();
    redirectUri = ToolkitAuth.pickRedirectUriForOrigin(authConfig, location.origin);
    try {
      if (new URL(redirectUri).origin !== location.origin) Log.warn(`Redirect URI ${redirectUri} is not on this origin (${location.origin}); sign-in will fail.`);
    } catch { Log.warn(`Redirect URI in .auth is not a valid URL: ${redirectUri}`); }
    await ToolkitAuth.ensureMsal();
    ToolkitAuth.purgeLegacyCache();
    msal = new window.msal.PublicClientApplication({
      auth: { clientId: authConfig.clientId, authority: ToolkitAuth.AUTHORITY, redirectUri, navigateToLoginRequestUrl: false },
      cache: { cacheLocation: 'sessionStorage', temporaryCacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
      system: { allowNativeBroker: false },
    });
    await msal.initialize();
    await msal.handleRedirectPromise().catch((e) => { Log.dbg('handleRedirectPromise:', e.message || e); return null; });

    // After a reload, re-attach the accounts this tab signed in per role.
    const map = loadMap();
    for (const role of ROLES) {
      const m = map[role];
      if (!m || !m.homeAccountId) continue;
      const acc = msal.getAllAccounts().find((a) => a.homeAccountId === m.homeAccountId);
      if (acc && await attach(role, acc, { tenantId: m.tenantId, upn: m.upn })) {
        Log.info(`${role} connection restored: ${connections[role].tenantName || connections[role].tenantHost} (${connections[role].upn})`);
      }
    }
    saveMap();
    return true;
  }

  // Bind a cached account to a role without any popup. Only used for accounts this tab
  // signed in itself, and only when tenant and UPN match what is expected.
  async function attach(role, account, expect = {}) {
    const conn = empty(role);
    conn.account = account;
    conn.upn = account.username || null;
    connections[role] = conn;
    try {
      await detectTenant(conn, { silent: true });
    } catch (e) {
      Log.dbg(`Silent restore of ${role} as ${account.username} failed: ${e.message}`);
      connections[role] = empty(role);
      return false;
    }
    if ((expect.tenantId && conn.tenantId !== expect.tenantId) || (expect.upn && !sameUpn(conn.upn, expect.upn))) {
      connections[role] = empty(role);
      return false;
    }
    conn.spoConsent = await hasSharePointConsent(role);
    return true;
  }

  // loginHint pre-fills the account; expectTenantId / expectTenantHost reject a sign-in to
  // another tenant (reconnecting for an existing job), so files never go to the wrong tenant.
  async function signIn(role, { loginHint, expectTenantId, expectTenantHost } = {}) {
    if (!ROLES.includes(role)) throw new Error(`Unknown connection role: ${role}`);
    if (!msal) throw new Error('Auth not initialised');
    Log.info(`Opening sign-in popup for the ${role} tenant...`);
    const req = { scopes: GRAPH_SCOPES };
    if (loginHint) req.loginHint = loginHint; else req.prompt = 'select_account';
    let resp;
    try {
      resp = await msal.loginPopup(req);
    } catch (e) {
      if (ToolkitAuth.isPopupBlocked(e)) throw new Error('The sign-in popup was blocked. Allow popups for this site and try again.');
      if (ToolkitAuth.isUserCancel(e)) throw new Error('Sign-in was cancelled.');
      throw e;
    }
    const conn = empty(role);
    conn.account = resp.account;
    conn.upn = resp.account ? resp.account.username : null;
    connections[role] = conn;
    try {
      await detectTenant(conn);
    } catch (e) {
      connections[role] = empty(role);
      saveMap();
      throw e;
    }
    if ((expectTenantHost && conn.tenantHost && conn.tenantHost !== expectTenantHost) ||
        (expectTenantId && conn.tenantId !== expectTenantId)) {
      const got = conn.tenantName || conn.tenantHost || conn.upn;
      connections[role] = empty(role);
      saveMap();
      throw new Error(`Signed in to ${got}, but this job's ${role} tenant is ${expectTenantHost || expectTenantId}. Use an account in the right tenant.`);
    }
    Log.info(`${role} tenant connected: ${conn.tenantName || conn.tenantHost} (${conn.upn})`);
    conn.spoConsent = await hasSharePointConsent(role);
    if (!conn.spoConsent) Log.warn(`SharePoint access not approved yet for the ${role} tenant.`);
    saveMap();
    return conn;
  }

  // Clears both connections and this tab's token cache.
  async function signOutAll() {
    connections.source = empty('source');
    connections.target = empty('target');
    saveMap();
    inflight.clear();
    if (msal) await msal.clearCache().catch((e) => Log.dbg('clearCache:', e.message || e));
  }

  async function detectTenant(conn, { silent = false } = {}) {
    try {
      const org = await Http.json('https://graph.microsoft.com/v1.0/organization?$select=id,displayName,verifiedDomains', {
        label: `Graph[${conn.role}]`, token: () => getToken(conn.role, GRAPH_SCOPES, { silent }), retries: 3,
      });
      const o = (org && org.value && org.value[0]) || {};
      conn.tenantId = o.id || (conn.account && conn.account.tenantId) || null;
      conn.tenantName = o.displayName || null;
      const verified = o.verifiedDomains || [];
      const initial = (verified.find((d) => d.isInitial) || verified[0] || {}).name || '';
      const prefix = initial.split('.')[0];
      if (!/^[a-z0-9]+$/i.test(prefix)) throw new Error('no initial onmicrosoft.com domain found');
      conn.tenantHost = `${prefix.toLowerCase()}.sharepoint.com`;
      conn.adminHost = `${prefix.toLowerCase()}-admin.sharepoint.com`;
      Log.info(`Detected ${conn.role} SharePoint host: ${conn.tenantHost}`);
    } catch (e) {
      throw new Error(`Could not resolve the SharePoint host for the ${conn.role} tenant: ${e.message}`);
    }
  }

  // Re-attach a role for an existing job from accounts this tab signed in. Only an account
  // with the job's tenant id AND UPN qualifies; anything else means a fresh sign-in.
  async function restoreConnection(role, loc) {
    if (!msal || !loc || !loc.tenantId || !loc.upn) return false;
    const candidates = msal.getAllAccounts().filter((a) => MoveRules.accountMatches(a, loc));
    for (const acc of candidates) {
      if (await attach(role, acc, { tenantId: loc.tenantId, upn: loc.upn })) {
        Log.info(`Reconnected ${role} tenant: ${connections[role].tenantName || connections[role].tenantHost} (${connections[role].upn})`);
        saveMap();
        return true;
      }
    }
    return false;
  }

  // Is the live connection the one this job (or record) needs?
  function connectionMatches(role, loc, { requireUpn = true } = {}) {
    const c = connections[role];
    if (!c || !c.account || !loc) return false;
    if (loc.tenantId ? c.tenantId !== loc.tenantId : (loc.tenantHost && c.tenantHost !== loc.tenantHost)) return false;
    if (requireUpn && loc.upn && !sameUpn(c.upn, loc.upn)) return false;
    return true;
  }

  // ---------- tokens ----------
  const freeAccessTokens = ToolkitAuth.freeAccessTokens;

  function interactionNeeded(e) {
    return ToolkitAuth.needsInteraction(e) || /monitor_window_timeout|timed_out/i.test(`${(e && e.errorCode) || ''} ${(e && e.message) || ''}`);
  }

  async function silentToken(req) {
    try {
      return (await msal.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      if (!ToolkitAuth.isCacheQuota(e)) throw e;
      Log.warn(`Token cache full, cleared ${freeAccessTokens()} cached access token(s).`);
      return (await msal.acquireTokenSilent(req)).accessToken;
    }
  }

  async function popupToken(req) {
    try {
      return (await msal.acquireTokenPopup(req)).accessToken;
    } catch (e) {
      if (ToolkitAuth.isPopupBlocked(e)) throw new Error('The consent popup was blocked. Allow popups for this site and try again.');
      if (ToolkitAuth.isUserCancel(e)) throw new Error('Sign-in or consent was cancelled.');
      throw e;
    }
  }

  async function getToken(role, scopes, { silent = false } = {}) {
    const conn = connections[role];
    if (!conn || !conn.account) throw new Error(`Not signed in for the ${role} tenant`);
    const key = `${role}|${conn.account.homeAccountId}|${silent ? 's' : 'i'}|${scopes.join(' ')}`;
    if (inflight.has(key)) return inflight.get(key);
    const p = (async () => {
      const req = { scopes, account: conn.account };
      try {
        return await silentToken(req);
      } catch (e) {
        if (silent || !interactionNeeded(e)) throw e;
        Log.info(`Asking for consent or sign-in (${role}): ${scopes.join(', ')}`);
        return popupToken(req);
      }
    })();
    inflight.set(key, p);
    try { return await p; } finally { inflight.delete(key); }
  }

  function getGraphToken(role) { return getToken(role, GRAPH_SCOPES); }
  // Extra Graph scopes (User.ReadBasic.All for the OneDrive picker) only when used.
  function getGraphTokenScoped(role, scopes) {
    return getToken(role, [...new Set([...GRAPH_SCOPES, ...(scopes || [])])]);
  }
  function getSpoToken(role, host) {
    const conn = connections[role];
    const h = String(host || (conn && conn.tenantHost) || '').toLowerCase();
    if (!h) throw new Error(`SharePoint host unknown for the ${role} tenant`);
    if (!conn || !MoveRules.isTenantSpoHost(h, conn.tenantHost)) {
      throw new Error(`${h} is not a SharePoint host of the ${role} tenant (${(conn && conn.tenantHost) || 'not connected'}).`);
    }
    return getToken(role, [`https://${h}/.default`]);
  }
  function getSpoAdminToken(role) {
    const conn = connections[role];
    if (!conn || !conn.adminHost) throw new Error(`SharePoint admin host unknown for the ${role} tenant`);
    return getToken(role, [`https://${conn.adminHost}/.default`]);
  }

  async function hasSharePointConsent(role) {
    const conn = connections[role];
    if (!conn || !conn.account || !conn.tenantHost) return false;
    try { await silentToken({ scopes: [`https://${conn.tenantHost}/.default`], account: conn.account }); return true; }
    catch { return false; }
  }

  // All SharePoint hosts of a tenant (sites, -my, -admin) are one resource, so one .default
  // consent covers them. .default for different resources cannot be combined in one request.
  async function grantSharePointConsent(role) {
    const conn = connections[role];
    if (!conn || !conn.account) throw new Error(`Not signed in for the ${role} tenant`);
    if (!conn.tenantHost) throw new Error(`SharePoint host unknown for the ${role} tenant`);
    if (conn._spoConsentInflight) return conn._spoConsentInflight;
    conn._spoConsentInflight = (async () => {
      const req = { scopes: [`https://${conn.tenantHost}/.default`], account: conn.account };
      try { await silentToken(req); conn.spoConsent = true; return; } catch { /* needs a popup */ }
      Log.info(`Requesting SharePoint consent for the ${role} tenant...`);
      await popupToken(req);
      conn.spoConsent = true;
      for (const h of [conn.tenantHost.replace(/\.sharepoint\.com$/i, '-my.sharepoint.com'), conn.adminHost]) {
        try { await silentToken({ scopes: [`https://${h}/.default`], account: conn.account }); }
        catch (e) { Log.dbg(`Could not pre-warm ${h} token: ${e.message}`); }
      }
    })();
    try { return await conn._spoConsentInflight; } finally { conn._spoConsentInflight = null; }
  }

  // Header chip: both connections, with tenants.
  function userInfo(onSignOut) {
    const live = ROLES.map((r) => connections[r]).filter((c) => c.account);
    if (!live.length) return null;
    const upns = [...new Set(live.map((c) => c.upn))];
    return {
      name: live[0].upn,
      upn: upns.join(' + '),
      tenant: live.map((c) => `${c.role}: ${c.tenantName || c.tenantHost || c.tenantId}`).join(', '),
      onSignOut,
    };
  }

  window.Auth = {
    init, signIn, signOutAll, restoreConnection, connectionMatches,
    getGraphToken, getGraphTokenScoped, getSpoToken, getSpoAdminToken, grantSharePointConsent, hasSharePointConsent,
    getConnection: (role) => connections[role],
    isConnected: (role) => !!(connections[role] && connections[role].account),
    getRedirectUri: () => redirectUri,
    userInfo,
  };
})();
