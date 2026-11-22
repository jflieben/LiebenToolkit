// SimpleMaester auth: the shared ToolkitAuth instance plus scope helpers the tests use.
// Read-only tool: every set below is a read scope. Sign-in asks for 'read' only; the
// other sets are asked for when a selected test needs them.
(() => {
  const G = (s) => `https://graph.microsoft.com/${s}`;
  const READ = [
    'User.Read', 'Directory.Read.All', 'Policy.Read.All', 'RoleManagement.Read.Directory',
    'Application.Read.All', 'UserAuthenticationMethod.Read.All', 'AuditLog.Read.All', 'Reports.Read.All',
  ].map(G);

  const SETS = {
    read: READ,
    forms: [G('OrgSettings-Forms.Read.All'), ...READ],
    appsAndServices: [G('OrgSettings-AppsAndServices.Read.All'), ...READ],
    securityEvents: [G('SecurityEvents.Read.All'), ...READ],
    // Azure Resource Manager, only for the Azure tests. ARM has no read-only delegated
    // scope; the tool only sends GET requests.
    arm: ['https://management.azure.com/.default'],
  };

  const isDefault = (s) => /\/\.default$/i.test(s);
  function resourceOf(scope) {
    if (/^(openid|profile|offline_access|email)$/i.test(scope)) return 'oidc';
    const m = /^https?:\/\/([^/]+)/i.exec(scope);
    return m ? m[1].toLowerCase() : 'graph.microsoft.com';
  }

  // Turn requiredScopes values (set names, scope arrays, nested arrays) into one scope list.
  function expand(req, sets = SETS) {
    if (req == null) return [];
    if (Array.isArray(req)) return [...new Set(req.flatMap((r) => expand(r, sets)))];
    const s = String(req);
    if (Object.prototype.hasOwnProperty.call(sets, s)) return sets[s].slice();
    return [/^https?:\/\//i.test(s) || resourceOf(s) === 'oidc' ? s : G(s)];
  }

  // Entra ID refuses one token request for several resources or for .default mixed
  // with named scopes, so consent is asked per resource and per .default scope.
  function consentGroups(scopes) {
    const groups = [];
    const byResource = new Map();
    for (const s of [...new Set(scopes || [])].filter(Boolean)) {
      if (isDefault(s)) { groups.push([s]); continue; }
      const key = resourceOf(s);
      if (!byResource.has(key)) byResource.set(key, []);
      byResource.get(key).push(s);
    }
    return [...byResource.values(), ...groups];
  }

  window.SMAuth = { SETS, expand, consentGroups, resourceOf };

  if (!window.ToolkitAuth) return;
  const auth = ToolkitAuth.create({ scopes: SETS });

  // Get every token the selected tests need before the scan starts, so consent
  // prompts do not interrupt it. A blocked popup falls back to a redirect, after
  // which app.js resumes the pending run.
  async function ensureScopes(required) {
    const groups = consentGroups(expand(required && (Array.isArray(required) ? required : [required])));
    for (const g of (groups.length ? groups : [SETS.read])) await auth.getToken(g);
    return true;
  }

  window.Auth = Object.assign(auth, {
    ensureScopes,
    // Arrays kept for the test modules, which build their own scope lists from these.
    SCOPES: {
      graphFull: SETS.read,
      graphForms: SETS.forms,
      graphAppsAndServices: SETS.appsAndServices,
      graphSecurityEvents: SETS.securityEvents,
    },
  });
})();
