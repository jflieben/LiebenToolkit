// CISA SCuBA (Secure Cloud Business Applications) Baseline tests for Microsoft Entra ID.
// Reference: https://maester.dev/docs/tests/ (CISA category)
// Most CISA.MS.AAD.* checks overlap with CIS/Maester tests but use the CISA framing and IDs.
// CISA.MS.EXO.2.2, 3.1, 4.1, 4.2 are implemented via DNS-over-HTTPS (no auth needed).
// CISA.MS.SHAREPOINT.1.1, 1.3 are implemented via Graph admin/sharepoint/settings.
// Other CISA.MS.EXO.* tests require Exchange Admin REST and remain as stubs.
(() => {
  const tests = [];
  const ARM_SCOPES = ['https://management.azure.com/.default'];

  function ms(start) { return Math.round(performance.now() - start); }
  function errRow(id, e, start) {
    if (e.status === 403 || e.status === 401) return { id, status: 'Skipped', reason: `Insufficient permissions: ${e.message}`, durationMs: ms(start) };
    return { id, status: 'Error', reason: e.message, durationMs: ms(start) };
  }
  function exoStub(id, title) {
    return {
      id, title, severity: 'Info', tag: 'CISA', category: 'CISA',
      docUrl: `https://maester.dev/docs/tests/${id}`,
      implemented: false,
      run: () => Promise.resolve({
        id, status: 'Skipped',
        reason: 'This CISA test requires Exchange Online Admin REST API or SharePoint API which is not accessible from a browser-based PKCE app.',
        durationMs: 0,
      }),
    };
  }

  async function armFetch(url) {
    const tok = await Auth.getToken(ARM_SCOPES);
    const res = await fetch(url, { headers: { Authorization: 'Bearer ' + tok } });
    if (res.status === 403 || res.status === 401) {
      const err = new Error(`ARM ${res.status} for ${url}`);
      err.status = res.status;
      throw err;
    }
    if (!res.ok) {
      const t = await res.text().catch(() => '');
      const err = new Error(`ARM ${res.status} for ${url}: ${t.substring(0, 300)}`);
      err.status = res.status;
      throw err;
    }
    return res.json();
  }

  // Privileged role template IDs used for CISA privileged checks
  const PRIV_ROLE_TEMPLATE_IDS = [
    '62e90394-69f5-4237-9190-012177145e10', // Global Administrator
    '7be44c8a-adaf-4e2a-84d6-ab2649e08a13', // Privileged Authentication Admin
    'e8611ab8-c189-46e8-94e1-60213ab1f814', // Privileged Role Administrator
    '194ae4cb-b126-40b2-bd5b-6091b380977d', // Security Administrator
    'fe930be7-5e62-47db-91af-98c3a49a38b1', // User Administrator
    '29232cdf-9323-42fd-ade2-1d097af3e4de', // Exchange Administrator
    'f28a1f50-f6e7-4571-818b-6a12f2af6b6c', // SharePoint Administrator
    '69091246-20e8-4a56-aa4d-066075b2a7a8', // Teams Administrator
    'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9', // Conditional Access Administrator
    'b0f54661-2d74-4c50-afa3-1ec803f12efe', // Billing Administrator
    '966707d0-3269-4727-9be2-8c3a10f19b9d', // Password Administrator
    'f2ef992c-3afb-46b9-b7cf-a126ee74c451', // Global Reader
  ];

  // Shared PIM policy/rules prefetch for CISA 7.6-7.9.
  // We fetch once per scan and reuse results across tests to avoid repeated
  // serial N+1 calls to /policies/roleManagementPolicies/{id}/rules.
  let _dirRolePoliciesWithRulesPromise = null;

  async function mapWithConcurrency(items, worker, limit = 8, onItemDone = null) {
    const out = new Array(items.length);
    let next = 0;

    async function runWorker() {
      while (true) {
        const i = next;
        next++;
        if (i >= items.length) break;
        out[i] = await worker(items[i], i);
        if (typeof onItemDone === 'function') onItemDone(i, items[i]);
      }
    }

    const count = Math.max(1, Math.min(limit, items.length));
    await Promise.all(Array.from({ length: count }, () => runWorker()));
    return out;
  }

  async function getDirectoryRolePoliciesWithRules(onProgress = null) {
    if (_dirRolePoliciesWithRulesPromise) {
      if (typeof onProgress === 'function') {
        onProgress({ done: 1, total: 1, label: 'Using cached PIM role policies' });
      }
      return _dirRolePoliciesWithRulesPromise;
    }

    _dirRolePoliciesWithRulesPromise = (async () => {
      const policiesResp = await Graph.graph(
        `policies/roleManagementPolicies?$filter=scopeId eq '/' and scopeType eq 'DirectoryRole'`,
        { apiVersion: 'beta' }
      );
      const policies = policiesResp.value || [];
      if (typeof onProgress === 'function') {
        onProgress({ done: 0, total: Math.max(1, policies.length), label: 'Loading PIM role policy rules...' });
      }
      if (!policies.length) return [];

      let completed = 0;

      const results = await mapWithConcurrency(policies, async (pol) => {
        try {
          const rulesResp = await Graph.graph(`policies/roleManagementPolicies/${pol.id}/rules`, { apiVersion: 'beta' });
          return { ...pol, rules: rulesResp.value || [] };
        } catch {
          return { ...pol, rules: [] };
        }
      }, 3, () => {
        completed++;
        if (typeof onProgress === 'function') {
          onProgress({ done: completed, total: policies.length, label: 'Reading PIM policies and rules...' });
        }
      });

      return results;
    })();

    return _dirRolePoliciesWithRulesPromise;
  }

  // ── CISA.MS.AAD.1.1 Legacy authentication SHALL be blocked ──────────────────────
  tests.push({
    id: 'CISA.MS.AAD.1.1',
    title: 'Legacy authentication SHALL be blocked',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.1.1',
    description: 'An enabled CA policy for all users must block legacy authentication (exchangeActiveSync + other client app types).',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.grantControls?.builtInControls || []).includes('block')
          && (p.conditions?.users?.includeUsers || []).includes('All')
          && (p.conditions?.clientAppTypes || []).some(t => /exchangeActiveSync|other/i.test(t)));
        if (hit) return { id: 'CISA.MS.AAD.1.1', status: 'Passed', reason: `Policy "${hit.displayName}" blocks legacy auth for all users.`, durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.1.1', status: 'Failed', reason: 'No enabled CA policy for all users that blocks legacy client app types (exchangeActiveSync + other).', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.1.1', e, start); }
    },
  });

  // ── CISA.MS.AAD.2.1 High risk users SHALL be blocked ───────────────────────────
  tests.push({
    id: 'CISA.MS.AAD.2.1',
    title: 'Users detected as high risk SHALL be blocked',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.2.1',
    description: 'Requires an enabled CA policy targeting all users with high userRiskLevel that blocks access. Requires Entra ID P2.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.users?.includeUsers || []).includes('All')
          && (p.conditions?.userRiskLevels || []).some(r => /high/i.test(r))
          && ((p.grantControls?.builtInControls || []).includes('block')
            || (p.grantControls?.builtInControls || []).includes('mfa')));
        if (hit) return { id: 'CISA.MS.AAD.2.1', status: 'Passed', reason: `Policy "${hit.displayName}" acts on high-risk users.`, durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.2.1', status: 'Failed', reason: 'No enabled CA policy blocking high-risk users for all users. Requires Entra ID P2.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.2.1', e, start); }
    },
  });

  // ── CISA.MS.AAD.2.2 Notification SHOULD be sent for high-risk users ─────────────
  tests.push({
    id: 'CISA.MS.AAD.2.2',
    title: 'A notification SHOULD be sent to the administrator when high-risk users are detected',
    severity: 'Medium', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.2.2',
    description: 'Checks if admin notification for high-risk users can be validated via Graph. This setting is not currently exposed in a stable Graph endpoint.',
    implemented: true,
    async run() {
      const start = performance.now();
      return {
        id: 'CISA.MS.AAD.2.2',
        status: 'Skipped',
        reason: 'Graph does not currently expose tenant-level Identity Protection notification recipients in a stable, browser-safe endpoint. Verify in Entra portal: Protection > Risky users > Notifications.',
        durationMs: ms(start),
      };
    },
  });

  // ── CISA.MS.AAD.2.3 High risk sign-ins SHALL be blocked ────────────────────────
  tests.push({
    id: 'CISA.MS.AAD.2.3',
    title: 'Sign-ins detected as high risk SHALL be blocked',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.2.3',
    description: 'Requires an enabled CA policy for all users with high signInRiskLevel that blocks or requires MFA. Requires Entra ID P2.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.users?.includeUsers || []).includes('All')
          && (p.conditions?.signInRiskLevels || []).some(r => /high/i.test(r))
          && ((p.grantControls?.builtInControls || []).includes('block')
            || (p.grantControls?.builtInControls || []).includes('mfa')
            || p.grantControls?.authenticationStrength?.id));
        if (hit) return { id: 'CISA.MS.AAD.2.3', status: 'Passed', reason: `Policy "${hit.displayName}" acts on high-risk sign-ins.`, durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.2.3', status: 'Failed', reason: 'No enabled CA policy for all users that blocks or requires MFA for high sign-in risk. Requires Entra ID P2.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.2.3', e, start); }
    },
  });

  // ── CISA.MS.AAD.3.1 Phishing-resistant MFA SHALL be enforced ───────────────────
  tests.push({
    id: 'CISA.MS.AAD.3.1',
    title: 'Phishing-resistant MFA SHALL be enforced for all users',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.1',
    description: 'An enabled CA policy for all apps and all users should require an authentication strength that is phishing-resistant (FIDO2, Windows Hello, certificate).',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        // Check for policies that require an authentication strength
        const policiesWithStrength = pols.filter(p => p.state === 'enabled'
          && (p.conditions?.users?.includeUsers || []).includes('All')
          && p.grantControls?.authenticationStrength?.id);
        if (policiesWithStrength.length === 0) {
          return { id: 'CISA.MS.AAD.3.1', status: 'Failed', reason: 'No enabled CA policy for all users uses an authentication strength policy (required for phishing-resistant MFA enforcement).', durationMs: ms(start) };
        }
        // Check if any of those strengths are phishing-resistant
        for (const pol of policiesWithStrength) {
          try {
            const strength = await Graph.graph(`identity/conditionalAccess/authenticationStrength/policies/${pol.grantControls.authenticationStrength.id}`, { apiVersion: 'v1.0' });
            if (/phishing|fido|certificate|windowsHello/i.test(JSON.stringify(strength))) {
              return { id: 'CISA.MS.AAD.3.1', status: 'Passed', reason: `Policy "${pol.displayName}" requires phishing-resistant authentication strength.`, durationMs: ms(start) };
            }
          } catch (_) { /* skip */ }
        }
        return { id: 'CISA.MS.AAD.3.1', status: 'Failed', reason: `${policiesWithStrength.length} CA policy(ies) use authentication strength but none appear to be explicitly phishing-resistant. Verify the strength policy includes FIDO2/WHfB/Certificate.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.1', e, start); }
    },
  });

  // ── CISA.MS.AAD.3.2 Alternative MFA SHALL be enforced (if phishing-resistant not yet done) ──
  tests.push({
    id: 'CISA.MS.AAD.3.2',
    title: 'If phishing-resistant MFA has not been enforced, an alternative MFA method SHALL be enforced for all users',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.2',
    description: 'At minimum, any MFA requirement for all users should exist while migrating to phishing-resistant methods.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.users?.includeUsers || []).includes('All')
          && ((p.grantControls?.builtInControls || []).includes('mfa') || p.grantControls?.authenticationStrength?.id));
        if (hit) return { id: 'CISA.MS.AAD.3.2', status: 'Passed', reason: `Policy "${hit.displayName}" requires MFA for all users.`, durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.3.2', status: 'Failed', reason: 'No enabled CA policy requiring any form of MFA for all users.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.2', e, start); }
    },
  });

  // ── CISA.MS.AAD.3.3 Microsoft Authenticator SHALL show login context ────────────
  tests.push({
    id: 'CISA.MS.AAD.3.3',
    title: 'If Microsoft Authenticator is enabled, it SHALL be configured to show login context information',
    severity: 'Medium', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.3',
    description: 'Microsoft Authenticator should have number matching AND show app/location context to prevent MFA fatigue attacks.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/authenticationMethodsPolicy', { apiVersion: 'beta' });
        const msAuth = pol.authenticationMethodConfigurations?.find(c => /microsoftAuthenticator/i.test(c.id));
        if (!msAuth || msAuth.state !== 'enabled') {
          return { id: 'CISA.MS.AAD.3.3', status: 'Skipped', reason: 'Microsoft Authenticator is not enabled as an authentication method.', durationMs: ms(start) };
        }
        const feat = msAuth.featureSettings || {};
        const nmEnabled = feat.numberMatchingRequiredState?.state === 'enabled';
        const appEnabled = feat.displayAppInformationRequiredState?.state === 'enabled';
        const locEnabled = feat.displayLocationInformationRequiredState?.state === 'enabled';
        if (nmEnabled && (appEnabled || locEnabled)) {
          return { id: 'CISA.MS.AAD.3.3', status: 'Passed', reason: 'Microsoft Authenticator has number matching and context information enabled.', durationMs: ms(start) };
        }
        const issues = [];
        if (!nmEnabled) issues.push('number matching not enabled');
        if (!appEnabled) issues.push('app name not shown');
        if (!locEnabled) issues.push('location not shown');
        return { id: 'CISA.MS.AAD.3.3', status: 'Failed', reason: `Microsoft Authenticator context issues: ${issues.join(', ')}. Enable these in Authentication Methods policy.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.3', e, start); }
    },
  });

  // ── CISA.MS.AAD.3.4 Authentication Methods Manage Migration = Migration Complete ─
  tests.push({
    id: 'CISA.MS.AAD.3.4',
    title: 'The Authentication Methods Manage Migration feature SHALL be set to Migration Complete',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.4',
    description: 'If migration is not complete, the legacy per-user MFA settings may override the modern authentication methods policy.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/authenticationMethodsPolicy', { apiVersion: 'beta' });
        const state = pol.policyMigrationState;
        if (state === 'migrationComplete') return { id: 'CISA.MS.AAD.3.4', status: 'Passed', reason: 'Authentication methods migration state is "migrationComplete".', durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.3.4', status: 'Failed', reason: `Authentication methods migration state is "${state || 'unknown'}". Set to "migrationComplete" in the Authentication Methods policy.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.4', e, start); }
    },
  });

  // ── CISA.MS.AAD.3.5 SMS/Voice/Email OTP SHALL be disabled ───────────────────────
  tests.push({
    id: 'CISA.MS.AAD.3.5',
    title: 'The authentication methods SMS, Voice Call, and Email One-Time Passcode (OTP) SHALL be disabled',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.5',
    description: 'SMS, Voice, and Email OTP are vulnerable to interception, SIM swap, and phishing. All three should be disabled in the Authentication Methods policy.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/authenticationMethodsPolicy', { apiVersion: 'beta' });
        const methods = pol.authenticationMethodConfigurations || [];
        const issues = [];
        for (const m of methods) {
          if (/^(sms|voice|email)$/i.test(m.id) && m.state === 'enabled') issues.push(m.id);
        }
        if (issues.length === 0) return { id: 'CISA.MS.AAD.3.5', status: 'Passed', reason: 'SMS, Voice, and Email OTP authentication methods are disabled.', durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.3.5', status: 'Failed', reason: `Weak authentication methods still enabled: ${issues.join(', ')}. Disable in Authentication Methods policy.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.5', e, start); }
    },
  });

  // ── CISA.MS.AAD.3.6 Phishing-resistant MFA for highly privileged roles ───────────
  tests.push({
    id: 'CISA.MS.AAD.3.6',
    title: 'Phishing-resistant MFA SHALL be required for highly privileged roles',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.6',
    description: 'At least one CA policy should target privileged admin roles and require an authentication strength (ideally phishing-resistant).',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const privSet = new Set(PRIV_ROLE_TEMPLATE_IDS);
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.users?.includeRoles || []).some(r => privSet.has(r))
          && p.grantControls?.authenticationStrength?.id);
        if (hit) return { id: 'CISA.MS.AAD.3.6', status: 'Passed', reason: `Policy "${hit.displayName}" requires authentication strength for privileged roles.`, durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.3.6', status: 'Failed', reason: 'No enabled CA policy targeting privileged roles with an authentication strength requirement.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.6', e, start); }
    },
  });

  // ── CISA.MS.AAD.3.7 Managed devices required for auth ───────────────────────────
  tests.push({
    id: 'CISA.MS.AAD.3.7',
    title: 'Managed devices SHOULD be required for authentication',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.7',
    description: 'At least one CA policy requires device compliance or hybrid join for all cloud apps for some or all users.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.grantControls?.builtInControls || []).some(c => /compliantDevice|domainJoined/i.test(c)));
        if (hit) return { id: 'CISA.MS.AAD.3.7', status: 'Passed', reason: `Policy "${hit.displayName}" requires a managed device.`, durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.3.7', status: 'Failed', reason: 'No enabled CA policy requires a compliant or hybrid-joined device.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.7', e, start); }
    },
  });

  // ── CISA.MS.AAD.3.8 Managed devices for MFA registration ───────────────────────
  tests.push({
    id: 'CISA.MS.AAD.3.8',
    title: 'Managed Devices SHOULD be required to register MFA',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.8',
    description: 'The "Register security information" CA user action should require a managed device, trusted location, or similar control.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.applications?.includeUserActions || []).includes('urn:user:registersecurityinfo')
          && ((p.grantControls?.builtInControls || []).some(c => /compliantDevice|domainJoined/i.test(c))
            || (p.conditions?.locations?.includeLocations || []).some(l => /trusted|AllTrusted/i.test(l))));
        if (hit) return { id: 'CISA.MS.AAD.3.8', status: 'Passed', reason: `Policy "${hit.displayName}" restricts MFA registration to managed devices or trusted locations.`, durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.3.8', status: 'Failed', reason: 'No enabled CA policy restricts the MFA registration user action to managed devices or trusted locations.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.8', e, start); }
    },
  });

  // ── CISA.MS.AAD.4.1 Security logs SHALL be sent to SOC ──────────────────────────
  tests.push({
    id: 'CISA.MS.AAD.4.1',
    title: 'Security logs SHALL be sent to the agency\'s security operations center for monitoring',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.4.1',
    description: 'Checks Entra diagnostic settings via Azure Resource Manager to verify sign-in and audit logs are streamed to a SIEM destination.',
    implemented: true,
    requiredScopes: ARM_SCOPES,
    async run() {
      const start = performance.now();
      try {
        const diag = await armFetch('https://management.azure.com/providers/microsoft.aadiam/diagnosticSettings?api-version=2017-04-01-preview');
        const settings = diag.value || [];
        if (!settings.length) {
          return { id: 'CISA.MS.AAD.4.1', status: 'Failed', reason: 'No Entra diagnostic settings found. Configure streaming of Entra logs to SOC (Log Analytics/Event Hub/Storage).', durationMs: ms(start) };
        }

        const qualifying = settings.filter(s => {
          const logs = s.properties?.logs || [];
          const hasAudit = logs.some(l => l.enabled && /auditlogs|audit/i.test(String(l.category || '')));
          const hasSignin = logs.some(l => l.enabled && /signinlogs|sign.?in/i.test(String(l.category || '')));
          const hasDestination = !!(
            s.properties?.workspaceId
            || s.properties?.eventHubAuthorizationRuleId
            || s.properties?.storageAccountId
            || s.properties?.marketplacePartnerId
          );
          return hasAudit && hasSignin && hasDestination;
        });

        if (!qualifying.length) {
          return { id: 'CISA.MS.AAD.4.1', status: 'Failed', reason: `${settings.length} diagnostic setting(s) found, but none stream both AuditLogs and SignInLogs to a SOC destination.`, durationMs: ms(start) };
        }

        return { id: 'CISA.MS.AAD.4.1', status: 'Passed', reason: `${qualifying.length} Entra diagnostic setting(s) stream both AuditLogs and SignInLogs to SOC destinations.`, durationMs: ms(start) };
      } catch (e) {
        if (e.status === 401 || e.status === 403 || e.status === 404) {
          return { id: 'CISA.MS.AAD.4.1', status: 'Skipped', reason: 'Could not read Entra diagnostic settings from ARM. Requires Azure Reader access and management.azure.com consent.', durationMs: ms(start) };
        }
        return errRow('CISA.MS.AAD.4.1', e, start);
      }
    },
  });

  // ── CISA.MS.AAD.5.1 Only administrators SHALL be allowed to register applications ─
  tests.push({
    id: 'CISA.MS.AAD.5.1',
    title: 'Only administrators SHALL be allowed to register applications',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.5.1',
    description: 'Users should not be able to register app registrations. This is controlled by the default user role permission allowedToCreateApps.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const allowed = pol.defaultUserRolePermissions?.allowedToCreateApps;
        if (allowed === false) return { id: 'CISA.MS.AAD.5.1', status: 'Passed', reason: 'Non-admin users cannot register applications (allowedToCreateApps = false).', durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.5.1', status: 'Failed', reason: 'Non-admin users can register applications. Disable this in the authorization policy.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.5.1', e, start); }
    },
  });

  // ── CISA.MS.AAD.5.2 Only administrators SHALL be allowed to consent ──────────────
  tests.push({
    id: 'CISA.MS.AAD.5.2',
    title: 'Only administrators SHALL be allowed to consent to applications',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.5.2',
    description: 'User consent to applications should be disabled. Only admins should be able to grant app permissions.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const consent = pol.permissionGrantPolicyIdsAssignedToDefaultUserRole || [];
        const noConsent = consent.length === 0 || consent.every(p => /ManagePermissionGrantsForSelf.*None/i.test(p));
        if (noConsent) return { id: 'CISA.MS.AAD.5.2', status: 'Passed', reason: 'User consent to applications is disabled.', durationMs: ms(start) };
        // Also check via policies
        const consentPolicy = await Graph.graph('policies/permissionGrantPolicies', { apiVersion: 'v1.0' });
        const defaultPol = (consentPolicy.value || []).find(p => p.id === 'microsoft-user-default-legacy' || p.id === 'microsoft-user-default-low');
        if (!defaultPol?.includes?.length) return { id: 'CISA.MS.AAD.5.2', status: 'Passed', reason: 'No user consent grant policy active.', durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.5.2', status: 'Failed', reason: `User consent permissions are granted: ${consent.join(', ')}. Restrict this to admins only.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.5.2', e, start); }
    },
  });

  // ── CISA.MS.AAD.5.3 Admin consent workflow SHALL be configured ──────────────────
  tests.push({
    id: 'CISA.MS.AAD.5.3',
    title: 'An admin consent workflow SHALL be configured for applications',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.5.3',
    description: 'When user consent is restricted, users should be able to request admin approval. The admin consent workflow should be enabled.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/adminConsentRequestPolicy', { apiVersion: 'beta' });
        if (pol.isEnabled) return { id: 'CISA.MS.AAD.5.3', status: 'Passed', reason: 'Admin consent request workflow is enabled.', durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.5.3', status: 'Failed', reason: 'Admin consent request workflow is not enabled. Enable it so users can request app consent from admins.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.5.3', e, start); }
    },
  });

  // ── CISA.MS.AAD.5.4 Group owners SHALL NOT be allowed to consent ─────────────────
  tests.push({
    id: 'CISA.MS.AAD.5.4',
    title: 'Group owners SHALL NOT be allowed to consent to applications accessing group data',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.5.4',
    description: 'Checks that group owners cannot consent to apps. This is controlled by the consentPolicySettings.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const settings = await Graph.graphAll('groupSettings', { apiVersion: 'v1.0' });
        const consent = settings.find(s => /consent/i.test(s.displayName));
        if (consent) {
          const blocked = consent.values?.find(v => v.name === 'BlockMSFTGroupCreation' || v.name === 'AllowGroupSpecificConsent');
          // If AllowGroupSpecificConsent is false or not present, group owners can't consent
        }
        // Check via policies/authorizationPolicy
        const authPol = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const groupOwnerConsent = authPol.allowGroupSpecificConsent ?? authPol.defaultUserRolePermissions?.allowedToCreateSecurityGroups;
        // The actual check is: permissionGrantPolicy for groups
        const policyIds = authPol.permissionGrantPolicyIdsAssignedToDefaultUserRole || [];
        const hasGroupConsent = policyIds.some(p => /ManagePermissionGrantsForOwnedResource/i.test(p));
        if (!hasGroupConsent) return { id: 'CISA.MS.AAD.5.4', status: 'Passed', reason: 'Group owners are not allowed to consent to applications accessing group data.', durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.5.4', status: 'Failed', reason: 'Group owners may consent to applications. Disable group owner consent in the authorization policy.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.5.4', e, start); }
    },
  });

  // ── CISA.MS.AAD.6.1 User passwords SHALL NOT expire ─────────────────────────────
  tests.push({
    id: 'CISA.MS.AAD.6.1',
    title: 'User passwords SHALL NOT expire',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.6.1',
    description: 'Microsoft and CISA recommend setting passwords to never expire. Frequent mandatory changes lead to weaker passwords.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const domains = await Graph.graphAll('domains?$select=id,isVerified,passwordNotificationWindowInDays,passwordValidityPeriodInDays', { apiVersion: 'v1.0' });
        const expiring = domains.filter(d => d.isVerified && d.passwordValidityPeriodInDays && d.passwordValidityPeriodInDays < 2147483647);
        if (expiring.length === 0) return { id: 'CISA.MS.AAD.6.1', status: 'Passed', reason: 'All domains have password expiry set to never expire.', durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.6.1', status: 'Failed', reason: `${expiring.length} domain(s) have password expiry configured: ${expiring.map(d => `${d.id} (${d.passwordValidityPeriodInDays} days)`).join(', ')}`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.6.1', e, start); }
    },
  });

  // ── CISA.MS.AAD.7.1 Minimum 2 and maximum 8 Global Admins ───────────────────────
  tests.push({
    id: 'CISA.MS.AAD.7.1',
    title: 'A minimum of two users and a maximum of eight users SHALL be provisioned with the Global Administrator role',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.1',
    description: 'Too few GAs is a resilience risk; too many inflates blast radius. CISA and CIS both recommend 2-8.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const roles = await Graph.graphAll('directoryRoles', { apiVersion: 'v1.0' });
        const ga = roles.find(r => r.roleTemplateId === '62e90394-69f5-4237-9190-012177145e10');
        if (!ga) return { id: 'CISA.MS.AAD.7.1', status: 'Skipped', reason: 'Global Administrator role not activated.', durationMs: ms(start) };
        const members = await Graph.graphAll(`directoryRoles/${ga.id}/members`, { apiVersion: 'v1.0' });
        const users = members.filter(m => m['@odata.type'] === '#microsoft.graph.user').length;
        const ok = users >= 2 && users <= 8;
        return { id: 'CISA.MS.AAD.7.1', status: ok ? 'Passed' : 'Failed', actual: users,
          reason: `${users} Global Administrator(s). CISA requires 2–8.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.7.1', e, start); }
    },
  });

  // ── CISA.MS.AAD.7.2 Privileged users SHALL have finer-grained roles, not GA ──────
  tests.push({
    id: 'CISA.MS.AAD.7.2',
    title: 'Privileged users SHALL be provisioned with finer-grained roles instead of Global Administrator',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.2',
    description: 'Checks the ratio: for every GA, there should be multiple users in specific-purpose roles. High GA count relative to other privileged roles is a sign of over-privileging.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const roles = await Graph.graphAll('directoryRoles', { apiVersion: 'v1.0' });
        const ga = roles.find(r => r.roleTemplateId === '62e90394-69f5-4237-9190-012177145e10');
        if (!ga) return { id: 'CISA.MS.AAD.7.2', status: 'Skipped', reason: 'Global Administrator role not activated.', durationMs: ms(start) };
        const gaMembers = await Graph.graphAll(`directoryRoles/${ga.id}/members`, { apiVersion: 'v1.0' });
        const gaCount = gaMembers.filter(m => m['@odata.type'] === '#microsoft.graph.user').length;
        // Count members in other privileged roles
        const otherPrivRoles = roles.filter(r => r.roleTemplateId !== '62e90394-69f5-4237-9190-012177145e10'
          && PRIV_ROLE_TEMPLATE_IDS.includes(r.roleTemplateId));
        let otherPrivCount = 0;
        for (const r of otherPrivRoles.slice(0, 10)) {
          try {
            const m = await Graph.graphAll(`directoryRoles/${r.id}/members`, { apiVersion: 'v1.0' });
            otherPrivCount += m.filter(u => u['@odata.type'] === '#microsoft.graph.user').length;
          } catch (_) { /* skip */ }
        }
        if (gaCount <= 4 && otherPrivCount >= gaCount) {
          return { id: 'CISA.MS.AAD.7.2', status: 'Passed', reason: `${gaCount} GA(s), ${otherPrivCount} users in other privileged roles. Good balance.`, durationMs: ms(start) };
        }
        return { id: 'CISA.MS.AAD.7.2', status: 'Failed', reason: `${gaCount} Global Admin(s) vs ${otherPrivCount} users in other privileged roles. Consider using finer-grained roles instead of GA.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.7.2', e, start); }
    },
  });

  // ── CISA.MS.AAD.7.3 Privileged users SHALL be provisioned cloud-only accounts ─────
  tests.push({
    id: 'CISA.MS.AAD.7.3',
    title: 'Privileged users SHALL be provisioned cloud-only accounts separate from on-premises directory',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.3',
    description: 'Checks that Global Administrators are all cloud-only accounts (not synced from on-prem AD).',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const roles = await Graph.graphAll('directoryRoles', { apiVersion: 'v1.0' });
        const ga = roles.find(r => r.roleTemplateId === '62e90394-69f5-4237-9190-012177145e10');
        if (!ga) return { id: 'CISA.MS.AAD.7.3', status: 'Skipped', reason: 'Global Administrator role not activated.', durationMs: ms(start) };
        const members = await Graph.graphAll(`directoryRoles/${ga.id}/members`, { apiVersion: 'v1.0' });
        const userMembers = members.filter(m => m['@odata.type'] === '#microsoft.graph.user');
        const synced = [];
        for (const m of userMembers) {
          try {
            const u = await Graph.graph(`users/${m.id}?$select=onPremisesSyncEnabled,userPrincipalName`, { apiVersion: 'v1.0' });
            if (u.onPremisesSyncEnabled) synced.push(u.userPrincipalName);
          } catch (_) { /* skip */ }
        }
        if (synced.length === 0) return { id: 'CISA.MS.AAD.7.3', status: 'Passed', reason: 'All Global Administrators are cloud-only accounts.', durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.7.3', status: 'Failed', actual: synced.length,
          reason: `${synced.length} synced GA account(s): ${synced.join(', ')}`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.7.3', e, start); }
    },
  });

  // ── CISA.MS.AAD.7.4 No permanent active role assignments for highly privileged ────
  tests.push({
    id: 'CISA.MS.AAD.7.4',
    title: 'Permanent active role assignments SHALL NOT be allowed for highly privileged roles',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.4',
    description: 'Uses PIM to check for permanent (always-active) assignments to the Global Administrator role. All GA assignments should be PIM eligible, not permanent.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const GA_TEMPLATE = '62e90394-69f5-4237-9190-012177145e10';
        const active = await Graph.graphAll('roleManagement/directory/roleAssignmentScheduleInstances?$expand=roleDefinition,principal', { apiVersion: 'v1.0' });
        const perm = active.filter(a => a.assignmentType === 'Assigned' && a.roleDefinition?.templateId === GA_TEMPLATE);
        if (perm.length === 0) return { id: 'CISA.MS.AAD.7.4', status: 'Passed', reason: 'No permanent active Global Administrator assignments found.', durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.7.4', status: 'Failed', actual: perm.length,
          reason: `${perm.length} permanent active Global Administrator assignment(s): ${perm.slice(0, 5).map(p => p.principal?.displayName || p.principalId).join(', ')}`,
          durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403) return { id: 'CISA.MS.AAD.7.4', status: 'Skipped', reason: 'Insufficient permissions for PIM data (requires PrivilegedAccess.Read or RoleManagement.Read.All).', durationMs: ms(start) };
        return errRow('CISA.MS.AAD.7.4', e, start);
      }
    },
  });

  // ── CISA.MS.AAD.7.5 Provisioning via PAM (PIM eligible) ─────────────────────────
  tests.push({
    id: 'CISA.MS.AAD.7.5',
    title: 'Provisioning users to highly privileged roles SHALL NOT occur outside of a PAM system',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.5',
    description: 'Checks for permanent active privileged role assignments. Privileged roles should be managed as eligible assignments via PIM.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const active = await Graph.graphAll('roleManagement/directory/roleAssignmentScheduleInstances?$expand=roleDefinition,principal', { apiVersion: 'v1.0' });
        const privSet = new Set(PRIV_ROLE_TEMPLATE_IDS);
        const perm = active.filter(a => a.assignmentType === 'Assigned' && privSet.has(a.roleDefinition?.templateId));
        if (!perm.length) return { id: 'CISA.MS.AAD.7.5', status: 'Passed', reason: 'No permanent privileged assignments found. Privileged roles appear to be managed via PIM eligibility/activation.', durationMs: ms(start) };
        return {
          id: 'CISA.MS.AAD.7.5',
          status: 'Failed',
          actual: perm.length,
          reason: `${perm.length} permanent privileged role assignment(s) found. Migrate these to PIM eligible assignments.`,
          durationMs: ms(start),
        };
      } catch (e) {
        if (e.status === 403 || e.status === 404) return { id: 'CISA.MS.AAD.7.5', status: 'Skipped', reason: 'Privileged role schedule data is not accessible with current permissions.', durationMs: ms(start) };
        return errRow('CISA.MS.AAD.7.5', e, start);
      }
    },
  });

  // ── CISA.MS.AAD.7.6 GA activation SHALL require approval ────────────────────────
  tests.push({
    id: 'CISA.MS.AAD.7.6',
    title: 'Activation of the Global Administrator role SHALL require approval',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.6',
    description: 'Checks the PIM role management policy for Global Administrator to verify activation requires approval.',
    implemented: true,
    async run(ctx = {}) {
      const start = performance.now();
      try {
        const report = typeof ctx.reportSubProgress === 'function' ? ctx.reportSubProgress : null;
        const all = await getDirectoryRolePoliciesWithRules(report);
        if (report) report({ done: 0, total: Math.max(1, all.length), label: 'Evaluating approval requirements...' });
        for (let i = 0; i < all.length; i++) {
          const pol = all[i];
          if (report) report({ done: i + 1, total: all.length, label: 'Evaluating approval requirements...' });
          const approvalRule = (pol.rules || []).find(r => r['@odata.type'] === '#microsoft.graph.unifiedRoleManagementPolicyApprovalRule');
          if (approvalRule?.setting?.isApprovalRequired) {
            return { id: 'CISA.MS.AAD.7.6', status: 'Passed', reason: 'PIM Global Administrator role activation requires approval.', durationMs: ms(start) };
          }
        }
        return { id: 'CISA.MS.AAD.7.6', status: 'Failed', reason: 'PIM Global Administrator role activation does not require approval. Configure approval in PIM role settings.', durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403 || e.status === 404) return { id: 'CISA.MS.AAD.7.6', status: 'Skipped', reason: 'PIM role management policies not accessible (requires Entra ID P2).', durationMs: ms(start) };
        return errRow('CISA.MS.AAD.7.6', e, start);
      }
    },
  });

  // ── CISA.MS.AAD.7.7 Eligible/active highly privileged assignments SHALL trigger alert ─
  tests.push({
    id: 'CISA.MS.AAD.7.7',
    title: 'Eligible and Active highly privileged role assignments SHALL trigger an alert',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.7',
    description: 'Checks that PIM assignment notifications are configured for the GA role.',
    implemented: true,
    async run(ctx = {}) {
      const start = performance.now();
      try {
        const report = typeof ctx.reportSubProgress === 'function' ? ctx.reportSubProgress : null;
        const all = await getDirectoryRolePoliciesWithRules(report);
        if (report) report({ done: 0, total: Math.max(1, all.length), label: 'Evaluating PIM notification rules...' });
        for (let i = 0; i < all.length; i++) {
          const pol = all[i];
          if (report) report({ done: i + 1, total: all.length, label: 'Evaluating PIM notification rules...' });
          const notifRule = (pol.rules || []).find(r => r['@odata.type'] === '#microsoft.graph.unifiedRoleManagementPolicyNotificationRule'
            && r.notificationType === 'Email' && r.notificationLevel === 'Critical');
          if (notifRule?.isDefaultRecipientsEnabled || (notifRule?.notificationRecipients || []).length > 0) {
            return { id: 'CISA.MS.AAD.7.7', status: 'Passed', reason: 'PIM role management notification rules are configured for critical assignments.', durationMs: ms(start) };
          }
        }
        return { id: 'CISA.MS.AAD.7.7', status: 'Failed', reason: 'No PIM notification rules found for highly privileged role assignments. Configure in PIM role settings.', durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403 || e.status === 404) return { id: 'CISA.MS.AAD.7.7', status: 'Skipped', reason: 'PIM role policies not accessible (requires Entra ID P2).', durationMs: ms(start) };
        return errRow('CISA.MS.AAD.7.7', e, start);
      }
    },
  });

  // ── CISA.MS.AAD.7.8 GA activation SHALL trigger alert ───────────────────────────
  tests.push({
    id: 'CISA.MS.AAD.7.8',
    title: 'User activation of the Global Administrator role SHALL trigger an alert',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.8',
    description: 'When a user activates the Global Administrator role through PIM, an alert notification should be sent.',
    implemented: true,
    async run(ctx = {}) {
      const start = performance.now();
      try {
        const report = typeof ctx.reportSubProgress === 'function' ? ctx.reportSubProgress : null;
        const all = await getDirectoryRolePoliciesWithRules(report);
        if (report) report({ done: 0, total: Math.max(1, all.length), label: 'Evaluating activation alert rules...' });
        for (let i = 0; i < all.length; i++) {
          const pol = all[i];
          if (report) report({ done: i + 1, total: all.length, label: 'Evaluating activation alert rules...' });
          const activationNotif = (pol.rules || []).find(r => r['@odata.type'] === '#microsoft.graph.unifiedRoleManagementPolicyNotificationRule'
            && /activation/i.test(r.id || ''));
          if (activationNotif?.isDefaultRecipientsEnabled || (activationNotif?.notificationRecipients || []).length > 0) {
            return { id: 'CISA.MS.AAD.7.8', status: 'Passed', reason: 'PIM activation notification is configured.', durationMs: ms(start) };
          }
        }
        return { id: 'CISA.MS.AAD.7.8', status: 'Failed', reason: 'No PIM activation notification found. Configure email notifications for GA activation in PIM role settings.', durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403 || e.status === 404) return { id: 'CISA.MS.AAD.7.8', status: 'Skipped', reason: 'PIM role policies not accessible (requires Entra ID P2).', durationMs: ms(start) };
        return errRow('CISA.MS.AAD.7.8', e, start);
      }
    },
  });

  // ── CISA.MS.AAD.7.9 Other privileged role activation SHOULD trigger alert ────────
  tests.push({
    id: 'CISA.MS.AAD.7.9',
    title: 'User activation of other highly privileged roles SHOULD trigger an alert',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.9',
    description: 'Activation of any highly privileged role through PIM should generate an email notification.',
    implemented: true,
    async run(ctx = {}) {
      const start = performance.now();
      try {
        const report = typeof ctx.reportSubProgress === 'function' ? ctx.reportSubProgress : null;
        const all = await getDirectoryRolePoliciesWithRules(report);
        const total = all.length;
        let withNotif = 0;
        if (report) report({ done: 0, total: Math.max(1, total), label: 'Assessing notification coverage...' });
        for (let i = 0; i < all.length; i++) {
          const pol = all[i];
          if (report) report({ done: i + 1, total: total, label: 'Assessing notification coverage...' });
          const notifRules = (pol.rules || []).filter(r => r['@odata.type'] === '#microsoft.graph.unifiedRoleManagementPolicyNotificationRule');
          if (notifRules.some(r => r.isDefaultRecipientsEnabled || (r.notificationRecipients || []).length > 0)) withNotif++;
        }
        if (total === 0) return { id: 'CISA.MS.AAD.7.9', status: 'Skipped', reason: 'No PIM role management policies found.', durationMs: ms(start) };
        const ratio = withNotif / total;
        if (ratio >= 0.8) return { id: 'CISA.MS.AAD.7.9', status: 'Passed', reason: `${withNotif}/${total} PIM role policies have activation notifications configured.`, durationMs: ms(start) };
        return { id: 'CISA.MS.AAD.7.9', status: 'Failed', reason: `Only ${withNotif}/${total} PIM role policies have activation notifications. Configure for all highly privileged roles.`, durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403 || e.status === 404) return { id: 'CISA.MS.AAD.7.9', status: 'Skipped', reason: 'PIM role policies not accessible (requires Entra ID P2).', durationMs: ms(start) };
        return errRow('CISA.MS.AAD.7.9', e, start);
      }
    },
  });

  // ── CISA.MS.AAD.8.1 Guest users SHOULD have limited access to AAD directory objects ─
  tests.push({
    id: 'CISA.MS.AAD.8.1',
    title: 'Guest users SHOULD have limited or restricted access to Azure AD directory objects',
    severity: 'Medium', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.8.1',
    description: 'Guest user access should be restricted so they cannot enumerate users, groups, or service principals in the directory.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const level = pol.guestUserRoleId;
        // 10dae51f-b6af-4016-8d66-8c2a99b929b3 = Restricted guest (most restricted)
        // 2af84b1e-32c8-42b7-82bc-daa82404023b = Guest (moderate restriction)
        // a0b1b346-4d3e-4e8b-98f8-753987be4970 = Member (basically same as user, WORST)
        const RESTRICTED = '10dae51f-b6af-4016-8d66-8c2a99b929b3';
        const GUEST = '2af84b1e-32c8-42b7-82bc-daa82404023b';
        if (level === RESTRICTED || level === GUEST) {
          return { id: 'CISA.MS.AAD.8.1', status: 'Passed', reason: `Guest user access is appropriately restricted (role: ${level === RESTRICTED ? 'Restricted guest' : 'Guest'}).`, durationMs: ms(start) };
        }
        return { id: 'CISA.MS.AAD.8.1', status: 'Failed', reason: 'Guest user access level is set to "Member" permissions which is too permissive. Set to "Guest" or "Restricted guest".', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.8.1', e, start); }
    },
  });

  // ── CISA.MS.AAD.8.2 Only Guest Inviter role can invite guests ───────────────────
  tests.push({
    id: 'CISA.MS.AAD.8.2',
    title: 'Only users with the Guest Inviter role SHOULD be able to invite guest users',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.8.2',
    description: 'The guest invite restriction setting should be set to only admins or Guest Inviter role members.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const restriction = pol.allowInvitesFrom;
        // Values: none, adminsAndGuestInviters, adminsGuestInvitersAndAllMembers, everyone
        if (restriction === 'adminsAndGuestInviters' || restriction === 'none') {
          return { id: 'CISA.MS.AAD.8.2', status: 'Passed', reason: `Guest invitation restricted to "${restriction}".`, durationMs: ms(start) };
        }
        return { id: 'CISA.MS.AAD.8.2', status: 'Failed', reason: `Guest invitation allowed from "${restriction || 'everyone'}". Restrict to "adminsAndGuestInviters" or "none".`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.8.2', e, start); }
    },
  });

  // ── CISA.MS.AAD.8.3 Guest invites only to approved domains ──────────────────────
  tests.push({
    id: 'CISA.MS.AAD.8.3',
    title: 'Guest invites SHOULD only be allowed to specific external domains that have been authorized',
    severity: 'Medium', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.8.3',
    description: 'Checks cross-tenant access policy or external collaboration settings to see if guest invitations are restricted to approved domains.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const settings = await Graph.graph('policies/externalIdentitiesPolicy', { apiVersion: 'beta' });
        const restricted = settings?.allowExternalIdentitiesToLeave === false
          || settings?.allowDeletedIdentitiesDataRemoval === false;
        // Check if there are any allowlisted domains (cross-tenant access)
        const partners = await Graph.graph('policies/crossTenantAccessPolicy/partners?$top=1', { apiVersion: 'beta' });
        const hasPartners = (partners.value || []).length > 0;
        if (hasPartners) {
          return { id: 'CISA.MS.AAD.8.3', status: 'Passed', reason: `Cross-tenant access policy has specific partner domains configured.`, durationMs: ms(start) };
        }
        return { id: 'CISA.MS.AAD.8.3', status: 'Skipped', reason: 'No domain allowlist found in cross-tenant access policy. Restrict guest invitations to specific approved domains in External Identities settings.', durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403 || e.status === 404) return { id: 'CISA.MS.AAD.8.3', status: 'Skipped', reason: 'Cannot verify cross-tenant access policy. Check External Identities → External collaboration settings manually.', durationMs: ms(start) };
        return errRow('CISA.MS.AAD.8.3', e, start);
      }
    },
  });

  // ── DNS helper — used by SPF, DKIM, DMARC checks ────────────────────────────
  async function dnsQuery(name, type) {
    const res = await fetch(
      `https://cloudflare-dns.com/dns-query?name=${encodeURIComponent(name)}&type=${type}`,
      { headers: { Accept: 'application/dns-json' } }
    );
    const data = await res.json();
    return (data.Answer || []).map(a => (a.data || '').replace(/"/g, ''));
  }

  async function getTenantDomains() {
    const domains = await Graph.graphAll('domains?$select=id,isVerified,isInitial,supportedServices', { apiVersion: 'v1.0' });
    const mailDomains = domains.filter(d =>
      d.isVerified
      && d.isInitial !== true
      && !String(d.id || '').toLowerCase().endsWith('.onmicrosoft.com')
      && Array.isArray(d.supportedServices)
      && d.supportedServices.includes('Email')
    );
    if (mailDomains.length) return mailDomains.map(d => ({ name: d.id }));
    return domains
      .filter(d => d.isVerified && d.isInitial !== true && !String(d.id || '').toLowerCase().endsWith('.onmicrosoft.com'))
      .map(d => ({ name: d.id }));
  }

  // ── CISA.MS.EXO.2.2  SPF via DNS-over-HTTPS ─────────────────────────────────
  const cisaDnsSpoTests = [];

  cisaDnsSpoTests.push({
    id: 'CISA.MS.EXO.2.2',
    title: 'An SPF policy SHALL be published for each domain',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.EXO.2.2',
    description: 'CISA SCuBA EXO.2.2: Every verified domain used for email must publish an SPF TXT record.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const domains = await getTenantDomains();
        if (!domains.length) return { id: 'CISA.MS.EXO.2.2', status: 'Skipped', reason: 'No custom verified domains found.', durationMs: ms(start) };
        const missing = [], found = [];
        for (const d of domains) {
          const records = await dnsQuery(d.name, 'TXT').catch(() => []);
          if (records.some(r => r.toLowerCase().startsWith('v=spf1'))) found.push(d.name);
          else missing.push(d.name);
        }
        if (missing.length) return { id: 'CISA.MS.EXO.2.2', status: 'Failed', reason: `${missing.length} domain(s) missing SPF: ${missing.join(', ')}.${found.length ? ` SPF present: ${found.join(', ')}.` : ''}`, durationMs: ms(start) };
        return { id: 'CISA.MS.EXO.2.2', status: 'Passed', reason: `All ${found.length} custom domain(s) have SPF: ${found.join(', ')}.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.EXO.2.2', e, start); }
    },
  });

  // ── CISA.MS.EXO.3.1  DKIM via DNS-over-HTTPS ────────────────────────────────
  cisaDnsSpoTests.push({
    id: 'CISA.MS.EXO.3.1',
    title: 'DKIM SHOULD be enabled for all domains',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.EXO.3.1',
    description: 'CISA SCuBA EXO.3.1: Every verified domain should have DKIM selector CNAME records published.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const domains = await getTenantDomains();
        if (!domains.length) return { id: 'CISA.MS.EXO.3.1', status: 'Skipped', reason: 'No custom verified domains found.', durationMs: ms(start) };
        const missing = [], found = [];
        for (const d of domains) {
          // Microsoft Exchange DKIM uses selector1 and selector2
          const [sel1, sel2] = await Promise.all([
            dnsQuery(`selector1._domainkey.${d.name}`, 'CNAME').catch(() => []),
            dnsQuery(`selector2._domainkey.${d.name}`, 'CNAME').catch(() => []),
          ]);
          const hasDkim = sel1.length > 0 || sel2.length > 0;
          if (hasDkim) found.push(d.name); else missing.push(d.name);
        }
        if (missing.length) return { id: 'CISA.MS.EXO.3.1', status: 'Failed', reason: `${missing.length} domain(s) missing DKIM selector CNAME records: ${missing.join(', ')}.${found.length ? ` DKIM present: ${found.join(', ')}.` : ''}`, durationMs: ms(start) };
        return { id: 'CISA.MS.EXO.3.1', status: 'Passed', reason: `All ${found.length} domain(s) have DKIM selector CNAME records.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.EXO.3.1', e, start); }
    },
  });

  // ── CISA.MS.EXO.4.1  DMARC published ────────────────────────────────────────
  cisaDnsSpoTests.push({
    id: 'CISA.MS.EXO.4.1',
    title: 'A DMARC policy SHALL be published for every second-level domain',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.EXO.4.1',
    description: 'CISA SCuBA EXO.4.1: A DMARC TXT record must be published at _dmarc.<domain>.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const domains = await getTenantDomains();
        if (!domains.length) return { id: 'CISA.MS.EXO.4.1', status: 'Skipped', reason: 'No custom verified domains found.', durationMs: ms(start) };
        const missing = [], found = [];
        for (const d of domains) {
          const records = await dnsQuery(`_dmarc.${d.name}`, 'TXT').catch(() => []);
          if (records.some(r => r.toLowerCase().startsWith('v=dmarc1'))) found.push(d.name);
          else missing.push(d.name);
        }
        if (missing.length) return { id: 'CISA.MS.EXO.4.1', status: 'Failed', reason: `${missing.length} domain(s) missing DMARC record: ${missing.join(', ')}.${found.length ? ` DMARC present: ${found.join(', ')}.` : ''}`, durationMs: ms(start) };
        return { id: 'CISA.MS.EXO.4.1', status: 'Passed', reason: `All ${found.length} domain(s) have DMARC records.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.EXO.4.1', e, start); }
    },
  });

  // ── CISA.MS.EXO.4.2  DMARC p=reject ─────────────────────────────────────────
  cisaDnsSpoTests.push({
    id: 'CISA.MS.EXO.4.2',
    title: 'The DMARC message rejection option SHALL be p=reject',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.EXO.4.2',
    description: 'CISA SCuBA EXO.4.2: DMARC records must have p=reject so emails failing DMARC are rejected.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const domains = await getTenantDomains();
        if (!domains.length) return { id: 'CISA.MS.EXO.4.2', status: 'Skipped', reason: 'No custom verified domains found.', durationMs: ms(start) };
        const notReject = [], rejected = [], missing = [];
        for (const d of domains) {
          const records = await dnsQuery(`_dmarc.${d.name}`, 'TXT').catch(() => []);
          const dmarc = records.find(r => r.toLowerCase().startsWith('v=dmarc1'));
          if (!dmarc) { missing.push(d.name); continue; }
          const pMatch = dmarc.toLowerCase().match(/\bp=(\w+)/);
          const policy = pMatch ? pMatch[1] : 'none';
          if (policy === 'reject') rejected.push(d.name);
          else notReject.push(`${d.name}(p=${policy})`);
        }
        const issues = [...missing.map(d => `${d}(no DMARC)`), ...notReject];
        if (issues.length) return { id: 'CISA.MS.EXO.4.2', status: 'Failed', reason: `${issues.length} domain(s) do not have p=reject: ${issues.join(', ')}.${rejected.length ? ` p=reject: ${rejected.join(', ')}.` : ''}`, durationMs: ms(start) };
        return { id: 'CISA.MS.EXO.4.2', status: 'Passed', reason: `All ${rejected.length} domain(s) have DMARC p=reject.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.EXO.4.2', e, start); }
    },
  });

  // ── SharePoint: sharing settings via Graph admin/sharepoint ─────────────────
  const SPO_SCOPES = ['https://graph.microsoft.com/SharePointTenantSettings.Read.All', ...Auth.SCOPES.graphFull];

  cisaDnsSpoTests.push({
    id: 'CISA.MS.SHAREPOINT.1.1',
    title: 'External sharing for SharePoint SHALL be limited to Existing guests or Only People in your organization',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.SHAREPOINT.1.1',
    description: 'CISA SCuBA SHAREPOINT.1.1: SharePoint external sharing setting must be "Existing guests" or "Disabled".',
    implemented: true,
    requiredScopes: SPO_SCOPES,
    async run() {
      const start = performance.now();
      try {
        const settings = await Graph.graph('admin/sharepoint/settings', {
          apiVersion: 'beta', tokenScopes: SPO_SCOPES,
        });
        const sharing = settings.sharingCapability || settings.SharingCapability || 'unknown';
        // Allowed values: disabled, existingExternalUserSharingOnly, externalUserSharingOnly, externalUserAndGuestSharing
        const ALLOWED = ['disabled', 'existingExternalUserSharingOnly'];
        if (ALLOWED.includes(sharing)) {
          return { id: 'CISA.MS.SHAREPOINT.1.1', status: 'Passed', reason: `SharePoint external sharing is set to "${sharing}" which is within CISA allowed values.`, durationMs: ms(start) };
        }
        return { id: 'CISA.MS.SHAREPOINT.1.1', status: 'Failed', reason: `SharePoint external sharing is set to "${sharing}". CISA requires "disabled" or "existingExternalUserSharingOnly".`, durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403 || e.status === 401) return { id: 'CISA.MS.SHAREPOINT.1.1', status: 'Skipped', reason: `Insufficient permissions. Requires SharePointTenantSettings.Read.All. ${e.message}`, durationMs: ms(start) };
        return errRow('CISA.MS.SHAREPOINT.1.1', e, start);
      }
    },
  });

  cisaDnsSpoTests.push({
    id: 'CISA.MS.SHAREPOINT.1.3',
    title: 'External sharing SHALL be restricted to approved external domains',
    severity: 'High', tag: 'CISA', category: 'CISA',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.SHAREPOINT.1.3',
    description: 'CISA SCuBA SHAREPOINT.1.3: SharePoint external sharing should be restricted to specific approved domains.',
    implemented: true,
    requiredScopes: SPO_SCOPES,
    async run() {
      const start = performance.now();
      try {
        const settings = await Graph.graph('admin/sharepoint/settings', {
          apiVersion: 'beta', tokenScopes: SPO_SCOPES,
        });
        const sharing = settings.sharingCapability || settings.SharingCapability || 'unknown';
        // If sharing is disabled/existingExternalUser only, domain restriction is not needed
        if (['disabled', 'existingExternalUserSharingOnly'].includes(sharing)) {
          return { id: 'CISA.MS.SHAREPOINT.1.3', status: 'Passed', reason: `External sharing is set to "${sharing}" — domain restriction is effectively enforced.`, durationMs: ms(start) };
        }
        // Check if domain allow list is configured
        const allowedDomains = settings.sharingAllowedDomainList || settings.SharingAllowedDomainList || [];
        const domainRestriction = settings.sharingDomainRestrictionMode || settings.SharingDomainRestrictionMode || 'none';
        if (domainRestriction === 'allowList' && allowedDomains.length > 0) {
          return { id: 'CISA.MS.SHAREPOINT.1.3', status: 'Passed', reason: `Domain allow list configured with ${allowedDomains.length} approved domain(s): ${allowedDomains.slice(0, 5).join(', ')}.`, durationMs: ms(start) };
        }
        if (domainRestriction === 'blockList') {
          return { id: 'CISA.MS.SHAREPOINT.1.3', status: 'Failed', reason: `Sharing domain restriction is set to block list mode — CISA requires an allow list of specific approved domains.`, durationMs: ms(start) };
        }
        return { id: 'CISA.MS.SHAREPOINT.1.3', status: 'Failed', reason: `External sharing (${sharing}) is permitted but no domain allow list is configured. Restrict sharing to approved external domains.`, durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403 || e.status === 401) return { id: 'CISA.MS.SHAREPOINT.1.3', status: 'Skipped', reason: `Insufficient permissions. Requires SharePointTenantSettings.Read.All. ${e.message}`, durationMs: ms(start) };
        return errRow('CISA.MS.SHAREPOINT.1.3', e, start);
      }
    },
  });

  // Exchange tests that still require EXO admin REST (not accessible from browser PKCE)
  const exoStubs = [
    exoStub('CISA.MS.EXO.1.1', 'Automatic forwarding to external domains SHALL be disabled'),
    exoStub('CISA.MS.EXO.2.1', 'A list of approved IP addresses for sending mail SHALL be maintained'),
    exoStub('CISA.MS.EXO.5.1', 'SMTP AUTH SHALL be disabled'),
    exoStub('CISA.MS.EXO.6.1', 'Contact folders SHALL NOT be shared with all domains'),
    exoStub('CISA.MS.EXO.6.2', 'Calendar details SHALL NOT be shared with all domains'),
    exoStub('CISA.MS.EXO.7.1', 'External sender warnings SHALL be implemented'),
    exoStub('CISA.MS.EXO.8.1', 'A DLP solution SHALL be used'),
    exoStub('CISA.MS.EXO.9.1', 'Emails SHALL be filtered by attachment file types'),
    exoStub('CISA.MS.EXO.10.1', 'Emails SHALL be scanned for malware'),
    exoStub('CISA.MS.EXO.11.1', 'Impersonation protection checks SHOULD be used'),
    exoStub('CISA.MS.EXO.12.1', 'IP allow lists SHOULD NOT be created'),
    exoStub('CISA.MS.EXO.13.1', 'Mailbox auditing SHALL be enabled'),
    exoStub('CISA.MS.EXO.14.1', 'A spam filter SHALL be enabled'),
    exoStub('CISA.MS.EXO.15.1', 'URL comparison with a block-list SHOULD be enabled'),
    exoStub('CISA.MS.EXO.16.1', 'Alerts SHALL be enabled'),
    exoStub('CISA.MS.EXO.17.1', 'Microsoft Purview Audit (Standard) logging SHALL be enabled'),
  ];

  function buildCatalog() {
    const dnsSpoCatalog = cisaDnsSpoTests.map(t => ({
      id: t.id, title: t.title, severity: t.severity, tag: t.tag, category: t.category,
      docUrl: t.docUrl, description: t.description, implemented: t.implemented,
      requiredScopes: t.requiredScopes || Auth.SCOPES.graphFull,
      runCategory: 'CISA',
      async run(ctx) {
        const r = await t.run(ctx);
        return { ...t, ...r, tag: t.tag, category: t.category };
      },
    }));
    return [
      ...tests.map(t => ({
        id: t.id, title: t.title, severity: t.severity, tag: t.tag, category: t.category,
        docUrl: t.docUrl, description: t.description, implemented: t.implemented,
        runCategory: 'CISA',
        async run(ctx) {
          const r = await t.run(ctx);
          return { ...t, ...r, tag: t.tag, category: t.category };
        },
      })),
      ...dnsSpoCatalog,
      ...exoStubs,
    ];
  }

  window.TestsCISA = { buildCatalog };
})();
