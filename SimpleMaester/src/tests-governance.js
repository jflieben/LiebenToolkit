// Governance, Azure RBAC, Access Reviews, Entitlement Management, Purview, and
// Security tests that require additional OAuth scopes beyond the base graphFull set.
// Each test declares requiredScopes so the runner pre-consents before scanning.
//
// Reference: https://maester.dev/docs/tests/
(() => {
  // ── Scope constants ──────────────────────────────────────────────────────────
  const ACCESS_REVIEW_SCOPES  = ['https://graph.microsoft.com/AccessReview.Read.All', ...Auth.SCOPES.graphFull];
  const ENTITLEMENT_SCOPES    = ['https://graph.microsoft.com/EntitlementManagement.Read.All', ...Auth.SCOPES.graphFull];
  const PURVIEW_SCOPES        = ['https://graph.microsoft.com/InformationProtectionPolicy.Read', ...Auth.SCOPES.graphFull];
  const SECURITY_EVENTS_SCOPES = Auth.SCOPES.graphSecurityEvents;
  const ARM_SCOPES            = ['https://management.azure.com/.default'];

  function ms(t0) { return Math.round(performance.now() - t0); }
  function skipOn(id, e, t0, extra) {
    if (e.status === 403 || e.status === 401 || e.status === 404) {
      return { id, status: 'Skipped', reason: `Insufficient permissions (HTTP ${e.status})${extra ? ` — ${extra}` : ''}. ${e.message}`, durationMs: ms(t0) };
    }
    return { id, status: 'Error', reason: e.message, durationMs: ms(t0) };
  }

  const tests = [];

  // ─────────────────────────────────────────────────────────────────────────────
  // LICENSING
  // ─────────────────────────────────────────────────────────────────────────────

  // SKU part numbers that include Entra ID P1 or P2 (incomplete list but covers common ones)
  const P1_SKUS = new Set([
    'AAD_PREMIUM', 'EMS', 'EMSPREMIUM', 'SPE_E3', 'SPE_E5', 'SPE_F1',
    'M365EDU_A3_FACULTY', 'M365EDU_A3_STUDENT', 'M365EDU_A5_FACULTY', 'M365EDU_A5_STUDENT',
    'ENTERPRISEPREMIUM', 'ENTERPRISEPREMIUM_NOPSTNCONF', 'SPB',
  ]);
  const P2_SKUS = new Set([
    'AAD_PREMIUM_P2', 'EMSPREMIUM', 'SPE_E5', 'M365EDU_A5_FACULTY', 'M365EDU_A5_STUDENT',
    'ENTERPRISEPREMIUM', 'SPB',
  ]);

  tests.push({
    id: 'MT.1022',
    title: 'Entra ID P1 licensing is available for all users requiring CA',
    severity: 'Informational', category: 'Licensing', tag: 'Licensing',
    docUrl: 'https://maester.dev/docs/tests/MT.1022',
    description: 'Entra ID P1 (or Microsoft 365 E3+) licensing should be available. CA policies require P1.',
    async run() {
      const t0 = performance.now();
      try {
        const skus = await Graph.graphAll('subscribedSkus', { apiVersion: 'v1.0' });
        const p1 = skus.filter(s => P1_SKUS.has(s.skuPartNumber) && s.prepaidUnits?.enabled > 0);
        if (!p1.length) return { id: 'MT.1022', status: 'Failed', reason: `No Entra ID P1 SKUs found in subscribed licenses. Found: ${skus.map(s => s.skuPartNumber).join(', ')}`, durationMs: ms(t0) };
        const total = p1.reduce((n, s) => n + (s.prepaidUnits?.enabled || 0), 0);
        return { id: 'MT.1022', status: 'Passed', reason: `${p1.length} P1 SKU(s) found with ${total} total enabled unit(s): ${p1.map(s => s.skuPartNumber).join(', ')}.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1022', e, t0); }
    },
  });

  tests.push({
    id: 'MT.1023',
    title: 'Entra ID P2 licensing is available for PIM and Identity Protection features',
    severity: 'Informational', category: 'Licensing', tag: 'Licensing',
    docUrl: 'https://maester.dev/docs/tests/MT.1023',
    description: 'Entra ID P2 (or Microsoft 365 E5+) licensing should be available for PIM and Identity Protection.',
    async run() {
      const t0 = performance.now();
      try {
        const skus = await Graph.graphAll('subscribedSkus', { apiVersion: 'v1.0' });
        const p2 = skus.filter(s => P2_SKUS.has(s.skuPartNumber) && s.prepaidUnits?.enabled > 0);
        if (!p2.length) return { id: 'MT.1023', status: 'Failed', reason: `No Entra ID P2 SKUs found. PIM and Identity Protection features require P2. Found: ${skus.map(s => s.skuPartNumber).join(', ')}`, durationMs: ms(t0) };
        const total = p2.reduce((n, s) => n + (s.prepaidUnits?.enabled || 0), 0);
        return { id: 'MT.1023', status: 'Passed', reason: `${p2.length} P2 SKU(s) found with ${total} total enabled unit(s): ${p2.map(s => s.skuPartNumber).join(', ')}.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1023', e, t0); }
    },
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // PIM POLICY RULES (MT.1085, MT.1111, MT.1112)
  // These reuse the existing beta+P2 access; same 403/404 graceful skip as
  // the PIM alert tests already in tests-native.js.
  // ─────────────────────────────────────────────────────────────────────────────

  // Fetch all roleManagementPolicies (for GA) with their activation rules expanded.
  // Returns null on 403/404 (no P2) so callers can skip gracefully.
  async function getGaActivationRules() {
    // Get policies specifically scoped to the Global Administrator role template
    const GA_TEMPLATE = '62e90394-69f5-4237-9190-012177145e10';
    // Find GA directory role
    const roles = await Graph.graphAll('directoryRoles', { apiVersion: 'v1.0' });
    const ga = roles.find(r => r.roleTemplateId === GA_TEMPLATE);
    if (!ga) return null;
    // Get PIM policy assignment for GA
    const assignments = await Graph.graphAll(
      `policies/roleManagementPolicyAssignments?$filter=scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '${GA_TEMPLATE}'`,
      { apiVersion: 'beta' }
    );
    if (!assignments.length) return null;
    const policyId = assignments[0].policyId;
    // Get the policy rules
    const rulesResp = await Graph.graph(`policies/roleManagementPolicies/${policyId}/rules`, { apiVersion: 'beta' });
    return Array.isArray(rulesResp?.value) ? rulesResp.value : [];
  }

  tests.push({
    id: 'MT.1085',
    title: 'Privileged role activations require multi-factor authentication (PIM setting)',
    severity: 'High', category: 'Privileged', tag: 'Privileged',
    docUrl: 'https://maester.dev/docs/tests/MT.1085',
    description: 'The PIM policy for the Global Administrator role should require MFA on activation.',
    async run() {
      const t0 = performance.now();
      try {
        const rules = await getGaActivationRules();
        if (!rules) return { id: 'MT.1085', status: 'Skipped', reason: 'Could not retrieve PIM policy rules — requires Entra ID P2 and PIM access.', durationMs: ms(t0) };

        // Look for enablement rule that includes MultiFactorAuthentication
        const enabRule = rules.find(r => r['@odata.type'] === '#microsoft.graph.unifiedRoleManagementPolicyEnablementRule' && r.id?.includes('Activation_'));
        if (enabRule) {
          const hasMfa = (enabRule.enabledRules || []).includes('MultiFactorAuthentication');
          if (hasMfa) return { id: 'MT.1085', status: 'Passed', reason: 'PIM activation policy for Global Administrator requires MFA (MultiFactorAuthentication in enabledRules).', durationMs: ms(t0) };
        }
        // Also check for authentication context rule
        const authCtxRule = rules.find(r => r['@odata.type'] === '#microsoft.graph.unifiedRoleManagementPolicyAuthenticationContextRule' && r.id?.includes('Activation_'));
        if (authCtxRule && authCtxRule.isEnabled) {
          return { id: 'MT.1085', status: 'Passed', reason: 'PIM activation policy for Global Administrator requires authentication context (MFA/CA).', durationMs: ms(t0) };
        }
        const current = enabRule ? (enabRule.enabledRules || []).join(', ') || 'none' : 'no enablement rule found';
        return { id: 'MT.1085', status: 'Failed', reason: `PIM activation policy for Global Administrator does NOT require MFA. Enabled rules: ${current}.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1085', e, t0, 'Requires Entra ID P2 and PIM'); }
    },
  });

  tests.push({
    id: 'MT.1111',
    title: 'Privileged role activation requires justification (PIM)',
    severity: 'High', category: 'Privileged', tag: 'Privileged',
    docUrl: 'https://maester.dev/docs/tests/MT.1111',
    description: 'The PIM activation policy for privileged roles should require a justification text.',
    async run() {
      const t0 = performance.now();
      try {
        const rules = await getGaActivationRules();
        if (!rules) return { id: 'MT.1111', status: 'Skipped', reason: 'Could not retrieve PIM policy rules — requires Entra ID P2 and PIM access.', durationMs: ms(t0) };
        const enablementRules = rules.filter(r => r['@odata.type'] === '#microsoft.graph.unifiedRoleManagementPolicyEnablementRule');
        if (!enablementRules.length) return { id: 'MT.1111', status: 'Skipped', reason: 'PIM activation enablement rule not found.', durationMs: ms(t0) };

        const activationScoped = enablementRules.filter(r => /Activation_|EndUser_Assignment|EndUser/i.test(String(r.id || '')));
        const candidates = activationScoped.length ? activationScoped : enablementRules;
        const enabled = [...new Set(candidates.flatMap(r => Array.isArray(r.enabledRules) ? r.enabledRules : []))];
        const hasJustification = enabled.includes('Justification');
        if (!hasJustification) {
          return { id: 'MT.1111', status: 'Failed', reason: `PIM activation policy for Global Administrator does NOT require justification. Enabled rules: ${enabled.join(', ') || 'none'}.`, durationMs: ms(t0) };
        }
        return { id: 'MT.1111', status: 'Passed', reason: 'PIM activation policy for Global Administrator requires justification.', durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1111', e, t0, 'Requires Entra ID P2 and PIM'); }
    },
  });

  tests.push({
    id: 'MT.1112',
    title: 'PIM privileged roles have maximum activation time configured',
    severity: 'Medium', category: 'Privileged', tag: 'Privileged',
    docUrl: 'https://maester.dev/docs/tests/MT.1112',
    description: 'The PIM activation policy for Global Administrator should have a maximum activation duration of 8 hours or less.',
    async run() {
      const t0 = performance.now();
      try {
        const rules = await getGaActivationRules();
        if (!rules) return { id: 'MT.1112', status: 'Skipped', reason: 'Could not retrieve PIM policy rules — requires Entra ID P2 and PIM access.', durationMs: ms(t0) };
        const expRule = rules.find(r => r['@odata.type'] === '#microsoft.graph.unifiedRoleManagementPolicyExpirationRule' && r.id?.includes('Activation_'));
        if (!expRule) return { id: 'MT.1112', status: 'Skipped', reason: 'PIM expiration rule for activation not found.', durationMs: ms(t0) };
        if (!expRule.isExpirationRequired) {
          return { id: 'MT.1112', status: 'Failed', reason: 'PIM activation for Global Administrator has no expiration limit (isExpirationRequired=false).', durationMs: ms(t0) };
        }
        // maximumDuration is ISO 8601 duration e.g. PT8H, P1D
        const dur = expRule.maximumDuration || '';
        // Parse hours from PT8H / PT4H / P1D etc.
        let hours = null;
        const ptH = dur.match(/PT?(\d+)H/);
        const pD = dur.match(/P(\d+)D/);
        if (ptH) hours = parseInt(ptH[1], 10);
        else if (pD) hours = parseInt(pD[1], 10) * 24;
        if (hours === null) return { id: 'MT.1112', status: 'Passed', reason: `PIM activation expiration is required. Duration: ${dur || 'not parsed'}.`, durationMs: ms(t0) };
        if (hours > 8) return { id: 'MT.1112', status: 'Failed', reason: `PIM activation maximum duration is ${hours}h — should be 8 hours or less. Current: ${dur}.`, durationMs: ms(t0) };
        return { id: 'MT.1112', status: 'Passed', reason: `PIM activation maximum duration is ${hours}h (${dur}) — within recommended 8-hour limit.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1112', e, t0, 'Requires Entra ID P2 and PIM'); }
    },
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // FEDERATED IDENTITY CREDENTIALS ON PRIVILEGED APPS (MT.1147)
  // Uses existing Application.Read.All scope — no new scope needed.
  // ─────────────────────────────────────────────────────────────────────────────
  tests.push({
    id: 'MT.1147',
    title: 'Federated identity credentials on high-privilege apps are restricted',
    severity: 'High', category: 'App', tag: 'App',
    docUrl: 'https://maester.dev/docs/tests/MT.1147',
    description: 'Applications with privileged API permissions should not have federated identity credentials configured, as these allow external identities to impersonate them.',
    async run() {
      const t0 = performance.now();
      try {
        const PRIV_API_PERMS = new Set([
          'e8f29200-9a12-4040-b9a5-8a9c0af49f62', // Directory.ReadWrite.All (app)
          '741f803b-c850-494e-b5df-cde7c675a1ca', // User.ReadWrite.All
          '9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8', // RoleManagement.ReadWrite.Directory
          '1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9', // Application.ReadWrite.All
          '06b708a9-e830-4db3-a914-8e69da51d44f', // AppRoleAssignment.ReadWrite.All
          '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3', // Application.ReadWrite.OwnedBy
          '62a82d76-70ea-41e2-9197-370581804d09', // Group.ReadWrite.All
        ]);
        const apps = await Graph.graphAll('applications', { apiVersion: 'v1.0' });
        const privApps = apps.filter(a =>
          (a.requiredResourceAccess || []).some(rra =>
            (rra.resourceAccess || []).some(ra =>
              ra.type === 'Role' && PRIV_API_PERMS.has(ra.id)
            )
          )
        );
        if (!privApps.length) return { id: 'MT.1147', status: 'Passed', reason: 'No app registrations with high-privilege application permissions found.', durationMs: ms(t0) };

        const findings = [];
        for (const app of privApps.slice(0, 30)) {
          try {
            const creds = await Graph.graphAll(`applications/${app.id}/federatedIdentityCredentials`, { apiVersion: 'v1.0' });
            if (creds.length) findings.push({ name: app.displayName, creds: creds.length });
          } catch {}
        }
        if (!findings.length) return { id: 'MT.1147', status: 'Passed', reason: `${privApps.length} privileged app(s) checked — none have federated identity credentials.`, durationMs: ms(t0) };
        const detail = findings.map(f => `${f.name} (${f.creds} federated cred(s))`).join(', ');
        return { id: 'MT.1147', status: 'Failed', reason: `${findings.length} privileged app(s) have federated identity credentials: ${detail}. Verify these are intentional and from trusted sources.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1147', e, t0); }
    },
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // ACCESS REVIEWS (MT.1090, MT.1091) — requires AccessReview.Read.All
  // ─────────────────────────────────────────────────────────────────────────────
  tests.push({
    id: 'MT.1090',
    title: 'User access reviews are configured for privileged roles',
    severity: 'High', category: 'Privileged', tag: 'Privileged',
    docUrl: 'https://maester.dev/docs/tests/MT.1090',
    description: 'Access reviews should be configured to periodically review membership in privileged Entra ID roles.',
    requiredScopes: ACCESS_REVIEW_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const defs = await Graph.graphAll('identityGovernance/accessReviews/definitions', {
          apiVersion: 'v1.0', tokenScopes: ACCESS_REVIEW_SCOPES,
        });
        if (!defs.length) return { id: 'MT.1090', status: 'Failed', reason: 'No access review definitions found. Configure access reviews for privileged Entra ID roles.', durationMs: ms(t0) };
        // Look for reviews targeting privileged roles (scope type includes directoryRole/group membership)
        const roleReviews = defs.filter(d => {
          const scope = d.scope || {};
          const resources = d.instanceEnumerationScope?.query || d.scope?.query || '';
          return (
            (d.scope?.['@odata.type'] || '').includes('principal') ||
            resources.toLowerCase().includes('roleassignment') ||
            resources.toLowerCase().includes('directoryrole') ||
            (d.displayName || '').toLowerCase().match(/privileged|admin|role|pim/i)
          );
        });
        if (!roleReviews.length) {
          return { id: 'MT.1090', status: 'Failed', reason: `${defs.length} access review(s) found but none appear to target privileged roles. Configure reviews for privileged role memberships.`, durationMs: ms(t0) };
        }
        return { id: 'MT.1090', status: 'Passed', reason: `${roleReviews.length} access review(s) targeting privileged roles found out of ${defs.length} total.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1090', e, t0, 'Requires AccessReview.Read.All scope'); }
    },
  });

  tests.push({
    id: 'MT.1091',
    title: 'Stale privileged accounts are reviewed via access reviews',
    severity: 'High', category: 'Privileged', tag: 'Privileged',
    docUrl: 'https://maester.dev/docs/tests/MT.1091',
    description: 'Access reviews should be configured to periodically review and clean up stale privileged accounts.',
    requiredScopes: ACCESS_REVIEW_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const defs = await Graph.graphAll('identityGovernance/accessReviews/definitions', {
          apiVersion: 'v1.0', tokenScopes: ACCESS_REVIEW_SCOPES,
        });
        if (!defs.length) return { id: 'MT.1091', status: 'Failed', reason: 'No access review definitions found. Configure access reviews to remove stale privileged accounts.', durationMs: ms(t0) };
        // Look for recurring reviews (not one-time) — recurring indicates ongoing hygiene
        const recurring = defs.filter(d => d.settings?.recurrence?.pattern?.type && d.settings.recurrence.pattern.type !== 'none');
        if (!recurring.length) return { id: 'MT.1091', status: 'Failed', reason: `${defs.length} access review(s) found but none are recurring. Stale account review should be a recurring process.`, durationMs: ms(t0) };
        return { id: 'MT.1091', status: 'Passed', reason: `${recurring.length} recurring access review(s) found that help manage stale accounts.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1091', e, t0, 'Requires AccessReview.Read.All scope'); }
    },
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // ENTITLEMENT MANAGEMENT (MT.1106–MT.1110) — requires EntitlementManagement.Read.All
  // ─────────────────────────────────────────────────────────────────────────────
  async function getAssignmentPolicies() {
    return Graph.graphAll('identityGovernance/entitlementManagement/assignmentPolicies', {
      apiVersion: 'v1.0', tokenScopes: ENTITLEMENT_SCOPES,
    });
  }

  tests.push({
    id: 'MT.1106',
    title: 'Access packages require approval before assignment',
    severity: 'High', category: 'Governance', tag: 'Governance',
    docUrl: 'https://maester.dev/docs/tests/MT.1106',
    description: 'Entitlement management access package assignment policies should require approval.',
    requiredScopes: ENTITLEMENT_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const policies = await getAssignmentPolicies();
        if (!policies.length) return { id: 'MT.1106', status: 'Skipped', reason: 'No access package assignment policies found (Entitlement Management may not be in use).', durationMs: ms(t0) };
        const noApproval = policies.filter(p => !p.requestApprovalSettings?.isApprovalRequired);
        if (noApproval.length) {
          return { id: 'MT.1106', status: 'Failed', reason: `${noApproval.length}/${policies.length} assignment polic${noApproval.length === 1 ? 'y does' : 'ies do'} not require approval: ${noApproval.map(p => p.displayName).slice(0, 5).join(', ')}.`, durationMs: ms(t0) };
        }
        return { id: 'MT.1106', status: 'Passed', reason: `All ${policies.length} assignment polic${policies.length === 1 ? 'y requires' : 'ies require'} approval.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1106', e, t0, 'Requires EntitlementManagement.Read.All scope'); }
    },
  });

  tests.push({
    id: 'MT.1107',
    title: 'Access packages require periodic access review',
    severity: 'High', category: 'Governance', tag: 'Governance',
    docUrl: 'https://maester.dev/docs/tests/MT.1107',
    description: 'Access package assignment policies should require periodic review of assignments.',
    requiredScopes: ENTITLEMENT_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const policies = await getAssignmentPolicies();
        if (!policies.length) return { id: 'MT.1107', status: 'Skipped', reason: 'No assignment policies found.', durationMs: ms(t0) };
        const noReview = policies.filter(p => !p.accessReviewSettings || !p.accessReviewSettings.isEnabled);
        if (noReview.length) {
          return { id: 'MT.1107', status: 'Failed', reason: `${noReview.length}/${policies.length} polic${noReview.length === 1 ? 'y does' : 'ies do'} not have access review enabled.`, durationMs: ms(t0) };
        }
        return { id: 'MT.1107', status: 'Passed', reason: `All ${policies.length} polic${policies.length === 1 ? 'y has' : 'ies have'} access review enabled.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1107', e, t0, 'Requires EntitlementManagement.Read.All scope'); }
    },
  });

  tests.push({
    id: 'MT.1108',
    title: 'Access packages require justification on request',
    severity: 'Medium', category: 'Governance', tag: 'Governance',
    docUrl: 'https://maester.dev/docs/tests/MT.1108',
    description: 'Access package requests should require requestors to provide a justification.',
    requiredScopes: ENTITLEMENT_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const policies = await getAssignmentPolicies();
        if (!policies.length) return { id: 'MT.1108', status: 'Skipped', reason: 'No assignment policies found.', durationMs: ms(t0) };
        const noJustification = policies.filter(p => !p.requestApprovalSettings?.isRequestorJustificationRequired);
        if (noJustification.length) {
          return { id: 'MT.1108', status: 'Failed', reason: `${noJustification.length}/${policies.length} polic${noJustification.length === 1 ? 'y does' : 'ies do'} not require requestor justification.`, durationMs: ms(t0) };
        }
        return { id: 'MT.1108', status: 'Passed', reason: `All ${policies.length} polic${policies.length === 1 ? 'y requires' : 'ies require'} requestor justification.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1108', e, t0, 'Requires EntitlementManagement.Read.All scope'); }
    },
  });

  tests.push({
    id: 'MT.1109',
    title: 'Connected organizations require sponsor approval',
    severity: 'Medium', category: 'Governance', tag: 'Governance',
    docUrl: 'https://maester.dev/docs/tests/MT.1109',
    description: 'External/connected organizations in entitlement management should require sponsor approval.',
    requiredScopes: ENTITLEMENT_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const orgs = await Graph.graphAll('identityGovernance/entitlementManagement/connectedOrganizations', {
          apiVersion: 'v1.0', tokenScopes: ENTITLEMENT_SCOPES,
        });
        if (!orgs.length) return { id: 'MT.1109', status: 'Skipped', reason: 'No connected organizations found in Entitlement Management.', durationMs: ms(t0) };
        // Check policies for external requestors requiring sponsor approval
        const policies = await getAssignmentPolicies();
        const withExternal = policies.filter(p =>
          (p.requestorSettings?.allowedRequestors || []).some(r => r['@odata.type']?.includes('connectedOrganization'))
        );
        const withSponsor = withExternal.filter(p =>
          (p.requestApprovalSettings?.approvalStages || []).some(s =>
            (s.primaryApprovers || []).some(a => a['@odata.type']?.includes('requestorManager') || a['@odata.type']?.includes('singleUser'))
          )
        );
        if (withExternal.length && !withSponsor.length) {
          return { id: 'MT.1109', status: 'Failed', reason: `${orgs.length} connected org(s) exist but no policies with external requestors require sponsor approval.`, durationMs: ms(t0) };
        }
        return { id: 'MT.1109', status: 'Passed', reason: `${orgs.length} connected org(s) found. External requestor policies appear to have sponsor approval configured.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1109', e, t0, 'Requires EntitlementManagement.Read.All scope'); }
    },
  });

  tests.push({
    id: 'MT.1110',
    title: 'Access package policies do not allow assignment directly by requestor',
    severity: 'High', category: 'Governance', tag: 'Governance',
    docUrl: 'https://maester.dev/docs/tests/MT.1110',
    description: 'Assignment policies should require approval — requestors should not be able to assign themselves access.',
    requiredScopes: ENTITLEMENT_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const policies = await getAssignmentPolicies();
        if (!policies.length) return { id: 'MT.1110', status: 'Skipped', reason: 'No assignment policies found.', durationMs: ms(t0) };
        // Self-assignment: requestorSettings allows requestors AND no approval required
        const selfAssignable = policies.filter(p => {
          const approvers = p.requestApprovalSettings;
          const noApproval = !approvers?.isApprovalRequired;
          const requestorCanRequest = (p.requestorSettings?.allowedRequestors || []).some(r =>
            r['@odata.type']?.includes('allMemberUsers') || r['@odata.type']?.includes('groupMember')
          );
          return noApproval && requestorCanRequest;
        });
        if (selfAssignable.length) {
          return { id: 'MT.1110', status: 'Failed', reason: `${selfAssignable.length} assignment polic${selfAssignable.length === 1 ? 'y allows' : 'ies allow'} self-service access without approval: ${selfAssignable.map(p => p.displayName).slice(0, 5).join(', ')}.`, durationMs: ms(t0) };
        }
        return { id: 'MT.1110', status: 'Passed', reason: `None of ${policies.length} assignment polic${policies.length === 1 ? 'y allows' : 'ies allow'} self-assignment without approval.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1110', e, t0, 'Requires EntitlementManagement.Read.All scope'); }
    },
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // PURVIEW — SENSITIVITY LABELS (MT.PUR.1001) — requires InformationProtectionPolicy.Read
  // ─────────────────────────────────────────────────────────────────────────────
  tests.push({
    id: 'MT.PUR.1001',
    title: 'Sensitivity labels are published',
    severity: 'Medium', category: 'Purview', tag: 'Purview',
    docUrl: 'https://maester.dev/docs/tests/MT.PUR.1001',
    description: 'Microsoft Purview sensitivity labels should be configured and published to users.',
    requiredScopes: PURVIEW_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const labels = await Graph.graphAll('informationProtection/policy/labels', {
          apiVersion: 'beta', tokenScopes: PURVIEW_SCOPES,
        });
        if (!labels.length) return { id: 'MT.PUR.1001', status: 'Failed', reason: 'No sensitivity labels found. Configure and publish labels in Microsoft Purview.', durationMs: ms(t0) };
        const active = labels.filter(l => l.isActive !== false);
        if (!active.length) return { id: 'MT.PUR.1001', status: 'Failed', reason: `${labels.length} sensitivity label(s) found but none are active/published.`, durationMs: ms(t0) };
        return { id: 'MT.PUR.1001', status: 'Passed', reason: `${active.length} active sensitivity label(s) found: ${active.map(l => l.name).slice(0, 5).join(', ')}.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.PUR.1001', e, t0, 'Requires InformationProtectionPolicy.Read scope'); }
    },
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // MICROSOFT DEFENDER FOR IDENTITY (MT.1059)
  // Checks for the AATP/MDI service principal and security alerts.
  // ─────────────────────────────────────────────────────────────────────────────
  tests.push({
    id: 'MT.1059',
    title: 'Microsoft Defender for Identity is provisioned and connected',
    severity: 'High', category: 'Defender', tag: 'Defender',
    docUrl: 'https://maester.dev/docs/tests/MT.1059',
    description: 'Microsoft Defender for Identity (MDI) should be provisioned in the tenant for on-premises identity protection.',
    requiredScopes: SECURITY_EVENTS_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        // MDI registers a service principal with this app ID when provisioned
        const MDI_APP_IDS = [
          '7b7531ad-5926-4f2d-8a1d-38495ad33e17', // Azure Advanced Threat Protection (old)
          'fc780465-2017-40d4-a0c5-307022471b92', // Microsoft Defender for Identity
        ];
        let found = null;
        for (const appId of MDI_APP_IDS) {
          try {
            const res = await Graph.graphAll(`servicePrincipals?$filter=appId eq '${appId}'`, { apiVersion: 'v1.0' });
            if (res.length) { found = res[0]; break; }
          } catch {}
        }
        if (!found) {
          // Fallback: check secure score control profiles for MDI
          try {
            const score = await Graph.graph('security/secureScores?$top=1', { apiVersion: 'beta', tokenScopes: SECURITY_EVENTS_SCOPES });
            const latest = (score.value || [score])[0];
            const mdiControl = (latest?.controlScores || []).find(c =>
              (c.controlName || '').toLowerCase().includes('defender.identity') ||
              (c.controlName || '').toLowerCase().includes('mdi') ||
              (c.controlName || '').toLowerCase().includes('aatp')
            );
            if (mdiControl) {
              const score100 = Math.round((mdiControl.score / mdiControl.total) * 100);
              return { id: 'MT.1059', status: score100 > 50 ? 'Passed' : 'Failed', reason: `MDI found via Secure Score. Control: ${mdiControl.controlName}, score: ${mdiControl.score}/${mdiControl.total}.`, durationMs: ms(t0) };
            }
          } catch {}
          return { id: 'MT.1059', status: 'Failed', reason: 'Microsoft Defender for Identity service principal not found. MDI does not appear to be provisioned in this tenant.', durationMs: ms(t0) };
        }
        const enabled = found.accountEnabled !== false;
        if (!enabled) return { id: 'MT.1059', status: 'Failed', reason: `MDI service principal found (${found.displayName}) but is disabled.`, durationMs: ms(t0) };
        return { id: 'MT.1059', status: 'Passed', reason: `Microsoft Defender for Identity is provisioned: ${found.displayName} (appId: ${found.appId}).`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1059', e, t0); }
    },
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // AZURE RBAC (MT.1056, MT.1064, MT.1065) — requires management.azure.com token
  // ─────────────────────────────────────────────────────────────────────────────

  async function armFetch(url) {
    const tok = await Auth.getToken(ARM_SCOPES);
    const res = await fetch(url, { headers: { Authorization: 'Bearer ' + tok } });
    if (res.status === 403 || res.status === 401) {
      const err = new Error(`ARM 403 for ${url} — insufficient Azure RBAC read permissions`);
      err.status = 403; throw err;
    }
    if (!res.ok) {
      const t = await res.text().catch(() => '');
      const err = new Error(`ARM ${res.status}: ${t.substring(0, 200)}`);
      err.status = res.status; throw err;
    }
    return res.json();
  }

  async function listSubscriptions() {
    const data = await armFetch('https://management.azure.com/subscriptions?api-version=2022-12-01');
    return (data.value || []).filter(s => s.state === 'Enabled');
  }

  tests.push({
    id: 'MT.1056',
    title: 'Privileged Identity Management is used for Azure resource roles',
    severity: 'High', category: 'Azure', tag: 'Azure',
    docUrl: 'https://maester.dev/docs/tests/MT.1056',
    description: 'Azure RBAC privileged roles (Owner, Contributor) should be assigned as eligible (via PIM) rather than as permanent active assignments.',
    requiredScopes: ARM_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const subs = await listSubscriptions();
        if (!subs.length) return { id: 'MT.1056', status: 'Skipped', reason: 'No Azure subscriptions found or accessible.', durationMs: ms(t0) };
        let totalEligible = 0, totalPermanent = 0, subDetails = [];
        for (const sub of subs.slice(0, 5)) {
          try {
            // Check for PIM eligible role assignment schedule instances
            const eligible = await armFetch(
              `https://management.azure.com/subscriptions/${sub.subscriptionId}/providers/Microsoft.Authorization/roleEligibilityScheduleInstances?api-version=2020-10-01-preview`
            );
            const eligCount = (eligible.value || []).length;
            // Check for permanent active role assignments for Owner/Contributor
            const OWNER_ID = '8e3af657-a8ff-443c-a75c-2fe8c4bcb635';
            const CONTRIB_ID = 'b24988ac-6180-42a0-ab88-20f7382dd24c';
            const permanent = await armFetch(
              `https://management.azure.com/subscriptions/${sub.subscriptionId}/providers/Microsoft.Authorization/roleAssignments?api-version=2022-04-01&$filter=atScope()`
            );
            const privPermanent = (permanent.value || []).filter(ra => {
              const rdId = (ra.properties?.roleDefinitionId || '').toLowerCase();
              return rdId.endsWith(OWNER_ID) || rdId.endsWith(CONTRIB_ID);
            });
            totalEligible += eligCount;
            totalPermanent += privPermanent.length;
            subDetails.push(`${sub.displayName}: ${eligCount} eligible, ${privPermanent.length} permanent Owner/Contributor`);
          } catch { subDetails.push(`${sub.displayName}: could not read`); }
        }
        if (!totalEligible && totalPermanent > 0) {
          return { id: 'MT.1056', status: 'Failed', reason: `No PIM-eligible Azure role assignments found, but ${totalPermanent} permanent Owner/Contributor assignments exist. Use PIM for privileged Azure roles. ${subDetails.join('; ')}`, durationMs: ms(t0) };
        }
        const status = totalPermanent > 0 ? 'Failed' : 'Passed';
        const msg = status === 'Passed'
          ? `Azure PIM is in use — ${totalEligible} eligible role assignment(s) found across ${subs.length} subscription(s).`
          : `${totalPermanent} permanent privileged role assignment(s) found alongside ${totalEligible} eligible. Move to PIM. ${subDetails.join('; ')}`;
        return { id: 'MT.1056', status, reason: msg, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1056', e, t0, 'Requires Azure subscription reader access (management.azure.com)'); }
    },
  });

  tests.push({
    id: 'MT.1064',
    title: 'Owner role is not assigned at subscription or management group level',
    severity: 'High', category: 'Azure', tag: 'Azure',
    docUrl: 'https://maester.dev/docs/tests/MT.1064',
    description: 'Direct (non-PIM) Owner role assignments at subscription or management group scope should not exist.',
    requiredScopes: ARM_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const subs = await listSubscriptions();
        if (!subs.length) return { id: 'MT.1064', status: 'Skipped', reason: 'No Azure subscriptions found or accessible.', durationMs: ms(t0) };
        const OWNER_ID = '8e3af657-a8ff-443c-a75c-2fe8c4bcb635';
        const ownerAssignments = [];
        for (const sub of subs.slice(0, 10)) {
          try {
            const ra = await armFetch(
              `https://management.azure.com/subscriptions/${sub.subscriptionId}/providers/Microsoft.Authorization/roleAssignments?api-version=2022-04-01&$filter=atScope()`
            );
            const owners = (ra.value || []).filter(r => {
              const rdId = (r.properties?.roleDefinitionId || '').toLowerCase();
              return rdId.endsWith(OWNER_ID) && r.properties?.principalType === 'User';
            });
            owners.forEach(o => ownerAssignments.push({ sub: sub.displayName, principal: o.properties?.principalId }));
          } catch {}
        }
        if (ownerAssignments.length) {
          return { id: 'MT.1064', status: 'Failed', reason: `${ownerAssignments.length} direct Owner role assignment(s) on subscriptions. Use PIM eligible assignments instead. Affected subscriptions: ${[...new Set(ownerAssignments.map(o => o.sub))].join(', ')}.`, durationMs: ms(t0) };
        }
        return { id: 'MT.1064', status: 'Passed', reason: `No direct Owner role assignments found on ${subs.length} subscription(s). Owner access appears to be managed via PIM or no assignments exist.`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1064', e, t0, 'Requires Azure subscription reader access'); }
    },
  });

  tests.push({
    id: 'MT.1065',
    title: 'Classic administrator roles are not used (Legacy Azure RBAC)',
    severity: 'Medium', category: 'Azure', tag: 'Azure',
    docUrl: 'https://maester.dev/docs/tests/MT.1065',
    description: 'Classic administrator roles (Co-Administrator, Service Administrator) are deprecated and should not be in use.',
    requiredScopes: ARM_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const subs = await listSubscriptions();
        if (!subs.length) return { id: 'MT.1065', status: 'Skipped', reason: 'No Azure subscriptions found or accessible.', durationMs: ms(t0) };
        const classicAdmins = [];
        for (const sub of subs.slice(0, 10)) {
          try {
            const ca = await armFetch(
              `https://management.azure.com/subscriptions/${sub.subscriptionId}/providers/Microsoft.Authorization/classicAdministrators?api-version=2015-07-01`
            );
            const admins = (ca.value || []).filter(a => a.properties?.role !== 'AccountAdministrator');
            admins.forEach(a => classicAdmins.push({ sub: sub.displayName, role: a.properties?.role, email: a.properties?.emailAddress }));
          } catch {}
        }
        if (classicAdmins.length) {
          const detail = classicAdmins.map(a => `${a.email} (${a.role}) on ${a.sub}`).join('; ');
          return { id: 'MT.1065', status: 'Failed', reason: `${classicAdmins.length} classic administrator(s) found — these roles are deprecated: ${detail}`, durationMs: ms(t0) };
        }
        return { id: 'MT.1065', status: 'Passed', reason: `No classic administrators (Co-Admin, Service Admin) found across ${subs.length} subscription(s).`, durationMs: ms(t0) };
      } catch (e) { return skipOn('MT.1065', e, t0, 'Requires Azure subscription reader access'); }
    },
  });

  // ── Build catalog ────────────────────────────────────────────────────────────
  function buildCatalog() {
    return tests.map(t => ({
      id: t.id, title: t.title, severity: t.severity,
      tag: t.tag || t.category, category: t.category,
      runCategory: t.category,
      docUrl: t.docUrl, description: t.description,
      implemented: true,
      requiredScopes: t.requiredScopes || Auth.SCOPES.graphFull,
      async run(ctx) {
        const r = await t.run(ctx);
        return { ...t, ...r };
      },
    }));
  }

  window.TestsGovernance = { buildCatalog };
})();
