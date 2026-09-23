/*
 * Phase 2 — Bucket A: browser-feasible Microsoft Graph controls.
 *
 * These port the data-collection + pass/fail logic of the ZeroTrustAssessment
 * PowerShell tests (powershell/tests/Test-Assessment.<id>.ps1) to delegated Graph.
 * The PS report-generation (markdown tables) is represented as structured `evidence`
 * instead of markdown. `$Database` (DuckDB) reads become live Graph queries.
 */
(() => {
  const { pass, fail, skip, skippedReason, toArray, lower, daysBetween, uniqueBy } = window.ZtaLib;

  // GSA / networkAccess endpoints are license-gated; 400/403/404 => not enabled/licensed.
  async function gsaGraph(path, query) {
    try { return await Api.graph(path, { beta: true, query: query || undefined }); }
    catch (err) { if ([400, 403, 404].includes(err.status)) return null; throw err; }
  }
  async function gsaGraphAll(path, query) {
    try { return await Api.graphAll(path, { beta: true, query: query || {} }); }
    catch (err) { if ([400, 403, 404].includes(err.status)) return null; throw err; }
  }
  const isoZ = (d) => d.toISOString().slice(0, 19) + 'Z';

  // Approximates the module's Find-ZtProfilesLinkedToPolicy: returns the GSA baseline/security
  // profiles that link the given policy id, and whether that linkage is effectively applied.
  // Baseline profiles (priority 65000) apply when enabled+link-enabled; security profiles also
  // require an attached Conditional Access policy.
  function profilesLinkedToPolicy(profiles, policyId, linkType) {
    const type = `${linkType}`.toLowerCase();
    const out = [];
    for (const prof of toArray(profiles)) {
      const links = toArray(prof.policies).filter(l => lower(l['@odata.type'] || '').includes(type) && lower(l.policy?.id) === lower(policyId));
      if (!links.length) continue;
      const isBaseline = Number(prof.priority) === 65000;
      const profileEnabled = lower(prof.state) === 'enabled';
      const linkEnabled = links.some(l => lower(l.state) === 'enabled');
      const hasCA = toArray(prof.conditionalAccessPolicies).length > 0;
      out.push({ profileId: prof.id, profileName: prof.name, isBaseline, profileEnabled, linkEnabled, hasCA, passes: profileEnabled && linkEnabled && (isBaseline || hasCA) });
    }
    return out;
  }

  async function getConditionalAccessPolicies() {
    return Api.graphAll('identity/conditionalAccess/policies', { beta: true, query: { '$top': '500' } });
  }

  // Agent identities/blueprints are Agent 365 preview derived types. Reachable via the
  // cast segment; return null when the surface isn't available (unlicensed tenants).
  async function getAgentIdentities() {
    try {
      return await Api.graphAll('servicePrincipals/microsoft.graph.agentIdentity', {
        beta: true,
        query: { '$select': 'id,appId,displayName,agentIdentityBlueprintId,customSecurityAttributes,accountEnabled' },
      });
    } catch (err) {
      if ([400, 403, 404].includes(err.status)) return null;
      throw err;
    }
  }

  // AI control-plane role catalog (mirrors Get-ZtAiAdminRoleDefinitions).
  const AI_ADMIN_ROLES = [
    { name: 'AI Administrator', id: 'd2562ede-74db-457e-a7b6-544e236ebb61', tier: 'Admin' },
    { name: 'Agent ID Administrator', id: 'db506228-d27e-4b7d-95e5-295956d6615f', tier: 'Admin' },
    { name: 'Agent ID Developer', id: 'adb2368d-a9be-41b5-8667-d96778e081b0', tier: 'Admin' },
    { name: 'Agent Registry Administrator', id: '6b942400-691f-4bf0-9d12-d8a254a2baf5', tier: 'Admin' },
    { name: 'Application Administrator', id: '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3', tier: 'Admin' },
    { name: 'Compliance Administrator', id: '17315797-102d-40b4-93e0-432062caca18', tier: 'Admin' },
    { name: 'Compliance Data Administrator', id: 'e6d1a23a-da11-4be4-9570-befc86d067a7', tier: 'Admin' },
    { name: 'Conditional Access Administrator', id: 'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9', tier: 'Admin' },
    { name: 'Global Reader', id: 'f2ef992c-3afb-46b9-b7cf-a126ee74c451', tier: 'Reader' },
    { name: 'Global Secure Access Administrator', id: 'ac434307-12b9-4fa1-a708-88bf58caabc1', tier: 'Admin' },
    { name: 'Identity Governance Administrator', id: '45d8d3c5-c802-45c6-b32a-1d70b5e1e86e', tier: 'Admin' },
    { name: 'Intune Administrator', id: '3a2c62db-5318-420d-8d74-23affee5d9d5', tier: 'Admin' },
    { name: 'Power Platform Administrator', id: '11648597-926c-4cf3-9c36-bcebb0ba8dcc', tier: 'Admin' },
    { name: 'Security Administrator', id: '194ae4cb-b126-40b2-bd5b-6091b380977d', tier: 'Admin' },
    { name: 'Security Operator', id: '5f2222b1-57c3-48ba-8ad5-d4759f1fde6f', tier: 'Admin' },
    { name: 'Security Reader', id: '5d6b6bb7-de71-4623-b4af-96380a352509', tier: 'Reader' },
    { name: 'SharePoint Administrator', id: 'f28a1f50-f6e7-4571-818b-6a12f2af6b6c', tier: 'Admin' },
  ];

  const GSA_INTERNET_PLAN = '8d23cb83-ab07-418f-8517-d7aca77307dc';
  const GSA_PRIVATE_PLAN = 'f057aab1-b184-49b2-85c0-881b02a405c5';

  // Intune endpoints are license-gated; 400/403/404 => not licensed/available.
  async function intuneGraphAll(path, query, beta) {
    try { return await Api.graphAll(path, { beta: beta !== false, query: query || {} }); }
    catch (err) { if ([400, 403, 404].includes(err.status)) return null; throw err; }
  }
  // Wi-Fi enterprise config profiles for a platform (matches iosWiFiConfiguration etc.).
  async function wifiProfilesForPlatform(platformMatch) {
    const cfgs = await intuneGraphAll('deviceManagement/deviceConfigurations', { '$expand': 'assignments', '$top': '200' });
    if (cfgs === null) return null;
    return toArray(cfgs).filter(c => { const t = lower(c['@odata.type']); return t.includes('wifi') && t.includes(platformMatch); });
  }
  const wifiSecType = (c) => lower(c.wiFiSecurityType || c.wifiSecurityType || c.WifiSecurityType || '');

  const impl = {
    // 25370 — Source IP restoration (Conditional Access signaling) is enabled.
    '25370': async (test) => {
      const settings = await gsaGraph('networkAccess/settings/conditionalAccess');
      if (settings === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access is not enabled or not licensed in this tenant.`);
      if (lower(settings.signalingStatus) === 'enabled') return pass(test, 'Global Secure Access source IP restoration (Conditional Access signaling) is enabled.', { signalingStatus: settings.signalingStatus });
      return fail(test, 'Global Secure Access signaling is disabled — user source IP is not preserved for Conditional Access and risk detection.', { signalingStatus: settings.signalingStatus || null });
    },

    // 25371 — Universal Continuous Access Evaluation is enabled for network access.
    '25371': async (test) => {
      const [gsaSettings, forwardingProfiles, caPolicies] = await Promise.all([
        gsaGraph('networkAccess/settings/conditionalAccess'),
        gsaGraphAll('networkAccess/forwardingProfiles'),
        getConditionalAccessPolicies().catch(() => []),
      ]);
      if (!gsaSettings || lower(gsaSettings.signalingStatus) !== 'enabled') {
        return skip(test, `${skippedReason('NotApplicable')} Global Secure Access Conditional Access signaling is not configured, so Universal CAE does not apply.`);
      }
      const disabling = toArray(caPolicies)
        .filter(p => lower(p.state) === 'enabled')
        .filter(p => toArray(p.conditions?.applications?.includeApplications).includes('All')
          && lower(p.sessionControls?.continuousAccessEvaluation?.mode) === 'disabled')
        .map(p => ({ id: p.id, displayName: p.displayName }));
      const profiles = toArray(forwardingProfiles).map(p => ({ name: p.name, state: p.state, type: p.trafficForwardingType }));
      if (disabling.length) return fail(test, 'Universal CAE is disabled by one or more Conditional Access policies that target all applications.', { disablingPolicies: disabling, forwardingProfiles: profiles });
      return pass(test, 'Universal Continuous Access Evaluation is enabled for Global Secure Access.', { signalingStatus: gsaSettings.signalingStatus, forwardingProfiles: profiles });
    },

    // 25372 — GSA client deployed on all managed endpoints (>=90% coverage).
    '25372': async (test) => {
      const end = new Date();
      const start = new Date(end.getTime() - 7 * 86400000);
      const pivot = new Date(end.getTime() - 86400000);
      const summary = await gsaGraph(`networkAccess/reports/getDeviceUsageSummary(startDateTime=${isoZ(start)},endDateTime=${isoZ(end)},activityPivotDateTime=${isoZ(pivot)})`);
      let devices = null;
      try { devices = await Api.graphAll('devices', { query: { '$select': 'trustType', '$top': '999' } }); }
      catch (err) { if (![400, 403, 404].includes(err.status)) throw err; }
      if (summary === null || devices === null) return skip(test, 'Unable to retrieve Global Secure Access device usage or the Entra device inventory.');

      const totalGsa = Number(summary.totalDeviceCount || 0);
      const activeGsa = Number(summary.activeDeviceCount || 0);
      const managed = toArray(devices).filter(d => ['azuread', 'serverad'].includes(lower(d.trustType))).length;
      const evidence = { totalGsaDevices: totalGsa, activeGsaDevices: activeGsa, managedEntraDevices: managed };

      if (managed === 0 && totalGsa === 0) return fail(test, 'No Global Secure Access client deployment detected and no managed Entra devices found.', evidence);
      if (managed === 0 && totalGsa > 0) return skip(test, 'GSA devices were detected but no Entra joined/hybrid-joined devices were found — deployment coverage cannot be calculated.', evidence);
      const pct = Math.round((totalGsa / managed) * 1000) / 10;
      evidence.deploymentPercentage = pct;
      if (pct >= 90) return pass(test, `Global Secure Access client is deployed on ${pct}% of managed endpoints.`, evidence);
      return fail(test, `Global Secure Access client deployment coverage is ${pct}% of managed endpoints (target is 90%).`, evidence);
    },

    // 25375 — GSA licenses available and assigned to users.
    '25375': async (test) => {
      let skus = null;
      try { skus = await Api.graphAll('subscribedSkus', { beta: true }); }
      catch (err) { if (![400, 403, 404].includes(err.status)) throw err; }
      if (skus === null) return skip(test, 'Unable to retrieve tenant subscribed SKUs.');

      const gsaPlanSet = new Set([GSA_INTERNET_PLAN, GSA_PRIVATE_PLAN]);
      const gsaSkus = toArray(skus).filter(s => toArray(s.servicePlans).some(p => gsaPlanSet.has(lower(p.servicePlanId))));
      const enabledGsaSkus = gsaSkus.filter(s => lower(s.capabilityStatus) === 'enabled');
      if (!enabledGsaSkus.length) return skip(test, `${skippedReason('NotApplicable')} No Global Secure Access licenses are available (enabled) in this tenant.`);

      const gsaSkuIds = new Set(enabledGsaSkus.map(s => lower(s.skuId)));
      const planBySku = new Map(enabledGsaSkus.map(s => [lower(s.skuId), new Set(toArray(s.servicePlans).map(p => lower(p.servicePlanId)).filter(id => gsaPlanSet.has(id)))]));

      const users = await Api.graphAll('users', { query: { '$select': 'id,displayName,userPrincipalName,assignedLicenses', '$top': '999' } });
      let internet = 0, priv = 0, any = 0;
      for (const u of toArray(users)) {
        let hasI = false, hasP = false;
        for (const lic of toArray(u.assignedLicenses)) {
          const sid = lower(lic.skuId);
          if (!gsaSkuIds.has(sid)) continue;
          const disabled = new Set(toArray(lic.disabledPlans).map(lower));
          const plans = planBySku.get(sid) || new Set();
          if (plans.has(GSA_INTERNET_PLAN) && !disabled.has(GSA_INTERNET_PLAN)) hasI = true;
          if (plans.has(GSA_PRIVATE_PLAN) && !disabled.has(GSA_PRIVATE_PLAN)) hasP = true;
        }
        if (hasI) internet++;
        if (hasP) priv++;
        if (hasI || hasP) any++;
      }
      const evidence = { enabledGsaSkus: enabledGsaSkus.map(s => s.skuPartNumber), usersWithInternetAccess: internet, usersWithPrivateAccess: priv, usersWithAnyGsa: any };
      if (any === 0) return fail(test, 'GSA licenses are available in the tenant but not assigned to any user.', evidence);
      return pass(test, 'Global Secure Access licenses are available and assigned to at least one user.', evidence);
    },

    // 25376 — Microsoft 365 traffic actively flowing through GSA.
    '25376': async (test) => {
      const end = new Date();
      const start = new Date(end.getTime() - 7 * 86400000);
      const pivot = new Date(end.getTime() - 86400000);
      const [m365Profiles, transactions, deviceUsage] = await Promise.all([
        gsaGraphAll('networkAccess/forwardingProfiles', { '$filter': "trafficForwardingType eq 'm365'" }),
        gsaGraphAll(`networkAccess/reports/transactionSummaries(startDateTime=${isoZ(start)},endDateTime=${isoZ(end)})`),
        gsaGraph(`networkAccess/reports/getDeviceUsageSummary(startDateTime=${isoZ(start)},endDateTime=${isoZ(end)},activityPivotDateTime=${isoZ(pivot)})`),
      ]);
      if (m365Profiles === null && transactions === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access reporting is not available or not licensed.`);

      const profile = toArray(m365Profiles)[0] || null;
      const profileEnabled = !!profile && lower(profile.state) === 'enabled';
      const m365Entry = toArray(transactions).find(t => lower(t.trafficType) === 'microsoft365');
      const m365Total = Number(m365Entry?.totalCount || 0);
      const evidence = {
        m365ProfileState: profile?.state || 'Not found',
        m365Transactions7d: m365Total,
        activeDevices: Number(deviceUsage?.activeDeviceCount || 0),
        totalDevices: Number(deviceUsage?.totalDeviceCount || 0),
      };
      if (profileEnabled && m365Total > 0) return pass(test, 'Microsoft 365 traffic forwarding is enabled and Microsoft 365 traffic is flowing through Global Secure Access.', evidence);
      return fail(test, 'Microsoft 365 traffic forwarding is disabled or no Microsoft 365 traffic is being tunneled through Global Secure Access.', evidence);
    },
    // 25379 — Conditional Access policies use compliant network controls.
    '25379': async (test) => {
      const settings = await gsaGraph('networkAccess/settings/conditionalAccess');
      if (!settings || lower(settings.signalingStatus) !== 'enabled') {
        return fail(test, 'Global Secure Access signaling is disabled — compliant network controls cannot function without this prerequisite.', { signalingStatus: settings?.signalingStatus || null });
      }
      const namedLocations = await Api.graphAll('identity/conditionalAccess/namedLocations', { beta: true, query: { '$top': '200' } }).catch(() => []);
      const compliant = toArray(namedLocations).find(l => lower(l['@odata.type']) === '#microsoft.graph.compliantnetworknamedlocation' && lower(l.compliantNetworkType) === 'alltenantcompliantnetworks');
      if (!compliant) return fail(test, 'A compliant network named location does not exist or is not configured for all tenant compliant networks.');
      const policies = (await getConditionalAccessPolicies().catch(() => [])).filter(p => lower(p.state) === 'enabled');
      const standard = policies.filter(p => toArray(p.conditions?.locations?.includeLocations).includes('All')
        && toArray(p.conditions?.locations?.excludeLocations).includes(compliant.id)
        && toArray(p.grantControls?.builtInControls).includes('block'));
      if (standard.length) return pass(test, 'A Conditional Access policy blocks access from all locations except the compliant network.', { compliantLocation: compliant.displayName, policies: standard.map(p => ({ id: p.id, displayName: p.displayName })) });
      return fail(test, 'No Conditional Access policy blocks access from all locations except the compliant network.', { compliantLocation: compliant.displayName });
    },

    // 25380 — GSA Conditional Access signaling is enabled.
    '25380': async (test) => {
      const settings = await gsaGraph('networkAccess/settings/conditionalAccess');
      if (!settings || !settings.signalingStatus) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access does not appear to be deployed in this tenant.`);
      if (lower(settings.signalingStatus) === 'enabled') return pass(test, 'Global Secure Access signaling for Conditional Access is enabled.', { signalingStatus: settings.signalingStatus });
      return fail(test, 'Global Secure Access signaling for Conditional Access is not enabled.', { signalingStatus: settings.signalingStatus });
    },

    // 25381 — Network traffic routed through GSA (all forwarding profiles enabled).
    '25381': async (test) => {
      const profiles = await gsaGraphAll('networkAccess/forwardingProfiles');
      if (profiles === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access is not licensed/available.`);
      if (!profiles.length) return fail(test, 'No traffic forwarding profiles found — Global Secure Access is not configured.');
      const disabled = profiles.filter(p => lower(p.state) !== 'enabled');
      const evidence = { profiles: profiles.map(p => ({ name: p.name, type: p.trafficForwardingType, state: p.state })) };
      if (!disabled.length) return pass(test, 'All traffic forwarding profiles are enabled — traffic is captured by the Security Service Edge.', evidence);
      if (disabled.length === profiles.length) return fail(test, 'All traffic forwarding profiles are disabled — Global Secure Access is not protecting any traffic.', evidence);
      return fail(test, 'Some traffic forwarding profiles are disabled — only partial traffic is protected.', evidence);
    },

    // 25382 — Forwarding profiles are scoped to users/groups.
    '25382': async (test) => {
      const profiles = await gsaGraphAll('networkAccess/forwardingProfiles', { '$select': 'id,name,state,trafficForwardingType,associations,servicePrincipal' });
      if (profiles === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access is not licensed/available.`);
      if (!profiles.length) return fail(test, 'No traffic forwarding profiles found.');
      let hasDisabled = false, hasEnabledWithoutAssignments = false;
      const detail = [];
      for (const p of profiles) {
        if (lower(p.state) !== 'enabled') { hasDisabled = true; detail.push({ name: p.name, state: p.state, assignment: 'N/A' }); continue; }
        const sp = p.servicePrincipal;
        if (!sp?.id) { hasEnabledWithoutAssignments = true; detail.push({ name: p.name, assignment: 'no service principal' }); continue; }
        try {
          const spDetail = await Api.graph(`servicePrincipals/${sp.id}`, { query: { '$select': 'appRoleAssignmentRequired', '$expand': 'appRoleAssignedTo($select=principalType)' } });
          if (!spDetail.appRoleAssignmentRequired) detail.push({ name: p.name, assignment: 'All users' });
          else if (toArray(spDetail.appRoleAssignedTo).length) detail.push({ name: p.name, assignment: `${toArray(spDetail.appRoleAssignedTo).length} assignment(s)` });
          else { hasEnabledWithoutAssignments = true; detail.push({ name: p.name, assignment: 'none' }); }
        } catch { hasEnabledWithoutAssignments = true; detail.push({ name: p.name, assignment: 'error' }); }
      }
      if (hasEnabledWithoutAssignments) return fail(test, 'One or more enabled traffic forwarding profiles have no user/group assignments.', { profiles: detail });
      if (hasDisabled) return fail(test, 'Some traffic forwarding profiles are disabled — review whether this is intentional.', { profiles: detail });
      return pass(test, 'Traffic forwarding profiles are scoped to specific users/groups.', { profiles: detail });
    },

    // 25383 — GA / GSA Administrator privileges are tightly limited.
    '25383': async (test) => {
      const roleDefs = await Api.graphAll('roleManagement/directory/roleDefinitions', { beta: true, query: { '$select': 'id,displayName,templateId', '$top': '400' } });
      const gaRole = toArray(roleDefs).find(r => r.displayName === 'Global Administrator');
      const gsaRole = toArray(roleDefs).find(r => r.displayName === 'Global Secure Access Administrator');
      if (!gaRole) return fail(test, 'Global Administrator role definition not found.');
      const targetRoleIds = new Set([gaRole.id, gsaRole?.id].filter(Boolean));
      const assignments = await Api.graphAll('roleManagement/directory/roleAssignments', { beta: true, query: { '$select': 'id,principalId,roleDefinitionId', '$expand': 'principal', '$top': '400' } });
      const relevant = toArray(assignments).filter(a => targetRoleIds.has(a.roleDefinitionId));
      const issues = [], disabled = [], valid = [];
      for (const a of relevant) {
        const type = lower(a.principal?.['@odata.type'] || '').replace('#microsoft.graph.', '');
        const p = { displayName: a.principal?.displayName, upn: a.principal?.userPrincipalName, type };
        if (type.includes('group')) { p.issue = 'Group assignment'; issues.push(p); }
        else if (type.includes('serviceprincipal')) { p.issue = 'Service principal assignment'; issues.push(p); }
        else if (type.includes('user') && lower(a.principal?.userType) === 'guest') { p.issue = 'Guest user assignment'; issues.push(p); }
        else if (type.includes('user') && a.principal?.accountEnabled === false) { p.issue = 'Disabled account with privileged role'; disabled.push(p); }
        else valid.push(p);
      }
      const evidence = { totalAssignments: relevant.length, issues, disabledAccounts: disabled, validMembers: valid.length };
      if (issues.length) return fail(test, 'Global Administrator / Global Secure Access Administrator roles include groups, guests, or service principals.', evidence);
      if (disabled.length || relevant.length > 5) return pass(test, 'Privileged network admin roles are limited to member users, but disabled accounts or a high assignment count warrant review.', evidence);
      return pass(test, 'Global Administrator / Global Secure Access Administrator roles are limited to enabled, vetted member users.', evidence);
    },

    // 25384 — Application Administrator rights constrained to specific Private Access apps.
    '25384': async (test) => {
      const roleDefs = await Api.graphAll('roleManagement/directory/roleDefinitions', { beta: true, query: { '$select': 'id,displayName', '$top': '400' } });
      const appAdmin = toArray(roleDefs).find(r => r.displayName === 'Application Administrator');
      if (!appAdmin) return skip(test, 'Application Administrator role definition not found.');
      const assignments = await Api.graphAll('roleManagement/directory/roleAssignments', { beta: true, query: { '$filter': `roleDefinitionId eq '${appAdmin.id}'`, '$expand': 'principal', '$top': '200' } });
      const tenantWide = [], problematic = [], scoped = [];
      for (const a of toArray(assignments)) {
        const type = lower(a.principal?.['@odata.type'] || '').replace('#microsoft.graph.', '');
        const info = { displayName: a.principal?.displayName, type, scope: a.directoryScopeId };
        if (a.directoryScopeId === '/') tenantWide.push(info);
        else scoped.push(info);
        if (type.includes('group') || type.includes('serviceprincipal') || lower(a.principal?.userType) === 'guest') problematic.push(info);
      }
      const evidence = { tenantWide: tenantWide.length, problematic, scoped: scoped.length, total: toArray(assignments).length };
      if (tenantWide.length) return fail(test, 'Application Administrator is assigned tenant-wide rather than scoped to specific Private Access applications.', evidence);
      if (problematic.length) return fail(test, 'Application Administrator role includes groups, guests, or service principals.', evidence);
      return pass(test, 'Application Administrator rights are scoped (not tenant-wide) and assigned only to member users.', evidence);
    },
    // 25391 — Private network connectors are active and healthy.
    '25391': async (test) => {
      const connectors = await gsaGraphAll('onPremisesPublishingProfiles/applicationProxy/connectors');
      if (connectors === null) return skip(test, `${skippedReason('NotApplicable')} Application Proxy / Private Network connectors are not available.`);
      if (!connectors.length) return fail(test, 'No Private Network connectors are configured.');
      const inactive = connectors.filter(c => lower(c.status) !== 'active');
      const evidence = { total: connectors.length, active: connectors.length - inactive.length, inactive: inactive.map(c => ({ machineName: c.machineName, version: c.version, status: c.status })) };
      if (!inactive.length) return pass(test, 'All Private Network connectors are active and healthy.', evidence);
      return fail(test, 'One or more Private Network connectors are inactive or unhealthy.', evidence);
    },

    // 25392 — Connectors on the latest version. Latest version lives on a cross-origin MS Learn page.
    '25392': async (test) => {
      const connectors = await gsaGraphAll('onPremisesPublishingProfiles/applicationProxy/connectors');
      if (connectors === null || !connectors.length) return skip(test, `${skippedReason('NotApplicable')} No Private Access connectors were detected.`);
      return skip(test, `${skippedReason('NotSupported')} The latest connector version is published on a cross-origin Microsoft Learn page not fetchable from the browser; verify connector versions manually.`, { connectors: connectors.map(c => ({ machineName: c.machineName, version: c.version })) });
    },

    // 25393 — Quick Access is enabled and bound to a connector.
    '25393': async (test) => {
      const profiles = await gsaGraphAll('networkAccess/forwardingProfiles');
      if (profiles === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access is not available.`);
      const priv = toArray(profiles).find(p => lower(p.trafficForwardingType) === 'private');
      const profileEnabled = !!priv && lower(priv.state) === 'enabled';
      const qaSp = (await Api.graphAll('servicePrincipals', { beta: true, query: { '$filter': "tags/any(t:t eq 'NetworkAccessQuickAccessApplication')", '$select': 'id,appId,displayName', '$top': '5' } }).catch(() => []))[0];
      const groups = (await gsaGraphAll('onPremisesPublishingProfiles/applicationProxy/connectorGroups')) || [];
      let qaGroup = null;
      if (qaSp) {
        for (const g of groups) {
          const apps = (await gsaGraphAll(`onPremisesPublishingProfiles/applicationProxy/connectorGroups/${g.id}/applications`)) || [];
          if (apps.some(a => lower(a.appId) === lower(qaSp.appId))) { qaGroup = g; break; }
        }
      }
      let activeConnectors = 0;
      if (qaGroup) {
        const members = (await gsaGraphAll(`onPremisesPublishingProfiles/applicationProxy/connectorGroups/${qaGroup.id}/members`)) || [];
        activeConnectors = members.filter(m => lower(m.status) === 'active').length;
      }
      const evidence = { privateProfileState: priv?.state || 'Not found', quickAccessApp: qaSp?.displayName || null, activeConnectors };
      if (profileEnabled && qaSp && qaGroup && activeConnectors > 0) return pass(test, 'Quick Access is bound to a connector group with an active connector and the Private Access profile is enabled.', evidence);
      return fail(test, 'Quick Access is not bound to a connector group with active connectors, or the Private Access profile is not enabled.', evidence);
    },

    // 25394 — Quick Access is bound to a Conditional Access policy with meaningful controls.
    '25394': async (test) => {
      const qa = (await Api.graphAll('servicePrincipals', { beta: true, query: { '$filter': "tags/any(c:c eq 'NetworkAccessQuickAccessApplication')", '$select': 'appId,displayName,id', '$top': '5' } }).catch(() => []))[0];
      if (!qa) return skip(test, `${skippedReason('NotApplicable')} No Quick Access application is configured.`);
      const policies = (await getConditionalAccessPolicies().catch(() => [])).filter(p => lower(p.state) === 'enabled');
      const meaningful = new Set(['mfa', 'compliantdevice', 'domainjoineddevice', 'approvedapplication']);
      const applicable = policies.filter(p => {
        const inc = toArray(p.conditions?.applications?.includeApplications);
        if (!(inc.includes('All') || inc.includes(qa.appId))) return false;
        const controls = toArray(p.grantControls?.builtInControls).map(lower);
        return controls.some(c => meaningful.has(c)) || !!p.grantControls?.authenticationStrength;
      });
      if (applicable.length) return pass(test, 'Quick Access is protected by a Conditional Access policy with strong grant controls.', { quickAccessApp: qa.displayName, policies: applicable.map(p => ({ id: p.id, displayName: p.displayName })) });
      return fail(test, 'Quick Access application is not protected by a Conditional Access policy with meaningful grant controls.', { quickAccessApp: qa.displayName });
    },

    // 25395 — Private Access application segments enforce least-privilege (partial: CSA check omitted; broad-scope heuristics).
    '25395': async (test) => {
      const apps = await Api.graphAll('applications', { query: { '$filter': "tags/any(t:t eq 'PrivateAccessNonWebApplication')", '$select': 'id,appId,displayName,tags', '$top': '100' } }).catch(() => null);
      if (apps === null || !toArray(apps).length) return skip(test, `${skippedReason('NotApplicable')} No Entra Private Access per-app applications are configured.`);
      const isBroadCidr = (h) => { const m = `${h || ''}`.match(/\/(\d+)$/); return m ? Number(m[1]) < 24 : false; };
      const isBroadPort = (p) => { const s = `${p || ''}`; if (s.includes('-')) { const [a, b] = s.split('-').map(Number); return (b - a) > 100; } return false; };
      const broadApps = [], findings = [];
      for (const app of toArray(apps)) {
        let broad = false;
        const segs = await Api.graph(`applications/${app.id}/onPremisesPublishing/segmentsConfiguration/microsoft.graph.ipSegmentConfiguration/applicationSegments`, { beta: true }).then(r => toArray(r?.value)).catch(() => []);
        for (const s of segs) {
          const issues = [];
          if (lower(s.destinationType) === 'dnssuffix') issues.push('Wildcard DNS');
          if (lower(s.destinationType).includes('cidr') && isBroadCidr(s.destinationHost)) issues.push('Broad CIDR');
          if (toArray(s.ports).some(isBroadPort)) issues.push('Broad port range');
          if (issues.length) { broad = true; findings.push({ app: app.displayName, destination: s.destinationHost, issues: issues.join(', ') }); }
        }
        if (broad) broadApps.push(app.displayName);
      }
      const evidence = { privateAccessApps: toArray(apps).length, appsWithBroadSegments: broadApps.length, findings: findings.slice(0, 20) };
      if (broadApps.length) return fail(test, 'One or more Private Access applications have overly broad application segments (wildcard DNS, broad CIDR, or broad port ranges).', evidence);
      return pass(test, 'Private Access application segments are scoped for least-privilege access.', evidence);
    },

    // 25396 — Conditional Access enforces strong authentication for private apps (partial: CSA scoring omitted).
    '25396': async (test) => {
      const apps = await Api.graphAll('servicePrincipals', { query: { '$filter': "(tags/any(t:t eq 'PrivateAccessNonWebApplication') or tags/any(t:t eq 'NetworkAccessQuickAccessApplication'))", '$select': 'id,displayName,appId,tags' } }).catch(() => null);
      if (apps === null || !toArray(apps).length) return skip(test, `${skippedReason('NotApplicable')} No Private Access applications are configured.`);
      const policies = (await getConditionalAccessPolicies().catch(() => [])).filter(p => lower(p.state) === 'enabled');
      const strongIds = new Set(['00000000-0000-0000-0000-000000000002', '00000000-0000-0000-0000-000000000003', '00000000-0000-0000-0000-000000000004']);
      const unprotected = [];
      for (const app of toArray(apps)) {
        const ok = policies.some(p => {
          const inc = toArray(p.conditions?.applications?.includeApplications);
          if (!(inc.includes('All') || inc.includes(app.appId))) return false;
          const controls = toArray(p.grantControls?.builtInControls).map(lower);
          const authStrengthId = lower(p.grantControls?.authenticationStrength?.id || '');
          return controls.includes('mfa') || strongIds.has(authStrengthId) || !!p.grantControls?.authenticationStrength;
        });
        if (!ok) unprotected.push(app.displayName);
      }
      const evidence = { privateAccessApps: toArray(apps).length, unprotectedCount: unprotected.length, unprotectedApps: unprotected.slice(0, 20) };
      if (unprotected.length) return fail(test, 'One or more Private Access applications are not protected by a Conditional Access policy enforcing MFA or stronger authentication.', evidence);
      return pass(test, 'Private Access applications enforce strong authentication via Conditional Access.', evidence);
    },
    // 25398 — Domain controller RDP protected by phishing-resistant auth through GSA.
    '25398': async (test) => {
      const paApps = await Api.graphAll('applications', { query: { '$filter': "tags/any(t:t eq 'PrivateAccessNonWebApplication')", '$select': 'id,appId,displayName', '$top': '100' } }).catch(() => null);
      if (paApps === null || !toArray(paApps).length) return skip(test, `${skippedReason('NotApplicable')} No Private Access applications configured.`);
      const portIncludes = (ports, target) => toArray(ports).some(p => { const s = `${p}`; if (s.includes('-')) { const [a, b] = s.split('-').map(Number); return target >= a && target <= b; } return Number(s) === target; });
      const segsFor = (id) => Api.graph(`applications/${id}/onPremisesPublishing/segmentsConfiguration/microsoft.graph.ipSegmentConfiguration/applicationSegments`, { beta: true }).then(r => toArray(r?.value)).catch(() => []);
      const appSegs = [];
      for (const app of toArray(paApps)) appSegs.push({ app, segs: await segsFor(app.id) });
      const dcHosts = new Set();
      for (const { segs } of appSegs) {
        const w88 = new Set(), w389 = new Set();
        for (const s of segs) { if (portIncludes(s.ports, 88)) w88.add(s.destinationHost); if (portIncludes(s.ports, 389)) w389.add(s.destinationHost); }
        for (const h of w88) if (w389.has(h)) dcHosts.add(h);
      }
      const rdpApps = [];
      for (const { app, segs } of appSegs) for (const s of segs) {
        if (lower(s.protocol).includes('tcp') && portIncludes(s.ports, 3389) && (dcHosts.size === 0 || dcHosts.has(s.destinationHost))) rdpApps.push({ appId: app.appId, appName: app.displayName });
      }
      if (!rdpApps.length) return skip(test, `${skippedReason('NotApplicable')} No Private Access applications with RDP access (port 3389) were found.`);
      const authStrengths = await Api.graphAll('policies/authenticationStrengthPolicies', { beta: true }).catch(() => []);
      const prMfa = toArray(authStrengths).find(a => lower(a.displayName) === 'phishing-resistant mfa');
      if (!prMfa) return skip(test, `${skippedReason('NotApplicable')} Phishing-resistant MFA authentication strength policy not found.`);
      const policies = (await getConditionalAccessPolicies().catch(() => [])).filter(p => lower(p.state) === 'enabled' && lower(p.grantControls?.authenticationStrength?.id || '') === lower(prMfa.id));
      const unprotected = [];
      for (const r of uniqueBy(rdpApps, x => x.appId)) {
        const ok = policies.some(p => { const inc = toArray(p.conditions?.applications?.includeApplications); return inc.includes('All') || inc.includes(r.appId) || !!p.conditions?.applications?.applicationFilter; });
        if (!ok) unprotected.push(r.appName);
      }
      const evidence = { rdpApps: rdpApps.length, dcHostsIdentified: dcHosts.size, unprotectedCount: unprotected.length, unprotected: unprotected.slice(0, 20) };
      if (unprotected.length) return fail(test, 'One or more RDP-capable Private Access apps are not protected by a phishing-resistant MFA Conditional Access policy.', evidence);
      return pass(test, 'Domain controller RDP access is protected by phishing-resistant authentication through Global Secure Access.', evidence);
    },

    // 25399 — Private DNS configured for internal name resolution (Quick Access).
    '25399': async (test) => {
      const qa = (await Api.graphAll('applications', { beta: true, query: { '$filter': "tags/any(c:c eq 'NetworkAccessQuickAccessApplication')", '$top': '5' } }).catch(() => []))[0];
      if (!qa) return fail(test, "No Quick Access application found (NetworkAccessQuickAccessApplication tag).");
      const onPrem = await Api.graph(`applications/${qa.id}/onPremisesPublishing`, { beta: true }).catch(() => null);
      const dnsEnabled = onPrem?.isDnsResolutionEnabled === true;
      const segs = await Api.graph(`applications/${qa.id}/onPremisesPublishing/segmentsConfiguration/microsoft.graph.ipSegmentConfiguration/applicationSegments`, { beta: true }).then(r => toArray(r?.value)).catch(() => []);
      const suffixes = [...new Set(segs.filter(s => lower(s.destinationType) === 'dnssuffix' && s.destinationHost).map(s => s.destinationHost))];
      const evidence = { quickAccessApp: qa.displayName, dnsResolutionEnabled: dnsEnabled, dnsSuffixes: suffixes };
      if (dnsEnabled && suffixes.length) return pass(test, 'Private DNS is configured for internal name resolution in Entra Private Access.', evidence);
      return fail(test, 'Private DNS is not configured or DNS suffixes are missing.', evidence);
    },

    // 25400 — DNS traffic for internal domains routed through Private Access.
    '25400': async (test) => {
      const profiles = await gsaGraphAll('networkAccess/forwardingProfiles', { '$filter': "trafficForwardingType eq 'private'" });
      if (profiles === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access is not available.`);
      const priv = toArray(profiles).find(p => lower(p.state) === 'enabled');
      if (!priv) return skip(test, `${skippedReason('NotApplicable')} Private Access is not enabled in this tenant.`);
      const apps = await Api.graphAll('applications', { query: { '$filter': "(tags/any(t:t eq 'PrivateAccessNonWebApplication') or tags/any(t:t eq 'NetworkAccessQuickAccessApplication'))", '$select': 'id,displayName', '$top': '100' } }).catch(() => []);
      if (!toArray(apps).length) return fail(test, 'No Private Access applications configured.');
      const portIncludes = (ports, target) => toArray(ports).some(p => { const s = `${p}`; if (s.includes('-')) { const [a, b] = s.split('-').map(Number); return target >= a && target <= b; } return Number(s) === target; });
      let hasDnsSuffix = false, hasPort53 = false;
      for (const app of toArray(apps)) {
        const segs = await Api.graph(`applications/${app.id}/onPremisesPublishing/segmentsConfiguration/microsoft.graph.ipSegmentConfiguration/applicationSegments`, { beta: true }).then(r => toArray(r?.value)).catch(() => []);
        for (const s of segs) {
          if (lower(s.destinationType) === 'dnssuffix' && s.destinationHost) hasDnsSuffix = true;
          else if (s.protocol && /udp|tcp/.test(lower(s.protocol)) && portIncludes(s.ports, 53)) hasPort53 = true;
        }
      }
      const evidence = { privateAccessApps: toArray(apps).length, hasDnsSuffix, hasPort53 };
      if (hasDnsSuffix || hasPort53) return pass(test, 'DNS traffic for internal domains is routed through Private Access (DNS suffix or port 53 segments configured).', evidence);
      return fail(test, 'No DNS suffix or port 53 segments are configured for Private Access — internal DNS is not routed through GSA.', evidence);
    },

    // 25401 — Application Proxy apps require preauthentication.
    '25401': async (test) => {
      const apps = await Api.graphAll('applications', { beta: true, query: { '$filter': 'onPremisesPublishing/isOnPremPublishingEnabled eq true', '$select': 'id,appId,displayName,onPremisesPublishing', '$top': '100' } }).catch(() => null);
      if (apps === null) return skip(test, `${skippedReason('NotApplicable')} Unable to determine Application Proxy configuration.`);
      const proxyApps = toArray(apps).filter(a => a.onPremisesPublishing?.isOnPremPublishingEnabled === true);
      if (!proxyApps.length) return skip(test, `${skippedReason('NotApplicable')} No Application Proxy applications are configured.`);
      const noPreauth = proxyApps.filter(a => lower(a.onPremisesPublishing?.externalAuthenticationType) !== 'aadpreauthentication')
        .map(a => ({ displayName: a.displayName, authType: a.onPremisesPublishing?.externalAuthenticationType || 'unknown' }));
      const evidence = { total: proxyApps.length, withoutPreauth: noPreauth };
      if (!noPreauth.length) return pass(test, 'All Application Proxy applications require Microsoft Entra pre-authentication.', evidence);
      return fail(test, `${noPreauth.length} Application Proxy application(s) use passthrough authentication instead of Entra pre-authentication.`, evidence);
    },
    // 25403 — Private Access sensors are active and enforcing on domain controllers.
    '25403': async (test) => {
      const sensors = await gsaGraphAll('onPremisesPublishingProfiles/privateAccess/sensors');
      if (sensors === null) return skip(test, `${skippedReason('NotApplicable')} Private Access sensors are not available.`);
      if (!sensors.length) return fail(test, 'No Private Access sensors are deployed on domain controllers.');
      const enforcing = sensors.filter(s => lower(s.status) === 'active' && s.isAuditMode === false);
      const nonEnforcing = sensors.filter(s => lower(s.status) !== 'active' || s.isAuditMode === true);
      const evidence = { total: sensors.length, enforcing: enforcing.length, nonEnforcing: nonEnforcing.length };
      if (enforcing.length > 0 && nonEnforcing.length === 0) return pass(test, 'Private Access sensors are active and enforcing (not in audit mode) on domain controllers.', evidence);
      return fail(test, 'One or more Private Access sensors are inactive or in audit mode.', evidence);
    },

    // 25405 — Intelligent Local Access is enabled and configured.
    '25405': async (test) => {
      const nets = await gsaGraphAll('networkAccess/privateNetworks');
      if (nets === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access is not available.`);
      if (!nets.length) return fail(test, 'Intelligent Local Access is not configured (no private networks defined).');
      return pass(test, `Intelligent Local Access is enabled with ${nets.length} private network(s) configured.`, { privateNetworks: nets.length });
    },

    // 25406 — Internet access forwarding profile is enabled and assigned.
    '25406': async (test) => {
      const profs = await gsaGraphAll('networkAccess/forwardingProfiles', { '$filter': "trafficForwardingType eq 'internet'", '$select': 'name,state,servicePrincipal' });
      if (profs === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access is not available.`);
      const profile = toArray(profs)[0];
      if (!profile) return fail(test, 'No Internet Access forwarding profile found.');
      const enabled = lower(profile.state) === 'enabled';
      let hasAssignments = false;
      const spId = profile.servicePrincipal?.id;
      if (spId) {
        try {
          const sp = await Api.graph(`servicePrincipals/${spId}`, { query: { '$select': 'appId,appRoleAssignmentRequired', '$expand': 'appRoleAssignedTo($select=principalId,principalType)' } });
          if (sp.appRoleAssignmentRequired === false) hasAssignments = true;
          else if (toArray(sp.appRoleAssignedTo).length > 0) hasAssignments = true;
        } catch { /* ignore */ }
      }
      const evidence = { profileName: profile.name, state: profile.state, hasAssignments };
      if (enabled && hasAssignments) return pass(test, 'Internet Access forwarding profile is enabled and assigned to users.', evidence);
      return fail(test, 'Internet Access forwarding profile is disabled or has no user/group assignments.', evidence);
    },

    // 25407 — Web content filtering integrates with Conditional Access.
    '25407': async (test) => {
      const [policies, filteringProfiles] = await Promise.all([getConditionalAccessPolicies().catch(() => []), gsaGraphAll('networkAccess/filteringProfiles')]);
      if (filteringProfiles === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access filtering is not available.`);
      const profById = new Map(toArray(filteringProfiles).map(p => [p.id, p]));
      const gsaPolicies = toArray(policies).filter(p => lower(p.state) === 'enabled' && p.sessionControls?.globalSecureAccessFilteringProfile);
      const details = gsaPolicies.map(p => { const gp = p.sessionControls.globalSecureAccessFilteringProfile; return { policy: p.displayName, caLinkageEnabled: gp.isEnabled === true, profileState: profById.get(gp.profileId)?.state }; });
      const active = details.filter(d => lower(d.profileState) === 'enabled' && d.caLinkageEnabled);
      const evidence = { gsaPolicies: details.length, activeLinkages: active.length };
      if (active.length >= 1) return pass(test, 'Web content filtering is integrated with Conditional Access via an enabled Global Secure Access filtering profile.', evidence);
      return fail(test, 'No enabled Conditional Access policy links to an enabled Global Secure Access web content filtering profile.', evidence);
    },

    // 25408 — Web content filtering policies are configured (partial: profile↔policy linkage heuristic).
    '25408': async (test) => {
      const allFilteringPolicies = await gsaGraphAll('networkAccess/filteringPolicies');
      if (allFilteringPolicies === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access filtering is not available.`);
      const wcf = toArray(allFilteringPolicies).filter(p => p.name !== 'All websites');
      if (!wcf.length) return fail(test, 'No web content filtering policies are configured (only the default "All websites" policy exists).');
      const [profiles, policies] = await Promise.all([gsaGraphAll('networkAccess/filteringProfiles', { '$expand': 'policies' }), getConditionalAccessPolicies().catch(() => [])]);
      const enabledProfiles = toArray(profiles).filter(p => lower(p.state) === 'enabled');
      const caLinksGsa = toArray(policies).some(p => lower(p.state) === 'enabled' && p.sessionControls?.globalSecureAccessFilteringProfile?.isEnabled === true);
      const evidence = { wcfPolicies: wcf.length, enabledFilteringProfiles: enabledProfiles.length, caLinkedToGsaFiltering: caLinksGsa };
      if (wcf.length && enabledProfiles.length && caLinksGsa) return pass(test, 'Web content filtering policies are configured and applied via enabled filtering profiles linked to Conditional Access.', evidence);
      return fail(test, 'Web content filtering policies exist but are not fully applied through enabled filtering profiles linked to Conditional Access.', evidence);
    },
    // 25409 — Web content filtering uses category-based rules (partial: profile-linkage heuristic).
    '25409': async (test) => {
      const allPol = await gsaGraphAll('networkAccess/filteringPolicies');
      if (allPol === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access filtering is not available.`);
      const wcf = toArray(allPol).filter(p => p.name !== 'All websites');
      if (!wcf.length) return fail(test, 'No web content filtering policies are configured.');
      const profiles = (await gsaGraphAll('networkAccess/filteringProfiles', { '$expand': 'policies($expand=policy),conditionalAccessPolicies' })) || [];
      let anyPass = false; const withCat = [];
      for (const p of wcf) {
        const detail = await Api.graph(`networkAccess/filteringPolicies/${p.id}`, { beta: true, query: { '$expand': 'policyRules' } }).catch(() => null);
        if (!toArray(detail?.policyRules).some(r => lower(r.ruleType) === 'webcategory')) continue;
        withCat.push(p.name);
        if (profilesLinkedToPolicy(profiles, p.id, 'filteringPolicyLink').some(l => l.passes)) anyPass = true;
      }
      const evidence = { wcfWithWebCategoryRules: withCat };
      if (anyPass) return pass(test, 'Web content filtering uses category-based rules applied through enabled, linked security profiles.', evidence);
      return fail(test, 'Web content filtering does not use category-based rules applied through enabled, linked profiles.', evidence);
    },

    // 25410 — WCF policies are linked to security profiles (partial).
    '25410': async (test) => {
      const [policies, profiles] = await Promise.all([gsaGraphAll('networkAccess/filteringPolicies', { '$expand': 'policyRules' }), gsaGraphAll('networkAccess/filteringProfiles', { '$expand': 'policies($expand=policy),conditionalAccessPolicies' })]);
      if (policies === null || profiles === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access filtering is not available.`);
      if (!toArray(policies).length) return fail(test, 'No web content filtering policies are configured.');
      const linked = [];
      for (const p of toArray(policies)) linked.push(...profilesLinkedToPolicy(profiles, p.id, 'filteringPolicyLink'));
      const baselineLinked = linked.some(l => l.isBaseline);
      const nonBaselineWithCA = linked.some(l => !l.isBaseline && l.hasCA);
      const evidence = { policies: toArray(policies).length, baselineLinked, nonBaselineWithCA };
      if (baselineLinked || nonBaselineWithCA) return pass(test, 'Web content filtering policies are linked to security profiles (baseline or a Conditional Access-linked profile).', evidence);
      return fail(test, 'Web content filtering policies are not linked to a baseline or Conditional Access-linked security profile.', evidence);
    },

    // 25411 — TLS inspection enabled and correctly configured (partial).
    '25411': async (test) => {
      const tls = await gsaGraphAll('networkAccess/tlsInspectionPolicies');
      if (tls === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access TLS inspection is not available.`);
      if (!tls.length) return fail(test, 'No TLS inspection policies are configured.');
      const profiles = (await gsaGraphAll('networkAccess/filteringProfiles', { '$expand': 'policies($expand=policy),conditionalAccessPolicies' })) || [];
      let applied = 0;
      for (const t of tls) if (profilesLinkedToPolicy(profiles, t.id, 'tlsInspectionPolicyLink').some(l => l.passes)) applied++;
      const evidence = { tlsPolicies: tls.length, appliedPolicies: applied };
      if (applied > 0) return pass(test, 'TLS inspection is enabled and applied through an enabled baseline or Conditional Access-linked profile.', evidence);
      return fail(test, 'TLS inspection policies exist but are not applied through an enabled linked profile.', evidence);
    },

    // 25413 — File transfer policies prevent data exfiltration (partial).
    '25413': async (test) => {
      const filePolicies = await gsaGraphAll('networkAccess/filePolicies');
      if (filePolicies === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access file policies are not available.`);
      if (!filePolicies.length) return fail(test, 'No file transfer policies are configured.');
      const profiles = (await gsaGraphAll('networkAccess/filteringProfiles', { '$expand': 'policies($expand=policy),conditionalAccessPolicies' })) || [];
      let applied = 0;
      for (const p of filePolicies) if (profilesLinkedToPolicy(profiles, p.id, 'filePolicyLink').some(l => l.passes)) applied++;
      const evidence = { filePolicies: filePolicies.length, appliedPolicies: applied };
      if (applied > 0) return pass(test, 'File transfer policies are configured and enforced through a filtering profile to prevent data exfiltration.', evidence);
      return fail(test, 'File transfer policies exist but are not enforced through an enabled linked profile.', evidence);
    },

    // 25415 — AI Gateway Prompt Shield protects generative AI apps (partial).
    '25415': async (test) => {
      const promptPolicies = await gsaGraphAll('networkAccess/promptPolicies', { '$expand': 'policyRules' });
      if (promptPolicies === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access AI Gateway / prompt policies are not available.`);
      if (!promptPolicies.length) return fail(test, 'No AI Gateway prompt-protection policies are configured.');
      const profiles = (await gsaGraphAll('networkAccess/filteringProfiles', { '$expand': 'policies($expand=policy),conditionalAccessPolicies' })) || [];
      let applied = 0;
      for (const p of promptPolicies) if (profilesLinkedToPolicy(profiles, p.id, 'promptPolicyLink').some(l => l.passes)) applied++;
      const evidence = { promptPolicies: promptPolicies.length, appliedPolicies: applied };
      if (applied > 0) return pass(test, 'AI Gateway Prompt Shield protection is enabled and applied through an enabled linked profile.', evidence);
      return fail(test, 'AI Gateway prompt policies exist but are not applied through an enabled linked profile.', evidence);
    },

    // 25416 — GSA cloud firewall protects branch office internet traffic.
    '25416': async (test) => {
      const branches = await gsaGraphAll('networkAccess/connectivity/branches');
      if (branches === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access connectivity is not available.`);
      if (!branches.length) return skip(test, `${skippedReason('NotApplicable')} No remote networks (branches) are configured.`);
      const profiles = (await gsaGraphAll('networkAccess/filteringProfiles', { '$expand': 'policies($expand=policy)' })) || [];
      const baseline = toArray(profiles).find(p => Number(p.priority) === 65000);
      const cfLinks = baseline ? toArray(baseline.policies).filter(l => lower(l['@odata.type'] || '').includes('cloudfirewallpolicylink')) : [];
      const enabledCf = cfLinks.filter(l => lower(l.state) === 'enabled');
      const evidence = { branches: branches.length, cloudFirewallLinks: cfLinks.length, enabledCloudFirewallLinks: enabledCf.length };
      if (enabledCf.length > 0) return pass(test, 'Branch office internet traffic is protected by enabled cloud firewall policies through the baseline security profile.', evidence);
      return fail(test, 'Branch office internet traffic is not protected by enabled cloud firewall policies in the baseline profile.', evidence);
    },
    // 25422 — GSA deployment logs populated with no recent failures.
    '25422': async (test) => {
      const profiles = await gsaGraphAll('networkAccess/forwardingProfiles');
      if (profiles === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access is not available.`);
      if (!toArray(profiles).some(p => lower(p.state) === 'enabled')) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access is not enabled.`);
      const deployments = (await gsaGraphAll('networkAccess/deployments')) || [];
      const cutoff = Date.now() - 30 * 86400000;
      const recent = deployments.filter(d => { const t = new Date(d.deploymentEndDateTime || d.lastModifiedDateTime || 0).getTime(); return t && t >= cutoff; });
      const failed = recent.filter(d => lower(d.status?.deploymentStage) === 'failed');
      const evidence = { recentDeployments: recent.length, failed: failed.length };
      if (failed.length === 0) return pass(test, 'Global Secure Access deployment logs are populated with no failed deployments in the last 30 days.', evidence);
      return fail(test, `${failed.length} Global Secure Access deployment(s) failed in the last 30 days.`, evidence);
    },

    // 25466 — At least two active connectors per Private Access connector group.
    '25466': async (test) => {
      const groups = await gsaGraphAll('onPremisesPublishingProfiles/applicationProxy/connectorGroups');
      if (groups === null) return skip(test, `${skippedReason('NotApplicable')} Application Proxy is not available.`);
      const appProxyGroups = toArray(groups).filter(g => lower(g.connectorGroupType) === 'applicationproxy');
      if (!appProxyGroups.length) return skip(test, `${skippedReason('NotApplicable')} No Private Access connector groups are configured.`);
      const belowTwo = [];
      for (const g of appProxyGroups) {
        const members = (await gsaGraphAll(`onPremisesPublishingProfiles/applicationProxy/connectorGroups/${g.id}/members`)) || [];
        const active = members.filter(m => lower(m.status) === 'active').length;
        if (active < 2) belowTwo.push({ name: g.name, activeConnectors: active });
      }
      const evidence = { connectorGroups: appProxyGroups.length, groupsBelowTwo: belowTwo };
      if (!belowTwo.length) return pass(test, 'Every Private Access connector group has at least two active, healthy connectors.', evidence);
      return fail(test, 'One or more Private Access connector groups have fewer than two active connectors.', evidence);
    },

    // 25480 — Quick Access has user or group assignments.
    '25480': async (test) => {
      const qa = (await Api.graphAll('servicePrincipals', { beta: true, query: { '$filter': "tags/any(c:c eq 'NetworkAccessQuickAccessApplication')", '$select': 'id,appId,displayName', '$top': '5' } }).catch(() => []))[0];
      if (!qa) return skip(test, `${skippedReason('NotApplicable')} No Quick Access application is configured.`);
      const app = await Api.graph(`servicePrincipals/${qa.id}`, { beta: true, query: { '$select': 'id,appId,appRoleAssignmentRequired', '$expand': 'appRoleAssignedTo($select=principalId,principalType,principalDisplayName)' } }).catch(() => null);
      if (!app) return skip(test, `${skippedReason('NotApplicable')} Unable to read Quick Access assignments.`);
      const hasAssignments = toArray(app.appRoleAssignedTo).length > 0;
      const evidence = { quickAccessApp: qa.displayName, assignmentRequired: app.appRoleAssignmentRequired, assignments: toArray(app.appRoleAssignedTo).length };
      if (!app.appRoleAssignmentRequired || hasAssignments) return pass(test, 'Quick Access has user or group assignments (or is assigned to all users).', evidence);
      return fail(test, 'Quick Access requires assignment but has no user or group assignments.', evidence);
    },

    // 25481 — All Private Access apps have user or group assignments.
    '25481': async (test) => {
      const raw = await Api.graphAll('servicePrincipals', { beta: true, query: { '$filter': "tags/any(c:c eq 'IsAccessibleViaZTNAClient')", '$expand': 'appRoleAssignedTo', '$select': 'id,appId,displayName,tags,appRoleAssignmentRequired' } }).catch(() => null);
      if (raw === null) return skip(test, `${skippedReason('NotApplicable')} Unable to enumerate Private Access applications.`);
      const apps = toArray(raw).filter(a => toArray(a.tags).includes('IsAccessibleViaZTNAClient'));
      if (!apps.length) return skip(test, `${skippedReason('NotApplicable')} No Private Access applications are configured.`);
      const without = apps.filter(a => !toArray(a.appRoleAssignedTo).length).map(a => a.displayName);
      const evidence = { privateAccessApps: apps.length, appsWithoutAssignments: without };
      if (!without.length) return pass(test, 'All Private Access applications have user or group assignments.', evidence);
      return fail(test, `${without.length} Private Access application(s) have no user or group assignments.`, evidence);
    },

    // 27000 — Web content filtering blocks high-risk categories (partial: category rule shape best-effort).
    '27000': async (test) => {
      const policies = await gsaGraphAll('networkAccess/filteringPolicies', { '$expand': 'policyRules' });
      if (policies === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access filtering is not available.`);
      const named = toArray(policies).filter(p => lower(p.name) !== 'all websites');
      if (!named.length) return fail(test, 'No web content filtering policies are configured to block high-risk categories.');
      const required = ['criminalactivity', 'hacking', 'illegalsoftware'];
      const profiles = (await gsaGraphAll('networkAccess/filteringProfiles', { '$expand': 'policies($expand=policy),conditionalAccessPolicies' })) || [];
      const blocked = new Set();
      for (const p of named) {
        if (!profilesLinkedToPolicy(profiles, p.id, 'filteringPolicyLink').some(l => l.passes)) continue;
        for (const r of toArray(p.policyRules)) {
          if (lower(r.ruleType) === 'webcategory' && lower(r.action) === 'block') {
            for (const c of toArray(r.destinations || r.webCategories || [])) blocked.add(lower(c?.name || c));
          }
        }
      }
      const missing = required.filter(c => !blocked.has(c));
      const evidence = { requiredCategories: ['CriminalActivity', 'Hacking', 'IllegalSoftware'], notBlocked: missing };
      if (!missing.length) return pass(test, 'High-risk web content categories (criminal activity, hacking, illegal software) are blocked and applied through a security profile.', evidence);
      return fail(test, 'One or more high-risk web content categories are not blocked and applied through a security profile.', evidence);
    },

    // 27001 — TLS inspection bypass rules are regularly reviewed (within 90 days).
    '27001': async (test) => {
      const tls = await gsaGraphAll('networkAccess/tlsInspectionPolicies', { '$expand': 'policyRules' });
      if (tls === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access TLS inspection is not available.`);
      if (!tls.length) return skip(test, `${skippedReason('NotApplicable')} No TLS inspection policies are configured.`);
      const flagged = [];
      for (const p of tls) {
        const bypass = toArray(p.policyRules).filter(r => !/^Auto-created TLS rule/.test(`${r.description || ''}`) && lower(r.action) === 'bypass');
        if (!bypass.length) continue;
        const days = Math.floor((Date.now() - new Date(p.lastModifiedDateTime || 0).getTime()) / 86400000);
        if (days > 90) flagged.push({ policy: p.name, daysSinceModified: days, customBypassRules: bypass.length });
      }
      const evidence = { policiesRequiringReview: flagged };
      if (!flagged.length) return pass(test, 'TLS inspection bypass rules have been reviewed within the last 90 days (or none exist).', evidence);
      return fail(test, 'One or more TLS inspection policies with custom bypass rules have not been reviewed in over 90 days.', evidence);
    },
    // 27002 — TLS inspection certificates have sufficient validity (>90 days).
    '27002': async (test) => {
      const tls = await gsaGraphAll('networkAccess/tlsInspectionPolicies');
      if (tls === null) return skip(test, `${skippedReason('NotApplicable')} Global Secure Access TLS inspection is not available.`);
      if (!tls.length) return skip(test, `${skippedReason('NotApplicable')} No TLS inspection policies are configured.`);
      const certs = await gsaGraphAll('networkAccess/tls/externalCertificateAuthorityCertificates');
      if (certs === null || !certs.length) return pass(test, 'TLS inspection is configured; no external CA certificates have been uploaded yet.');
      const now = Date.now();
      const flagged = [];
      for (const c of certs) {
        const st = lower(c.status);
        if (['csrgenerated', 'enrolling', 'disabled'].includes(st)) continue;
        const end = c.validity?.endDateTime ? new Date(c.validity.endDateTime).getTime() : null;
        const days = end != null ? Math.floor((end - now) / 86400000) : null;
        if (st === 'expiring' || st === 'expired' || ((st === 'active' || st === 'enabled') && days != null && days <= 90)) {
          flagged.push({ status: c.status, daysUntilExpiration: days });
        }
      }
      const evidence = { certificates: certs.length, flagged };
      if (!flagged.length) return pass(test, 'All active TLS inspection certificates have more than 90 days of validity remaining.', evidence);
      return fail(test, 'One or more TLS inspection certificates are expired, expiring, or within 90 days of expiration.', evidence);
    },

    // 61005 — Agents deployed to Microsoft 365 Copilot are discoverable in the Agent Registry.
    '61005': async (test) => {
      try {
        const pkgs = await Api.graphAll('copilot/admin/catalog/packages', { beta: true, query: { '$filter': "supportedHosts/any(h:h eq 'Copilot')" } });
        if (!toArray(pkgs).length) return fail(test, 'No agents are discoverable in the Microsoft 365 Agent Registry.', { agents: 0 });
        return pass(test, `${toArray(pkgs).length} agent(s) are discoverable in the Microsoft 365 Agent Registry.`, { agents: toArray(pkgs).length });
      } catch (err) {
        if ([403, 404].includes(err.status)) return skip(test, `${skippedReason('NotApplicable')} The tenant is not enrolled in the Microsoft 365 Copilot Frontier preview or lacks an Agent 365 license.`);
        throw err;
      }
    },

    // 61006 — AI administrative roles have assigned (tenant-scoped) principals.
    '61006': async (test) => {
      const roleDefs = await Api.graphAll('roleManagement/directory/roleDefinitions', { beta: true, query: { '$select': 'id,displayName,templateId', '$top': '500' } }).catch(() => null);
      if (roleDefs === null) return skip(test, `${skippedReason('NotApplicable')} Unable to read directory role definitions.`);
      const byId = new Set(toArray(roleDefs).map(r => lower(r.id)));
      const assignments = await Api.graphAll('roleManagement/directory/roleAssignments', { beta: true, query: { '$select': 'roleDefinitionId,directoryScopeId,principalId', '$top': '500' } }).catch(() => []);
      const tenantScopedByRole = new Map();
      for (const a of toArray(assignments)) {
        if (a.directoryScopeId !== '/') continue;
        const rid = lower(a.roleDefinitionId);
        tenantScopedByRole.set(rid, (tenantScopedByRole.get(rid) || 0) + 1);
      }
      const inScope = AI_ADMIN_ROLES.filter(r => byId.has(lower(r.id)));
      if (!inScope.length) return skip(test, `${skippedReason('NotApplicable')} None of the in-scope AI administrative roles are present in this tenant.`);
      const unassigned = inScope.filter(r => !(tenantScopedByRole.get(lower(r.id)) > 0)).map(r => ({ role: r.name, tier: r.tier }));
      const evidence = { evaluatedRoles: inScope.length, rolesWithoutTenantScopedPrincipal: unassigned };
      if (!unassigned.length) return pass(test, 'Every in-scope AI administrative role has at least one tenant-scoped assigned principal.', evidence);
      return fail(test, 'One or more AI administrative roles have no tenant-scoped assigned principal.', evidence);
    },

    // 61008 — Agent identities carry lifecycle tagging (customSecurityAttributes present).
    '61008': async (test) => {
      const agents = await getAgentIdentities();
      if (agents === null) return skip(test, `${skippedReason('NotApplicable')} Agent identities are not available (Agent 365 not licensed, or CustomSecAttributeAssignment.Read.All not consented).`);
      if (!agents.length) return skip(test, `${skippedReason('NotApplicable')} No agent identity service principals exist in this tenant.`);
      const hasCsa = (o) => o && typeof o === 'object' && Object.keys(o).length > 0;
      const without = agents.filter(a => !hasCsa(a.customSecurityAttributes)).map(a => a.displayName);
      const evidence = { agentIdentities: agents.length, withoutCustomSecurityAttributes: without };
      if (!without.length) return pass(test, 'All agent identities carry custom security attributes for lifecycle governance.', evidence);
      return fail(test, `${without.length} agent identity(ies) have no custom security attributes assigned.`, evidence);
    },

    // 61009 — Conditional Access covers agent identities and agent users (partial).
    '61009': async (test) => {
      const agents = await getAgentIdentities();
      if (agents === null || !agents.length) return skip(test, `${skippedReason('NotApplicable')} No agent identities are configured in this tenant.`);
      const policies = (await getConditionalAccessPolicies().catch(() => [])).filter(p => lower(p.state) === 'enabled');
      // Agent identity coverage: a workload-identity CA policy targeting service principals.
      const coversAgentIdentities = policies.some(p => {
        const ca = p.conditions?.clientApplications;
        return ca && (toArray(ca.includeServicePrincipals).length > 0 || lower(ca.servicePrincipalFilter?.rule || '').length > 0);
      });
      // Agent user coverage: a user CA policy targeting all users with a grant control.
      const coversAgentUsers = policies.some(p => toArray(p.conditions?.users?.includeUsers).includes('All') && (toArray(p.grantControls?.builtInControls).length > 0 || p.grantControls?.authenticationStrength));
      const evidence = { agentIdentities: agents.length, coversAgentIdentities, coversAgentUsers };
      if (coversAgentIdentities && coversAgentUsers) return pass(test, 'Conditional Access policies cover both agent identities (workload identity policy) and agent users.', evidence);
      return fail(test, 'Conditional Access does not cover both agent identities and agent users.', evidence);
    },

    // 61011 — Agent identity sign-in activity is monitored (partial; agent-gated).
    '61011': async (test) => {
      const agents = await getAgentIdentities();
      if (agents === null || !agents.length) return skip(test, `${skippedReason('NotApplicable')} No agent identities are configured in this tenant.`);
      const ids = new Set(agents.map(a => lower(a.appId)).filter(Boolean));
      let signIns = null;
      try {
        signIns = await Api.graphAll('auditLogs/signIns', { beta: true, query: { '$filter': "signInEventTypes/any(t: t eq 'servicePrincipal')", '$top': '200' } });
      } catch (err) { if (![400, 403, 404].includes(err.status)) throw err; }
      if (signIns === null) return skip(test, `${skippedReason('NotApplicable')} Service principal sign-in logs are not available.`);
      const active = new Set(toArray(signIns).map(s => lower(s.appId)).filter(x => ids.has(x)));
      const inactive = agents.filter(a => !active.has(lower(a.appId))).map(a => a.displayName);
      const evidence = { agentIdentities: agents.length, withRecentSignIn: active.size, withoutRecentSignIn: inactive.slice(0, 20) };
      if (!inactive.length) return pass(test, 'All agent identities show recent sign-in activity in the monitored window.', evidence);
      return fail(test, `${inactive.length} agent identity(ies) show no recent sign-in activity — review for stale or unmonitored agents.`, evidence);
    },

    // 61012 — Risk-based Conditional Access blocks risky agent identities (partial).
    '61012': async (test) => {
      const agents = await getAgentIdentities();
      if (agents === null || !agents.length) return skip(test, `${skippedReason('NotApplicable')} No agent identities are configured in this tenant.`);
      const policies = (await getConditionalAccessPolicies().catch(() => [])).filter(p => lower(p.state) === 'enabled');
      const riskyWorkloadPolicy = policies.some(p => {
        const ca = p.conditions?.clientApplications;
        const targetsSp = ca && (toArray(ca.includeServicePrincipals).length > 0 || lower(ca.servicePrincipalFilter?.rule || '').length > 0);
        const usesRisk = toArray(p.conditions?.servicePrincipalRiskLevels).length > 0;
        const blocks = toArray(p.grantControls?.builtInControls).map(lower).includes('block');
        return targetsSp && usesRisk && blocks;
      });
      const evidence = { agentIdentities: agents.length, riskBasedWorkloadBlockPolicy: riskyWorkloadPolicy };
      if (riskyWorkloadPolicy) return pass(test, 'A risk-based Conditional Access policy blocks risky agent (workload) identities.', evidence);
      return fail(test, 'No risk-based Conditional Access policy blocks risky agent identities.', evidence);
    },

    // 61013 — Identity governance for agents: sponsors, entitlement channel, lifecycle automation (partial).
    '61013': async (test) => {
      const agents = await getAgentIdentities();
      const blueprints = await Api.graphAll('applications/microsoft.graph.agentIdentityBlueprint', { beta: true, query: { '$select': 'id,displayName' } }).catch(() => null);
      if ((agents === null || !agents.length) && (blueprints === null || !toArray(blueprints).length)) {
        return skip(test, `${skippedReason('NotApplicable')} No agent identities or blueprints exist in this tenant.`);
      }
      const [assignmentPolicies, workflows] = await Promise.all([
        Api.graphAll('identityGovernance/entitlementManagement/assignmentPolicies', { query: { '$select': 'id,displayName', '$top': '100' } }).catch(() => null),
        Api.graphAll('identityGovernance/lifecycleWorkflows/workflows', { query: { '$select': 'id,displayName,category', '$top': '100' } }).catch(() => null),
      ]);
      if (assignmentPolicies === null || workflows === null) return skip(test, `${skippedReason('NotLicensedEntraIDGovernance')}`);
      const hasEntitlementChannel = toArray(assignmentPolicies).length > 0;
      const hasLifecycleAutomation = toArray(workflows).length > 0;
      const evidence = { hasEntitlementChannel, hasLifecycleAutomation };
      if (hasEntitlementChannel && hasLifecycleAutomation) return pass(test, 'An entitlement-management channel and lifecycle-automation workflows exist to govern agent access.', evidence);
      return fail(test, 'Identity governance for agents is incomplete: an entitlement-management channel and/or lifecycle-automation workflows are missing.', evidence);
    },

    // 61014 — Agent identities/blueprint principals have owners and no disabled agents remain (partial).
    '61014': async (test) => {
      const agents = await getAgentIdentities();
      if (agents === null) return skip(test, `${skippedReason('NotApplicable')} Agent identities are not available in this tenant.`);
      if (!agents.length) return skip(test, `${skippedReason('NotApplicable')} No agent identities exist in this tenant.`);
      const ownerless = [], disabled = [];
      for (const a of agents) {
        if (a.accountEnabled === false) disabled.push(a.displayName);
        const owners = await Api.graphAll(`servicePrincipals/${a.id}/owners`, { query: { '$select': 'id', '$top': '5' } }).catch(() => []);
        if (!toArray(owners).length) ownerless.push(a.displayName);
      }
      const evidence = { agentIdentities: agents.length, ownerless, disabled };
      if (!ownerless.length && !disabled.length) return pass(test, 'All agent identities have technical owners and no disabled agents remain in the directory.', evidence);
      return fail(test, 'One or more agent identities lack a technical owner or remain in the directory while disabled.', evidence);
    },
    // 24546 — Windows automatic device enrollment (MDM) is enforced.
    '24546': async (test) => {
      const policies = await intuneGraphAll('policies/mobileDeviceManagementPolicies');
      if (policies === null) return skip(test, `${skippedReason('NotLicensedEntraIDP1')}`);
      const intune = toArray(policies).find(p => p.displayName === 'Microsoft Intune');
      const evidence = { intuneMdmScope: intune?.appliesTo || 'none' };
      if (intune && lower(intune.appliesTo) !== 'none') return pass(test, 'Windows automatic device enrollment (Microsoft Intune MDM) is enforced.', evidence);
      return fail(test, 'Windows automatic device enrollment is not enforced (Microsoft Intune MDM user scope is "none").', evidence);
    },

    // 24561 — macOS LAPS protects local admin credentials during enrollment.
    '24561': async (test) => {
      const depTokens = await intuneGraphAll('deviceManagement/depOnboardingSettings', { '$expand': 'enrollmentProfiles', '$select': 'id,appleIdentifier,tokenName' });
      if (depTokens === null) return skip(test, `${skippedReason('NotApplicable')} Apple DEP onboarding settings are not available.`);
      if (!depTokens.length) return fail(test, 'No Apple Automated Device Enrollment (DEP) tokens are configured.');
      let profilesWithLaps = 0, profilesWithAssignments = 0, total = 0;
      for (const token of depTokens) {
        const profiles = (await intuneGraphAll(`deviceManagement/depOnboardingSettings/${token.id}/enrollmentProfiles`)) || [];
        for (const p of profiles) {
          total++;
          if (`${p.adminAccountUserName || ''}`.trim()) {
            profilesWithLaps++;
            const profileId = `${p.id || ''}`.split('_')[1] || p.id;
            const assignments = (await intuneGraphAll(`deviceManagement/depOnboardingSettings/${token.id}/importedAppleDeviceIdentities`, { '$top': '5', '$filter': `discoverySource eq 'deviceEnrollmentProgram' and requestedEnrollmentProfileId eq '${profileId}'` })) || [];
            if (assignments.length) profilesWithAssignments++;
          }
        }
      }
      const evidence = { enrollmentProfiles: total, profilesWithLaps, profilesWithAssignments };
      if (total && profilesWithLaps && profilesWithAssignments) return pass(test, 'A macOS LAPS-enabled enrollment profile is configured and assigned.', evidence);
      return fail(test, 'No assigned macOS enrollment profile protects local administrator credentials with LAPS.', evidence);
    },

    // 24570 — Entra Connect Sync uses service principal credentials (no enabled sync user accounts).
    '24570': async (test) => {
      const org = (await Api.graphAll('organization', { query: { '$select': 'onPremisesSyncEnabled,onPremisesLastSyncDateTime' } }).catch(() => []))[0];
      const isHybrid = !!org?.onPremisesSyncEnabled;
      const dirSyncRole = (await Api.graphAll('directoryRoles', { query: { '$filter': "roleTemplateId eq 'd29b2b05-8046-44ba-8758-1e26182fcf32'", '$expand': 'members($select=id,displayName,userPrincipalName,accountEnabled,userType)' } }).catch(() => []))[0];
      const members = toArray(dirSyncRole?.members);
      const enabledUsers = members.filter(m => m.accountEnabled === true && lower(m['@odata.type']) === '#microsoft.graph.user');
      const evidence = { hybridIdentity: isHybrid, enabledDirSyncUserAccounts: enabledUsers.length };
      if (!isHybrid || enabledUsers.length === 0) return pass(test, isHybrid ? 'Entra Connect Sync uses service principal credentials (no enabled user accounts in the Directory Synchronization Accounts role).' : 'Tenant is not using hybrid identity synchronization.', evidence);
      return fail(test, 'Enabled user accounts are present in the Directory Synchronization Accounts role — Entra Connect Sync is not using service principal credentials.', evidence);
    },

    // 24823 — Company Portal branding and support settings are customized.
    '24823': async (test) => {
      const profiles = await intuneGraphAll('deviceManagement/intuneBrandingProfiles', { '$select': 'id,isDefaultProfile,profileName,displayName,contactITPhoneNumber,contactITEmailAddress' });
      if (profiles === null) return skip(test, `${skippedReason('NotLicensedIntune')}`);
      const complete = (p) => !!(p.displayName && p.contactITPhoneNumber && p.contactITEmailAddress);
      const def = toArray(profiles).find(p => p.isDefaultProfile);
      const nonDefault = toArray(profiles).filter(p => !p.isDefaultProfile);
      let nonDefaultAssignedComplete = false;
      for (const p of nonDefault.filter(complete)) {
        const assignments = (await intuneGraphAll(`deviceManagement/intuneBrandingProfiles/${p.id}/assignments`)) || [];
        if (assignments.length) { nonDefaultAssignedComplete = true; break; }
      }
      const evidence = { defaultProfileComplete: !!(def && complete(def)), nonDefaultAssignedComplete };
      if ((def && complete(def)) || nonDefaultAssignedComplete) return pass(test, 'Company Portal branding includes display name and IT support contact details.', evidence);
      return fail(test, 'Company Portal branding is missing display name or IT support contact details.', evidence);
    },

    // 24824 — Conditional Access blocks noncompliant devices (all platforms or unfiltered).
    '24824': async (test) => {
      const pols = await Api.graphAll('identity/conditionalAccess/policies', { beta: true, query: { '$filter': "state eq 'enabled' and grantControls/builtInControls/any(bc: bc eq 'compliantDevice')", '$select': 'id,displayName,grantControls,conditions' } }).catch(() => null);
      if (pols === null) return skip(test, `${skippedReason('NotApplicable')} Unable to read Conditional Access policies.`);
      const has = (plat) => toArray(pols).some(p => toArray(p.conditions?.platforms?.includePlatforms).map(lower).includes(plat));
      const unfiltered = toArray(pols).some(p => !p.conditions?.platforms?.includePlatforms);
      const passed = unfiltered || (has('android') && has('ios') && has('macos') && has('windows'));
      const evidence = { compliantDevicePolicies: toArray(pols).length, unfiltered, platforms: { android: has('android'), ios: has('ios'), macos: has('macos'), windows: has('windows') } };
      if (passed) return pass(test, 'Conditional Access policies block access from noncompliant devices across all platforms.', evidence);
      return fail(test, 'Conditional Access does not block noncompliant devices across all platforms.', evidence);
    },

    // 24827 — Conditional Access blocks unmanaged apps (compliantApplication for iOS/Android).
    '24827': async (test) => {
      const pols = await Api.graphAll('identity/conditionalAccess/policies', { beta: true, query: { '$filter': "state eq 'enabled' and grantControls/builtInControls/any(bc: bc eq 'compliantApplication')", '$select': 'id,displayName,grantControls,conditions' } }).catch(() => null);
      if (pols === null) return skip(test, `${skippedReason('NotApplicable')} Unable to read Conditional Access policies.`);
      const has = (plat) => toArray(pols).some(p => toArray(p.conditions?.platforms?.includePlatforms).map(lower).includes(plat));
      const unfiltered = toArray(pols).some(p => !p.conditions?.platforms?.includePlatforms);
      const passed = unfiltered || (has('android') && has('ios'));
      const evidence = { compliantAppPolicies: toArray(pols).length, unfiltered, platforms: { android: has('android'), ios: has('ios') } };
      if (passed) return pass(test, 'Conditional Access policies block access from unmanaged apps on iOS and Android.', evidence);
      return fail(test, 'Conditional Access does not require approved/managed apps on iOS and Android.', evidence);
    },

    // 24839 / 24840 / 24870 — Secure (enterprise) Wi-Fi profiles for iOS / Android / macOS.
    '24839': async (test) => wifiProfileCheck(test, 'ios', 'iOS'),
    '24840': async (test) => wifiProfileCheck(test, 'android', 'Android'),
    '24870': async (test) => wifiProfileCheck(test, 'macos', 'macOS'),

    // 24871 — Defender for Endpoint automatic enrollment enforced for Android.
    '24871': async (test) => {
      const connectors = await intuneGraphAll('deviceManagement/mobileThreatDefenseConnectors');
      if (connectors === null) return skip(test, `${skippedReason('NotLicensedIntune')}`);
      const defender = toArray(connectors).find(c => lower(c.id) === 'fc780465-2017-40d4-a0c5-307022471b92');
      const evidence = { partnerState: defender?.partnerState || null, androidEnabled: defender?.androidEnabled ?? null };
      if (defender && lower(defender.partnerState) === 'enabled' && defender.androidEnabled === true) return pass(test, 'Defender for Endpoint automatic enrollment is enforced for Android devices.', evidence);
      return fail(test, 'Defender for Endpoint automatic enrollment is not enforced for Android devices.', evidence);
    },
  };

  async function mamPolicies(kind) {
    const path = kind === 'ios' ? 'deviceAppManagement/iosManagedAppProtections' : 'deviceAppManagement/androidManagedAppProtections';
    return intuneGraphAll(path, { '$expand': 'assignments' });
  }
  const mamAssigned = (p) => p.isAssigned === true || toArray(p.assignments).length > 0;

  ZtaImpl.register({
    // 24518 — Enterprise applications (with Graph permissions) have >=2 owners (partial: privilege proxy).
    '24518': async (test) => {
      const apps = await Api.graphAll('applications', { query: { '$select': 'id,appId,displayName,requiredResourceAccess,signInAudience', '$top': '999' } }).catch(() => null);
      if (apps === null) return skip(test, `${skippedReason('NotApplicable')} Unable to enumerate applications.`);
      const graphAppId = '00000003-0000-0000-c000-000000000000';
      const inScope = toArray(apps).filter(a => toArray(a.requiredResourceAccess).some(r => lower(r.resourceAppId) === graphAppId));
      if (!inScope.length) return pass(test, 'No applications with Microsoft Graph permissions require owner review.');
      const insufficient = [];
      for (const a of inScope) {
        const owners = await Api.graphAll(`applications/${a.id}/owners`, { query: { '$select': 'id', '$top': '5' } }).catch(() => []);
        if (toArray(owners).length < 2) insufficient.push({ displayName: a.displayName, owners: toArray(owners).length });
      }
      const evidence = { inScopeApps: inScope.length, appsWithFewerThanTwoOwners: insufficient.slice(0, 30) };
      if (!insufficient.length) return pass(test, 'All in-scope enterprise applications have at least two owners.', evidence);
      return fail(test, `${insufficient.length} application(s) have fewer than two owners.`, evidence);
    },

    // 35001 — Conditional Access policies don't exclude Rights Management workloads.
    '35001': async (test) => {
      const policies = (await getConditionalAccessPolicies().catch(() => null));
      if (policies === null) return skip(test, `${skippedReason('NotApplicable')} Unable to read Conditional Access policies.`);
      const rmsAppId = '00000012-0000-0000-c000-000000000000';
      const blocking = toArray(policies).filter(p => lower(p.state) === 'enabled').filter(p => {
        const inc = toArray(p.conditions?.applications?.includeApplications);
        const exc = toArray(p.conditions?.applications?.excludeApplications);
        return (inc.includes('All') || inc.includes(rmsAppId)) && !exc.includes(rmsAppId);
      }).map(p => ({ id: p.id, displayName: p.displayName }));
      const evidence = { policiesTargetingRmsWithoutExclusion: blocking };
      if (!blocking.length) return pass(test, "Conditional Access policies don't inadvertently include Rights Management workloads without excluding them.", evidence);
      return fail(test, `${blocking.length} enabled Conditional Access policy(ies) target Rights Management workloads without excluding them.`, evidence);
    },

    // 41003 — Microsoft Defender for Identity sensors healthy on AD FS/AD CS/Entra Connect (partial).
    '41003': async (test) => {
      let sensors = null;
      try { sensors = await Api.graphAll('security/identities/sensors', { beta: true }); }
      catch (err) { if ([403, 404].includes(err.status)) return skip(test, `${skippedReason('NotApplicable')} Microsoft Defender for Identity is not onboarded or not accessible.`); throw err; }
      if (!toArray(sensors).length) return skip(test, `${skippedReason('NotApplicable')} No Microsoft Defender for Identity sensors are registered.`);
      const unhealthy = toArray(sensors).filter(s => !['running', 'healthy', 'upToDate'].includes(s.healthStatus) && lower(s.healthStatus || '') !== 'healthy').map(s => ({ sensor: s.displayName || s.id, health: s.healthStatus }));
      const evidence = { sensors: toArray(sensors).length, unhealthy };
      if (!unhealthy.length) return pass(test, 'Microsoft Defender for Identity sensors are installed and healthy.', evidence);
      return fail(test, 'One or more Microsoft Defender for Identity sensors are unhealthy.', evidence);
    },

    // 41018 — No open Microsoft Defender for Identity health issues.
    '41018': async (test) => {
      let issues = null;
      try { issues = await Api.graphAll('security/identities/healthIssues', { beta: true, query: { '$filter': "status eq 'open'" } }); }
      catch (err) { if ([403, 404].includes(err.status)) return skip(test, `${skippedReason('NotApplicable')} Microsoft Defender for Identity is not deployed or not accessible.`); throw err; }
      const critical = toArray(issues).filter(i => ['medium', 'high'].includes(lower(i.severity)));
      const evidence = { openIssues: toArray(issues).length, mediumHigh: critical.length };
      if (!critical.length) return pass(test, 'No open medium/high Microsoft Defender for Identity health issues are present.', evidence);
      return fail(test, `${critical.length} open medium/high Microsoft Defender for Identity health issue(s) are present.`, evidence);
    },

    // 51001 — Windows Endpoint Privilege Management configured and assigned.
    '51001': async (test) => {
      let skus = null;
      try { skus = await Api.graphAll('subscribedSkus'); } catch (err) { if (![400, 403, 404].includes(err.status)) throw err; }
      if (skus === null) return skip(test, `${skippedReason('NotApplicable')} Unable to determine Intune Suite (EPM) licensing.`);
      const epmLicensed = toArray(skus).some(s => lower(s.capabilityStatus) === 'enabled' && toArray(s.servicePlans).some(p => p.servicePlanName === 'Intune-EPM' && lower(p.provisioningStatus) === 'success'));
      if (!epmLicensed) return fail(test, 'The tenant has no active Intune Suite (Intune-EPM) license — Endpoint Privilege Management cannot be enabled.');
      const policies = await intuneGraphAll('deviceManagement/configurationPolicies', { '$expand': 'assignments', '$top': '200' });
      if (policies === null) return skip(test, `${skippedReason('NotLicensedIntune')}`);
      const epm = toArray(policies).filter(p => lower(p.templateReference?.templateFamily) === 'endpointsecurityendpointprivilegemanagement');
      const dn = (p) => lower(p.templateReference?.templateDisplayName || '');
      const assigned = (p) => toArray(p.assignments).length > 0;
      const settingsAssigned = epm.some(p => dn(p).includes('elevation settings') && assigned(p));
      const rulesAssigned = epm.some(p => dn(p).includes('elevation rules') && assigned(p));
      const evidence = { epmPolicies: epm.length, settingsAssigned, rulesAssigned };
      if (settingsAssigned && rulesAssigned) return pass(test, 'Windows Endpoint Privilege Management elevation settings and rules policies are configured and assigned.', evidence);
      return fail(test, 'Windows Endpoint Privilege Management is not fully configured (an assigned elevation settings and elevation rules policy are both required).', evidence);
    },

    // 51013 — App Protection Policies block managed-app access when re-authentication fails (partial).
    '51013': async (test) => {
      const [ios, android] = await Promise.all([mamPolicies('ios'), mamPolicies('android')]);
      if (ios === null && android === null) return skip(test, `${skippedReason('NotLicensedIntune')}`);
      const acceptable = new Set(['block', 'wipe']);
      const okIos = toArray(ios).some(p => mamAssigned(p) && acceptable.has(lower(p.appActionIfUnableToAuthenticateUser)));
      const okAndroid = toArray(android).some(p => mamAssigned(p) && acceptable.has(lower(p.appActionIfUnableToAuthenticateUser)));
      const iosInScope = toArray(ios).length > 0, androidInScope = toArray(android).length > 0;
      const evidence = { iosPolicies: toArray(ios).length, androidPolicies: toArray(android).length, iosEnforced: okIos, androidEnforced: okAndroid };
      const iosPass = !iosInScope || okIos, androidPass = !androidInScope || okAndroid;
      if (iosPass && androidPass && (iosInScope || androidInScope)) return pass(test, 'App Protection Policies block managed-app access when user re-authentication fails.', evidence);
      if (!iosInScope && !androidInScope) return skip(test, `${skippedReason('NotApplicable')} No App Protection Policies are configured.`);
      return fail(test, 'One or more platforms lack an assigned App Protection Policy that blocks access when re-authentication fails.', evidence);
    },

    // 51014 — App Protection Policies block on jailbroken/rooted devices.
    '51014': async (test) => {
      const [ios, android] = await Promise.all([mamPolicies('ios'), mamPolicies('android')]);
      if (ios === null && android === null) return skip(test, `${skippedReason('NotLicensedIntune')}`);
      const acceptable = new Set(['block', 'wipe']);
      const attest = new Set(['basicintegrity', 'basicintegrityanddevicecertification']);
      const okIos = toArray(ios).some(p => mamAssigned(p) && p.deviceComplianceRequired === true && acceptable.has(lower(p.appActionIfDeviceComplianceRequired)));
      const okAndroid = toArray(android).some(p => mamAssigned(p) && attest.has(lower(p.requiredAndroidSafetyNetDeviceAttestationType)) && acceptable.has(lower(p.appActionIfAndroidSafetyNetDeviceAttestationFailed)));
      const iosInScope = toArray(ios).length > 0, androidInScope = toArray(android).length > 0;
      const evidence = { iosEnforced: okIos, androidEnforced: okAndroid };
      const iosPass = !iosInScope || okIos, androidPass = !androidInScope || okAndroid;
      if (!iosInScope && !androidInScope) return skip(test, `${skippedReason('NotApplicable')} No App Protection Policies are configured.`);
      if (iosPass && androidPass) return pass(test, 'App Protection Policies block managed-app access on jailbroken or rooted devices.', evidence);
      return fail(test, 'One or more platforms lack an assigned App Protection Policy enforcing jailbreak/root protection.', evidence);
    },

    // 51015 — Multi Admin Approval is enabled in Intune.
    '51015': async (test) => {
      const policies = await intuneGraphAll('deviceManagement/operationApprovalPolicies', { '$select': 'id,displayName,policyType,policyPlatform' });
      if (policies === null) return skip(test, `${skippedReason('NotLicensedIntune')}`);
      const evidence = { approvalPolicies: toArray(policies).length };
      if (toArray(policies).length > 0) return pass(test, 'Intune Multi Admin Approval access policies are configured.', evidence);
      return fail(test, 'No Intune Multi Admin Approval access policies are configured.', evidence);
    },

    // 51018 — Windows App Control for Business (WDAC) configured and assigned (partial).
    '51018': async (test) => {
      const policies = await intuneGraphAll('deviceManagement/configurationPolicies', { '$expand': 'assignments', '$top': '200' });
      if (policies === null) return skip(test, `${skippedReason('NotLicensedIntune')}`);
      const appControl = toArray(policies).filter(p => {
        const t = `${lower(p.templateReference?.templateDisplayName || '')} ${lower(p.name || '')}`;
        return t.includes('app control') || t.includes('application control');
      });
      const assigned = appControl.filter(p => toArray(p.assignments).length > 0);
      const evidence = { appControlPolicies: appControl.length, assigned: assigned.length };
      if (assigned.length > 0) return pass(test, 'Windows App Control for Business (WDAC) policies are configured and assigned via Intune.', evidence);
      return fail(test, 'No assigned Windows App Control for Business (WDAC) policy is configured in Intune.', evidence);
    },

    // 51019 — MAM app access timeout / re-auth enforced after idle period.
    '51019': async (test) => {
      const [ios, android] = await Promise.all([mamPolicies('ios'), mamPolicies('android')]);
      if (ios === null && android === null) return skip(test, `${skippedReason('NotLicensedIntune')}`);
      const durationMinutes = (iso) => { const m = `${iso || ''}`.match(/PT(?:(\d+)H)?(?:(\d+)M)?/); if (!m) return null; return (Number(m[1] || 0) * 60) + Number(m[2] || 0); };
      const qualifies = (p) => { const mins = durationMinutes(p.periodOnlineBeforeAccessCheck); return mamAssigned(p) && mins != null && mins <= 30 && (p.pinRequired === true || p.organizationalCredentialsRequired === true); };
      const okIos = toArray(ios).some(qualifies);
      const okAndroid = toArray(android).some(qualifies);
      const iosInScope = toArray(ios).length > 0, androidInScope = toArray(android).length > 0;
      const evidence = { iosEnforced: okIos, androidEnforced: okAndroid };
      if (!iosInScope && !androidInScope) return skip(test, `${skippedReason('NotApplicable')} No App Protection Policies are configured.`);
      if ((!iosInScope || okIos) && (!androidInScope || okAndroid)) return pass(test, 'MAM app access timeout and re-authentication are enforced after an idle period (<=30 min).', evidence);
      return fail(test, 'One or more platforms lack an assigned App Protection Policy enforcing a <=30 minute access recheck with PIN/credentials.', evidence);
    },
  });

  async function wifiProfileCheck(test, platformMatch, label) {
    const profiles = await wifiProfilesForPlatform(platformMatch);
    if (profiles === null) return skip(test, `${skippedReason('NotLicensedIntune')}`);
    const enterprise = new Set(['wpa2enterprise', 'wpaenterprise']);
    const compliant = profiles.filter(p => enterprise.has(wifiSecType(p)) && toArray(p.assignments).length > 0);
    const evidence = { wifiProfiles: profiles.length, compliantAssignedProfiles: compliant.length };
    if (compliant.length > 0) return pass(test, `Secure (enterprise) Wi-Fi profiles protect ${label} devices and are assigned.`, evidence);
    return fail(test, `No assigned enterprise-secured Wi-Fi profile protects ${label} devices from unauthorized network access.`, evidence);
  }

  ZtaImpl.register(impl);
  // Expose shared GSA helpers for later network controls appended to this file.
  window.ZtaGraphLib = { gsaGraph, gsaGraphAll, isoZ, profilesLinkedToPolicy, getConditionalAccessPolicies };
})();
