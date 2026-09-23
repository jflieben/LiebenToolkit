/*
 * Phase 3 — Bucket B: Azure controls via Azure Resource Manager / Resource Graph.
 *
 * Ports the KQL Resource Graph queries and pass/fail logic of the ZeroTrustAssessment
 * Azure tests to the browser using a delegated ARM token (Api.armResourceGraph).
 * Skip semantics mirror the module: no ARM consent -> NotConnectedAzure; ARG error ->
 * NotSupported; zero rows -> NotApplicable.
 */
(() => {
  const { pass, fail, skip, skippedReason, toArray, lower } = window.ZtaLib;

  // Runs an ARG query with the module's skip semantics, then hands rows to `evaluate`.
  async function runArg(test, query, evaluate, emptyMsg) {
    let rows;
    try {
      rows = await Api.armResourceGraph(query);
    } catch (err) {
      if (err.armConsent || [401, 403].includes(err.status)) return skip(test, skippedReason('NotConnectedAzure'));
      if ([400, 404].includes(err.status)) return skip(test, `${skippedReason('NotSupported')} (Azure Resource Graph query failed).`);
      throw err;
    }
    if (!toArray(rows).length) return skip(test, `${skippedReason('NotApplicable')} ${emptyMsg || ''}`.trim());
    return evaluate(toArray(rows));
  }

  // Whether a resource has any enabled diagnostic setting of the requested kind.
  async function anyDiagnostic(resourceId, kind) {
    try {
      const resp = await Api.arm(`${resourceId}/providers/microsoft.insights/diagnosticSettings?api-version=2021-05-01-preview`);
      const settings = toArray(resp?.value);
      const logsOn = settings.some(s => toArray(s.properties?.logs).some(l => l.enabled));
      const metricsOn = settings.some(s => toArray(s.properties?.metrics).some(m => m.enabled));
      if (kind === 'logs') return logsOn;
      if (kind === 'metrics') return metricsOn;
      return logsOn || metricsOn;
    } catch { return false; }
  }
  // Evaluates diagnostic coverage across the given rows (partial: category granularity not checked).
  async function diagCheck(test, rows, idField, nameField, kind, label) {
    const missing = [];
    for (const r of rows) { if (!(await anyDiagnostic(r[idField], kind))) missing.push(r[nameField] || r[idField]); }
    const noun = kind === 'metrics' ? 'metrics' : 'logging';
    if (!missing.length) return pass(test, `Diagnostic ${noun} is enabled for all ${label}.`, { resources: rows.length });
    return fail(test, `${missing.length} ${label} lack diagnostic ${noun}.`, { resources: rows.length, missing: missing.slice(0, 20) });
  }

  // Sentinel-onboarded Log Analytics workspaces (best-effort; used by AI/SIEM controls).
  async function sentinelWorkspaces() {
    let ws;
    try { ws = await Api.armResourceGraph(`resources | where type =~ 'microsoft.operationalinsights/workspaces' | project id, name, subscriptionId`); }
    catch (err) { if (err.armConsent || [401, 403].includes(err.status)) return 'noaccess'; if ([400, 404].includes(err.status)) return 'unsupported'; throw err; }
    const onboarded = [];
    for (const w of toArray(ws)) {
      try { const ob = await Api.arm(`${w.id}/providers/Microsoft.SecurityInsights/onboardingStates/default?api-version=2023-11-01`); if (ob) onboarded.push(w); }
      catch { /* not onboarded / forbidden */ }
    }
    return { workspaces: toArray(ws), onboarded };
  }
  // Best-effort: does any Sentinel workspace have a data connector matching `keyword`?
  async function sentinelHasConnector(onboarded, keyword) {
    for (const w of onboarded) {
      let connectors = [];
      try { const r = await Api.arm(`${w.id}/providers/Microsoft.SecurityInsights/dataConnectors?api-version=2023-11-01`); connectors = toArray(r?.value); }
      catch { continue; }
      if (connectors.some(c => JSON.stringify(c).toLowerCase().includes(keyword))) return true;
    }
    return false;
  }
  async function connectorTest(test, keyword, label) {
    const sw = await sentinelWorkspaces();
    if (sw === 'noaccess') return skip(test, skippedReason('NotConnectedAzure'));
    if (sw === 'unsupported') return skip(test, `${skippedReason('NotSupported')}`);
    if (!sw.workspaces.length) return skip(test, `${skippedReason('NotApplicable')} No Log Analytics workspaces were found.`);
    if (!sw.onboarded.length) return fail(test, 'Microsoft Sentinel is not onboarded on any Log Analytics workspace.', { workspaces: sw.workspaces.length });
    const found = await sentinelHasConnector(sw.onboarded, keyword);
    const evidence = { sentinelWorkspaces: sw.onboarded.length, connectorEnabled: found };
    if (found) return pass(test, `The ${label} data connector is enabled on a Microsoft Sentinel workspace.`, evidence);
    return fail(test, `The ${label} data connector is not enabled on any Microsoft Sentinel workspace.`, evidence);
  }

  const impl = {
    // 26879 — Request Body Inspection enabled in Application Gateway WAF.
    '26879': async (test) => runArg(test,
      `resources
| where type =~ 'microsoft.network/ApplicationGatewayWebApplicationFirewallPolicies'
| extend wafPolicyId = tolower(id)
| join kind=inner (
    resources
    | where type =~ 'microsoft.network/applicationgateways'
    | where isnotempty(properties.firewallPolicy.id)
    | extend wafPolicyId = tolower(tostring(properties.firewallPolicy.id))
    | project wafPolicyId, GatewayName=name
) on wafPolicyId
| summarize PolicyName=any(name), RequestBodyCheck=any(tobool(properties.policySettings.requestBodyCheck)), EnabledState=any(tostring(properties.policySettings.state)), Mode=any(tostring(properties.policySettings.mode)) by wafPolicyId`,
      (rows) => {
        const bad = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.Mode) !== 'prevention' || r.RequestBodyCheck !== true);
        if (!bad.length) return pass(test, 'Request Body Inspection is enabled on all Application Gateway WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${bad.length} Application Gateway WAF policy(ies) are not in Prevention mode with request body inspection enabled.`, { policies: rows.length, nonCompliant: bad.map(r => ({ policy: r.PolicyName, state: r.EnabledState, mode: r.Mode, requestBodyCheck: r.RequestBodyCheck })) });
      },
      'No Application Gateway WAF policies attached to Application Gateways were found.'),

    // 26880 — Request Body Inspection enabled in Azure Front Door WAF.
    '26880': async (test) => runArg(test,
      `resources
| where type =~ 'microsoft.network/frontdoorwebapplicationfirewallpolicies'
| where isnotempty(properties.frontendEndpointLinks) or isnotempty(properties.securityPolicyLinks)
| project PolicyName=name, subscriptionId, RequestBodyCheck=tostring(properties.policySettings.requestBodyCheck), EnabledState=tostring(properties.policySettings.enabledState), Mode=tostring(properties.policySettings.mode)`,
      (rows) => {
        const bad = rows.filter(r => !(lower(r.RequestBodyCheck) === 'enabled' && lower(r.EnabledState) === 'enabled' && lower(r.Mode) === 'prevention'));
        if (!bad.length) return pass(test, 'Request Body Inspection is enabled on all Azure Front Door WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${bad.length} Azure Front Door WAF policy(ies) are not in Prevention mode with request body inspection enabled.`, { policies: rows.length, nonCompliant: bad.map(r => ({ policy: r.PolicyName, state: r.EnabledState, mode: r.Mode, requestBodyCheck: r.RequestBodyCheck })) });
      },
      'No Azure Front Door WAF policies attached to Front Door were found.'),

    // 26881 — Default Ruleset enabled in Application Gateway WAF.
    '26881': async (test) => runArg(test,
      `resources
| where type =~ 'microsoft.network/applicationgatewaywebapplicationfirewallpolicies'
| where coalesce(array_length(properties.applicationGateways), 0) >= 1
| project PolicyName=name, subscriptionId, EnabledState=tostring(properties.policySettings.state), Mode=tostring(properties.policySettings.mode), ManagedRuleSets=properties.managedRules.managedRuleSets`,
      (rows) => {
        const failing = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.Mode) !== 'prevention' || !toArray(r.ManagedRuleSets).some(rs => ['microsoft_defaultruleset', 'owasp'].includes(lower(rs.ruleSetType))));
        if (!failing.length) return pass(test, 'The Default/OWASP managed ruleset is enabled on all Application Gateway WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${failing.length} Application Gateway WAF policy(ies) lack the default managed ruleset or are not in Prevention mode.`, { policies: rows.length, nonCompliant: failing.map(r => ({ policy: r.PolicyName, state: r.EnabledState, mode: r.Mode })) });
      },
      'No Application Gateway WAF policies attached to Application Gateways were found.'),

    // 26882 — Bot protection ruleset enabled in Application Gateway WAF.
    '26882': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/applicationgatewaywebapplicationfirewallpolicies' | where coalesce(array_length(properties.applicationGateways), 0) >= 1 | project PolicyName=name, EnabledState=tostring(properties.policySettings.state), Mode=tostring(properties.policySettings.mode), ManagedRuleSets=properties.managedRules.managedRuleSets`,
      (rows) => {
        const failing = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.Mode) !== 'prevention' || !toArray(r.ManagedRuleSets).some(rs => lower(rs.ruleSetType) === 'microsoft_botmanagerruleset'));
        if (!failing.length) return pass(test, 'The Bot Manager ruleset is enabled on all Application Gateway WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${failing.length} Application Gateway WAF policy(ies) lack the Bot Manager ruleset or are not in Prevention mode.`, { policies: rows.length, nonCompliant: failing.map(r => r.PolicyName) });
      },
      'No Application Gateway WAF policies attached to Application Gateways were found.'),

    // 26883 — Default rule set assigned in Azure Front Door WAF (partial: override depth not checked).
    '26883': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/frontdoorwebapplicationfirewallpolicies' | where array_length(properties.frontendEndpointLinks) > 0 or array_length(properties.securityPolicyLinks) > 0 | project PolicyName=name, EnabledState=tostring(properties.policySettings.enabledState), WafMode=tostring(properties.policySettings.mode), ManagedRuleSets=properties.managedRules.managedRuleSets`,
      (rows) => {
        const failing = rows.filter(r => !toArray(r.ManagedRuleSets).some(rs => lower(rs.ruleSetType) === 'microsoft_defaultruleset'));
        if (!failing.length) return pass(test, 'The Microsoft Default Rule Set is assigned on all Azure Front Door WAF policies.', { policies: rows.length });
        return fail(test, `${failing.length} Azure Front Door WAF policy(ies) do not have the Microsoft Default Rule Set assigned.`, { policies: rows.length, nonCompliant: failing.map(r => r.PolicyName) });
      },
      'No Azure Front Door WAF policies attached to Front Door were found.'),

    // 26884 — Bot protection rule set enabled in Azure Front Door WAF (partial).
    '26884': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/frontdoorwebapplicationfirewallpolicies' | where sku.name =~ 'Premium_AzureFrontDoor' | project PolicyName=name, ManagedRuleSets=properties.managedRules.managedRuleSets`,
      (rows) => {
        const failing = rows.filter(r => !toArray(r.ManagedRuleSets).some(rs => lower(rs.ruleSetType) === 'microsoft_botmanagerruleset'));
        if (!failing.length) return pass(test, 'The Bot Manager rule set is enabled on all Premium Azure Front Door WAF policies.', { policies: rows.length });
        return fail(test, `${failing.length} Premium Azure Front Door WAF policy(ies) lack the Bot Manager rule set.`, { policies: rows.length, nonCompliant: failing.map(r => r.PolicyName) });
      },
      'No Premium Azure Front Door WAF policies were found (Bot protection requires the Premium SKU).'),

    // 26885 — Metrics enabled for DDoS-protected public IPs.
    '26885': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/publicipaddresses' | where properties.provisioningState =~ 'Succeeded' | project PublicIpName=name, PublicIpId=id, ProtectionMode=tostring(properties.ddosSettings.protectionMode)`,
      async (rows) => {
        const protectedIps = rows.filter(r => ['enabled', 'virtualnetworkinherited'].includes(lower(r.ProtectionMode)));
        if (!protectedIps.length) return skip(test, `${skippedReason('NotApplicable')} No DDoS-protected public IP addresses were found.`);
        return diagCheck(test, protectedIps, 'PublicIpId', 'PublicIpName', 'metrics', 'DDoS-protected public IPs');
      },
      'No public IP addresses were found.'),

    // 26886 — Diagnostic logging enabled for DDoS-protected public IPs.
    '26886': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/publicipaddresses' | where properties.provisioningState =~ 'Succeeded' | project PublicIpName=name, PublicIpId=id, ProtectionMode=tostring(properties.ddosSettings.protectionMode)`,
      async (rows) => {
        const protectedIps = rows.filter(r => ['enabled', 'virtualnetworkinherited'].includes(lower(r.ProtectionMode)));
        if (!protectedIps.length) return skip(test, `${skippedReason('NotApplicable')} No DDoS-protected public IP addresses were found.`);
        return diagCheck(test, protectedIps, 'PublicIpId', 'PublicIpName', 'logs', 'DDoS-protected public IPs');
      },
      'No public IP addresses were found.'),

    // 26887 — Diagnostic logging enabled in Azure Firewall.
    '26887': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/azurefirewalls' | where properties.provisioningState =~ 'Succeeded' | project FirewallName=name, FirewallId=id`,
      async (rows) => diagCheck(test, rows, 'FirewallId', 'FirewallName', 'logs', 'Azure Firewalls'),
      'No Azure Firewalls were found.'),

    // 26888 — Diagnostic logging enabled in Application Gateway WAF.
    '26888': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/applicationgateways' | where properties.provisioningState =~ 'Succeeded' | where properties.sku.tier in~ ('WAF', 'WAF_v2') | project GatewayName=name, GatewayId=id`,
      async (rows) => diagCheck(test, rows, 'GatewayId', 'GatewayName', 'logs', 'WAF-enabled Application Gateways'),
      'No WAF-enabled Application Gateways were found.'),

    // 26889 — Diagnostic logging enabled in Azure Front Door WAF.
    '26889': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.cdn/profiles' | where sku.name in~ ('Standard_AzureFrontDoor', 'Premium_AzureFrontDoor') | project FrontDoorId=tolower(id), FrontDoorName=name`,
      async (rows) => diagCheck(test, rows, 'FrontDoorId', 'FrontDoorName', 'logs', 'Azure Front Door profiles'),
      'No Azure Front Door profiles were found.'),

    // 27015 — HTTP DDoS Protection Ruleset enabled in Application Gateway WAF.
    '27015': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/applicationgatewaywebapplicationfirewallpolicies' | where coalesce(array_length(properties.applicationGateways), 0) >= 1 | project PolicyName=name, EnabledState=tostring(properties.policySettings.state), Mode=tostring(properties.policySettings.mode), ManagedRuleSets=properties.managedRules.managedRuleSets`,
      (rows) => {
        const hasHttpDdos = (r) => toArray(r.ManagedRuleSets).some(rs => {
          if (lower(rs.ruleSetType) !== 'microsoft_httpddosruleset') return false;
          const overrides = toArray(rs.ruleGroupOverrides).flatMap(o => toArray(o.rules)).filter(Boolean);
          const disabled = overrides.filter(x => lower(x.state) === 'disabled');
          return !(overrides.length > 0 && disabled.length === overrides.length);
        });
        const failing = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.Mode) !== 'prevention' || !hasHttpDdos(r));
        if (!failing.length) return pass(test, 'The HTTP DDoS protection ruleset is enabled on all Application Gateway WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${failing.length} Application Gateway WAF policy(ies) lack an enabled HTTP DDoS protection ruleset.`, { policies: rows.length, nonCompliant: failing.map(r => r.PolicyName) });
      },
      'No Application Gateway WAF policies attached to Application Gateways were found.'),

    // 27016 — Rate Limiting enabled in Application Gateway WAF.
    '27016': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/applicationgatewaywebapplicationfirewallpolicies' | where coalesce(array_length(properties.applicationGateways), 0) >= 1 | project PolicyName=name, EnabledState=tostring(properties.policySettings.state), Mode=tostring(properties.policySettings.mode), CustomRules=properties.customRules`,
      (rows) => {
        const failing = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.Mode) !== 'prevention' || !toArray(r.CustomRules).some(cr => lower(cr.ruleType) === 'ratelimitrule' && lower(cr.state) === 'enabled'));
        if (!failing.length) return pass(test, 'Rate limiting is enabled on all Application Gateway WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${failing.length} Application Gateway WAF policy(ies) do not have an enabled rate limit rule.`, { policies: rows.length, nonCompliant: failing.map(r => r.PolicyName) });
      },
      'No Application Gateway WAF policies attached to Application Gateways were found.'),
    // 27017 — JavaScript Challenge enabled in Application Gateway WAF.
    '27017': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/applicationgatewaywebapplicationfirewallpolicies' | where coalesce(array_length(properties.applicationGateways), 0) >= 1 | project PolicyName=name, EnabledState=tostring(properties.policySettings.state), Mode=tostring(properties.policySettings.mode), CustomRules=properties.customRules`,
      (rows) => {
        const failing = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.Mode) !== 'prevention' || !toArray(r.CustomRules).some(cr => lower(cr.action) === 'jschallenge' && lower(cr.state) === 'enabled'));
        if (!failing.length) return pass(test, 'JavaScript Challenge is enabled on all Application Gateway WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${failing.length} Application Gateway WAF policy(ies) do not have an enabled JavaScript Challenge rule.`, { policies: rows.length, nonCompliant: failing.map(r => r.PolicyName) });
      },
      'No Application Gateway WAF policies attached to Application Gateways were found.'),

    // 27018 — Rate Limiting enabled in Azure Front Door WAF.
    '27018': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/frontdoorwebapplicationfirewallpolicies' | where array_length(properties.frontendEndpointLinks) > 0 or array_length(properties.securityPolicyLinks) > 0 | project PolicyName=name, EnabledState=tostring(properties.policySettings.enabledState), WafMode=tostring(properties.policySettings.mode), CustomRules=properties.customRules.rules`,
      (rows) => {
        const failing = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.WafMode) !== 'prevention' || !toArray(r.CustomRules).some(cr => lower(cr.ruleType) === 'ratelimitrule' && lower(cr.enabledState) === 'enabled'));
        if (!failing.length) return pass(test, 'Rate limiting is enabled on all Azure Front Door WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${failing.length} Azure Front Door WAF policy(ies) do not have an enabled rate limit rule.`, { policies: rows.length, nonCompliant: failing.map(r => r.PolicyName) });
      },
      'No Azure Front Door WAF policies attached to Front Door were found.'),

    // 27019 — JavaScript Challenge enabled in Azure Front Door WAF.
    '27019': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/frontdoorwebapplicationfirewallpolicies' | where array_length(properties.frontendEndpointLinks) > 0 or array_length(properties.securityPolicyLinks) > 0 | project PolicyName=name, EnabledState=tostring(properties.policySettings.enabledState), WafMode=tostring(properties.policySettings.mode), CustomRules=properties.customRules.rules`,
      (rows) => {
        const failing = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.WafMode) !== 'prevention' || !toArray(r.CustomRules).some(cr => lower(cr.action) === 'jschallenge' && lower(cr.enabledState) === 'enabled'));
        if (!failing.length) return pass(test, 'JavaScript Challenge is enabled on all Azure Front Door WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${failing.length} Azure Front Door WAF policy(ies) do not have an enabled JavaScript Challenge rule.`, { policies: rows.length, nonCompliant: failing.map(r => r.PolicyName) });
      },
      'No Azure Front Door WAF policies attached to Front Door were found.'),

    // 27020 — CAPTCHA challenge enabled in Azure Front Door WAF.
    '27020': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/frontdoorwebapplicationfirewallpolicies' | where array_length(properties.frontendEndpointLinks) > 0 or array_length(properties.securityPolicyLinks) > 0 | project PolicyName=name, EnabledState=tostring(properties.policySettings.enabledState), WafMode=tostring(properties.policySettings.mode), CustomRules=properties.customRules.rules`,
      (rows) => {
        const failing = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.WafMode) !== 'prevention' || !toArray(r.CustomRules).some(cr => lower(cr.action) === 'captcha' && lower(cr.enabledState) === 'enabled'));
        if (!failing.length) return pass(test, 'CAPTCHA challenge is enabled on all Azure Front Door WAF policies (Prevention mode).', { policies: rows.length });
        return fail(test, `${failing.length} Azure Front Door WAF policy(ies) do not have an enabled CAPTCHA challenge rule.`, { policies: rows.length, nonCompliant: failing.map(r => r.PolicyName) });
      },
      'No Azure Front Door WAF policies attached to Front Door were found.'),
    // 25419 — Network access activity visible to SecOps (GSA traffic logs) (partial).
    '25419': async (test) => {
      let settings;
      try { const r = await Api.arm('providers/microsoft.aadiam/diagnosticSettings?api-version=2017-04-01-preview'); settings = toArray(r?.value); }
      catch (err) { if (err.armConsent || [401, 403].includes(err.status)) return skip(test, skippedReason('NotConnectedAzure')); if ([400, 404].includes(err.status)) return skip(test, `${skippedReason('NotSupported')}`); throw err; }
      const hasNetworkLogs = settings.some(s => toArray(s.properties?.logs).some(l => l.enabled && lower(l.category).includes('networkaccess')));
      const evidence = { diagnosticSettings: settings.length, networkAccessLogsEnabled: hasNetworkLogs };
      if (hasNetworkLogs) return pass(test, 'Network access (Global Secure Access) traffic logs are exported for security operations.', evidence);
      return fail(test, 'Global Secure Access network access logs are not being exported to a monitoring destination.', evidence);
    },

    // 25420 — Network access logs retained in Log Analytics (partial).
    '25420': async (test) => {
      let settings;
      try { const r = await Api.arm('providers/microsoft.aadiam/diagnosticSettings?api-version=2017-04-01-preview'); settings = toArray(r?.value); }
      catch (err) { if (err.armConsent || [401, 403].includes(err.status)) return skip(test, skippedReason('NotConnectedAzure')); if ([400, 404].includes(err.status)) return skip(test, `${skippedReason('NotSupported')}`); throw err; }
      const retained = settings.some(s => s.properties?.workspaceId && toArray(s.properties?.logs).some(l => l.enabled && lower(l.category).includes('networkaccess')));
      const evidence = { diagnosticSettings: settings.length, retainedInLogAnalytics: retained };
      if (retained) return pass(test, 'Network access logs are retained in a Log Analytics workspace for security analysis and compliance.', evidence);
      return fail(test, 'Network access logs are not retained in a Log Analytics workspace.', evidence);
    },

    // 25533 — DDoS Protection enabled for public IPs attached in VNETs (partial: VNet plan inheritance assumed protected).
    '25533': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/publicipaddresses' | where isnotempty(properties.ipConfiguration.id) | project PublicIpName=name, ProtectionMode=tostring(properties.ddosSettings.protectionMode)`,
      (rows) => {
        const unprotected = rows.filter(r => lower(r.ProtectionMode || 'disabled') === 'disabled').map(r => r.PublicIpName);
        if (!unprotected.length) return pass(test, 'All attached public IP addresses have DDoS protection enabled (directly or inherited from the VNet).', { publicIps: rows.length });
        return fail(test, `${unprotected.length} attached public IP address(es) have DDoS protection disabled.`, { publicIps: rows.length, unprotected: unprotected.slice(0, 20) });
      },
      'No attached public IP addresses were found.'),

    // 25535 — Outbound traffic routed through Azure Firewall (partial: default-route heuristic).
    '25535': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/routetables' | mv-expand route = properties.routes | where tostring(route.properties.addressPrefix) == '0.0.0.0/0' | project RouteTable=name, NextHopType=tostring(route.properties.nextHopType), NextHop=tostring(route.properties.nextHopIpAddress)`,
      (rows) => {
        const viaAppliance = rows.filter(r => lower(r.NextHopType) === 'virtualappliance');
        if (viaAppliance.length > 0) return pass(test, 'Outbound (0.0.0.0/0) traffic is routed through a virtual appliance (Azure Firewall) via user-defined routes.', { defaultRoutes: rows.length, viaVirtualAppliance: viaAppliance.length });
        return fail(test, 'No user-defined default route forces outbound traffic through Azure Firewall (virtual appliance).', { defaultRoutes: rows.length });
      },
      'No route tables with a default (0.0.0.0/0) route were found.'),

    // 25537 — Threat intelligence in Deny mode on Azure Firewall.
    '25537': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/firewallpolicies' | where tostring(properties.sku.tier) in ('Standard', 'Premium') | project PolicyName=name, ThreatIntelMode=tostring(properties.threatIntelMode)`,
      (rows) => {
        const notDeny = rows.filter(r => lower(r.ThreatIntelMode) !== 'deny').map(r => ({ policy: r.PolicyName, mode: r.ThreatIntelMode || 'Unknown' }));
        if (!notDeny.length) return pass(test, 'Threat intelligence is enabled in Deny mode on all Azure Firewall policies.', { policies: rows.length });
        return fail(test, `${notDeny.length} Azure Firewall policy(ies) do not have threat intelligence in Deny mode.`, { policies: rows.length, nonCompliant: notDeny });
      },
      'No Standard/Premium Azure Firewall policies were found.'),

    // 25539 — IDPS in Deny mode on Azure Firewall (partial: reads intrusionDetection.mode via ARG).
    '25539': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/firewallpolicies' | where tostring(properties.sku.tier) =~ 'Premium' | project PolicyName=name, IdpsMode=tostring(properties.intrusionDetection.mode)`,
      (rows) => {
        const notDeny = rows.filter(r => lower(r.IdpsMode) !== 'deny').map(r => ({ policy: r.PolicyName, mode: r.IdpsMode || 'Off' }));
        if (!notDeny.length) return pass(test, 'IDPS is enabled in Deny mode on all Premium Azure Firewall policies.', { policies: rows.length });
        return fail(test, `${notDeny.length} Premium Azure Firewall policy(ies) do not have IDPS in Deny mode.`, { policies: rows.length, nonCompliant: notDeny });
      },
      'No Premium Azure Firewall policies were found.'),

    // 25541 — Application Gateway WAF enabled in Prevention mode.
    '25541': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/ApplicationGatewayWebApplicationFirewallPolicies' | project PolicyName=name, EnabledState=tostring(properties.policySettings.state), Mode=tostring(properties.policySettings.mode)`,
      (rows) => {
        const bad = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.Mode) !== 'prevention').map(r => r.PolicyName);
        if (!bad.length) return pass(test, 'All Application Gateway WAF policies are enabled in Prevention mode.', { policies: rows.length });
        return fail(test, `${bad.length} Application Gateway WAF policy(ies) are not enabled in Prevention mode.`, { policies: rows.length, nonCompliant: bad });
      },
      'No Application Gateway WAF policies were found.'),

    // 25543 — Azure Front Door WAF enabled in Prevention mode.
    '25543': async (test) => runArg(test,
      `resources | where type =~ 'microsoft.network/frontdoorwebapplicationfirewallpolicies' | project PolicyName=name, EnabledState=tostring(properties.policySettings.enabledState), Mode=tostring(properties.policySettings.mode)`,
      (rows) => {
        const bad = rows.filter(r => lower(r.EnabledState) !== 'enabled' || lower(r.Mode) !== 'prevention').map(r => r.PolicyName);
        if (!bad.length) return pass(test, 'All Azure Front Door WAF policies are enabled in Prevention mode.', { policies: rows.length });
        return fail(test, `${bad.length} Azure Front Door WAF policy(ies) are not enabled in Prevention mode.`, { policies: rows.length, nonCompliant: bad });
      },
      'No Azure Front Door WAF policies were found.'),

    // 25550 — Outbound TLS inspection enabled on Azure Firewall (partial: cert + terminateTLS rule).
    '25550': async (test) => {
      const [policies, tlsGroups] = await Promise.all([
        Api.armResourceGraph(`resources | where type =~ 'microsoft.network/firewallpolicies' | project PolicyId=tolower(id), PolicyName=name, SkuTier=tostring(properties.sku.tier), CertName=tostring(properties.transportSecurity.certificateAuthority.name)`).catch(e => { throw e; }),
        Api.armResourceGraph(`resources | where type =~ 'microsoft.network/firewallpolicies/rulecollectiongroups' | mv-expand ruleCollection = properties.ruleCollections | mv-expand rule = ruleCollection.rules | where tostring(rule.ruleType) =~ 'ApplicationRule' and tobool(rule.terminateTLS) == true | extend lowerId = tolower(id) | extend policyId = substring(lowerId, 0, indexof(lowerId, '/rulecollectiongroups/')) | distinct policyId`).catch(() => []),
      ]).catch((err) => { if (err?.armConsent || [401, 403].includes(err?.status)) return 'noaccess'; if ([400, 404].includes(err?.status)) return 'unsupported'; throw err; });
      if (policies === 'noaccess') return skip(test, skippedReason('NotConnectedAzure'));
      if (policies === 'unsupported') return skip(test, `${skippedReason('NotSupported')}`);
      const [pol, tls] = [policies, tlsGroups];
      if (!toArray(pol).length) return skip(test, `${skippedReason('NotApplicable')} No Azure Firewall policies were found.`);
      const tlsPolicyIds = new Set(toArray(tls).map(t => lower(t.policyId)));
      const compliant = toArray(pol).filter(p => p.CertName && tlsPolicyIds.has(lower(p.PolicyId)));
      if (compliant.length > 0) return pass(test, 'Outbound TLS inspection is enabled on Azure Firewall (certificate configured and TLS-terminating application rules present).', { firewallPolicies: toArray(pol).length, withTlsInspection: compliant.length });
      return fail(test, 'No Azure Firewall policy has outbound TLS inspection fully configured (CA certificate + TLS-terminating application rule).', { firewallPolicies: toArray(pol).length });
    },
    // 27003 — TLS inspection failure rate below 1% (requires Log Analytics KQL, not browser-reachable).
    '27003': async (test) => {
      let tls = null;
      try { tls = await Api.graphAll('networkAccess/tlsInspectionPolicies', { beta: true }); }
      catch (err) { if (![400, 403, 404].includes(err.status)) throw err; }
      if (tls === null || !tls.length) return skip(test, `${skippedReason('NotApplicable')} TLS inspection is not configured in this tenant.`);
      return skip(test, `${skippedReason('NotSupported')} Computing the TLS inspection failure rate requires querying the NetworkAccessTraffic table in Log Analytics, which is not reachable from a browser-only tool.`);
    },

    // 41207 — Active analytics rules configured in Microsoft Sentinel.
    '41207': async (test) => {
      const sw = await sentinelWorkspaces();
      if (sw === 'noaccess') return skip(test, skippedReason('NotConnectedAzure'));
      if (sw === 'unsupported') return skip(test, `${skippedReason('NotSupported')}`);
      if (!sw.workspaces.length) return skip(test, `${skippedReason('NotApplicable')} No Log Analytics workspaces were found.`);
      if (!sw.onboarded.length) return skip(test, `${skippedReason('NotApplicable')} Microsoft Sentinel is not onboarded on any workspace.`);
      let enabledRules = 0;
      for (const w of sw.onboarded) {
        try {
          const r = await Api.arm(`${w.id}/providers/Microsoft.SecurityInsights/alertRules?api-version=2023-11-01`);
          enabledRules += toArray(r?.value).filter(rule => rule.properties?.enabled === true && ['scheduled', 'nrt'].includes(lower(rule.kind))).length;
        } catch { /* forbidden */ }
      }
      const evidence = { sentinelWorkspaces: sw.onboarded.length, enabledScheduledRules: enabledRules };
      if (enabledRules > 0) return pass(test, 'Active (enabled) scheduled/NRT analytics rules are configured in Microsoft Sentinel.', evidence);
      return fail(test, 'No enabled scheduled or NRT analytics rules are configured in Microsoft Sentinel.', evidence);
    },

    // 50001 — Microsoft Defender for Cloud recommendations (partial: high-severity unhealthy assessments).
    '50001': async (test) => runArg(test,
      `securityresources | where type == 'microsoft.security/assessments' | project name, status=tostring(properties.status.code), severity=tostring(properties.metadata.severity)`,
      (rows) => {
        const unhealthyHigh = rows.filter(r => lower(r.status) === 'unhealthy' && lower(r.severity) === 'high');
        const evidence = { assessments: rows.length, unhealthyHigh: unhealthyHigh.length };
        if (!unhealthyHigh.length) return pass(test, 'No high-severity unhealthy Microsoft Defender for Cloud recommendations are present.', evidence);
        return fail(test, `${unhealthyHigh.length} high-severity Microsoft Defender for Cloud recommendation(s) are unhealthy.`, evidence);
      },
      'No Microsoft Defender for Cloud assessments found — ensure Defender for Cloud is enabled.'),

    // 61002 — Microsoft Sentinel onboarded on at least one Log Analytics workspace.
    '61002': async (test) => {
      const sw = await sentinelWorkspaces();
      if (sw === 'noaccess') return skip(test, skippedReason('NotConnectedAzure'));
      if (sw === 'unsupported') return skip(test, `${skippedReason('NotSupported')}`);
      if (!sw.workspaces.length) return skip(test, `${skippedReason('NotApplicable')} No Log Analytics workspaces were found.`);
      const evidence = { workspaces: sw.workspaces.length, sentinelOnboarded: sw.onboarded.length };
      if (sw.onboarded.length >= 1) return pass(test, 'Microsoft Sentinel is onboarded on at least one Log Analytics workspace.', evidence);
      return fail(test, 'Microsoft Sentinel is not onboarded on any Log Analytics workspace.', evidence);
    },

    // 61004 — Defender for Cloud CSPM plan enabled on all subscriptions.
    '61004': async (test) => {
      let subs;
      try { subs = await Api.armResourceGraph(`resourcecontainers | where type =~ 'microsoft.resources/subscriptions' | where properties.state =~ 'Enabled' | project subscriptionId, displayName=name`); }
      catch (err) { if (err.armConsent || [401, 403].includes(err.status)) return skip(test, skippedReason('NotConnectedAzure')); if ([400, 404].includes(err.status)) return skip(test, `${skippedReason('NotSupported')}`); throw err; }
      if (!toArray(subs).length) return skip(test, `${skippedReason('NotApplicable')} No enabled Azure subscriptions were found.`);
      const missing = [];
      for (const s of subs) {
        try {
          const p = await Api.arm(`subscriptions/${s.subscriptionId}/providers/Microsoft.Security/pricings/CloudPosture?api-version=2024-01-01`);
          if (lower(p?.properties?.pricingTier) !== 'standard') missing.push(s.displayName);
        } catch { missing.push(s.displayName); }
      }
      const evidence = { subscriptions: subs.length, withoutCspm: missing.slice(0, 20) };
      if (!missing.length) return pass(test, 'The Defender CSPM plan is enabled on all Azure subscriptions.', evidence);
      return fail(test, `${missing.length} Azure subscription(s) do not have the Defender CSPM plan enabled.`, evidence);
    },

    // 61016 — Entra ID Protection risk events flowing to Sentinel (partial).
    '61016': async (test) => connectorTest(test, 'azureactivedirectory', 'Microsoft Entra ID Protection'),

    // 61018 — Purview Information Protection data connector on Sentinel (partial).
    '61018': async (test) => connectorTest(test, 'informationprotection', 'Microsoft Purview Information Protection'),

    // 61021 — Microsoft 365 Copilot data connector on Sentinel (partial).
    '61021': async (test) => connectorTest(test, 'copilot', 'Microsoft 365 Copilot'),

    // 61022 — Defender for AI Services enabled on subscriptions hosting AI accounts (partial).
    '61022': async (test) => {
      let aiSubs;
      try { aiSubs = await Api.armResourceGraph(`resources | where type =~ 'microsoft.cognitiveservices/accounts' | distinct subscriptionId`); }
      catch (err) { if (err.armConsent || [401, 403].includes(err.status)) return skip(test, skippedReason('NotConnectedAzure')); if ([400, 404].includes(err.status)) return skip(test, `${skippedReason('NotSupported')}`); throw err; }
      if (!toArray(aiSubs).length) return skip(test, `${skippedReason('NotApplicable')} No Azure OpenAI / Azure AI Services accounts were found.`);
      const missing = [];
      for (const s of aiSubs) {
        try {
          const p = await Api.arm(`subscriptions/${s.subscriptionId}/providers/Microsoft.Security/pricings/AI?api-version=2024-01-01`);
          if (lower(p?.properties?.pricingTier) !== 'standard') missing.push(s.subscriptionId);
        } catch { missing.push(s.subscriptionId); }
      }
      const evidence = { subscriptionsWithAi: aiSubs.length, withoutDefenderForAi: missing.length };
      if (!missing.length) return pass(test, 'Microsoft Defender for AI Services is enabled on every subscription hosting Azure OpenAI / AI Services accounts.', evidence);
      return fail(test, `${missing.length} subscription(s) hosting AI accounts do not have Microsoft Defender for AI Services enabled.`, evidence);
    },

    // 61023 — Agent 365 data connector on Sentinel (partial).
    '61023': async (test) => connectorTest(test, 'agent', 'Agent 365'),

    // 61024 — Defender XDR (unified) data connector on Sentinel (partial).
    '61024': async (test) => connectorTest(test, 'threatprotection', 'Microsoft Defender XDR (unified)'),
  };

  ZtaImpl.register(impl);
  window.ZtaAzureLib = { runArg, anyDiagnostic, diagCheck, sentinelWorkspaces };
})();
