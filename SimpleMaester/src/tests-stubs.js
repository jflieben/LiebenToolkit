// Stub catalog: tests that cannot be implemented in a browser PKCE SPA.
// These tests are shown in the catalog with "Not implemented in SimpleMaester" status
// so users can see exactly what additional coverage the PowerShell Maester module provides.
//
// Note: Many tests originally listed here have been promoted to full implementations:
//  • Intune tests → tests-intune.js (DeviceManagement Graph namespace)
//  • Governance/Azure/Licensing/Purview/PIM/Access Reviews → tests-governance.js
//  • SPF/DMARC DNS checks → tests-native.js (DNS-over-HTTPS, no auth needed)
//  • CISA SPF/DKIM/DMARC/SharePoint → tests-cisa.js
//
// Remaining stubs:
//  • Teams: requires Teams Admin REST API (not accessible via delegated browser auth)
//  • Exchange admin config: requires Exchange Admin REST API (CORS-blocked from browser)
//  • XSPM: requires Defender CSPM APIs
//  • AI/Copilot: requires Copilot admin APIs not yet in Graph stable
//  • DLP/Purview complex: requires compliance API
//  • Intune diagnostics: requires dedicated diagnostic scope
//  • Drift: requires stored baseline
//
// Reference: https://maester.dev/docs/tests/
(() => {
  function stub(id, title, category, severity = 'Info', reason = null) {
    const defaultReason = reason || 'Not implemented in SimpleMaester (requires PowerShell or admin API access). Run using the Maester PowerShell module.';
    return {
      id, title, severity, tag: category, category, runCategory: category,
      docUrl: `https://maester.dev/docs/tests/${id}`,
      description: title,
      implemented: false,
      run: () => Promise.resolve({ id, status: 'Skipped', reason: defaultReason, durationMs: 0 }),
    };
  }

  const EXO_REASON = 'Requires Exchange Admin REST API which is CORS-blocked from browser. Run using the Maester PowerShell module.';
  const TEAMS_REASON = 'Requires Teams Admin REST API which is not accessible from browser PKCE auth. Run using the Maester PowerShell module.';
  const DEFENDER_REASON = 'Requires Microsoft Security Exposure Management (XSPM) API (api.security.microsoft.com) which needs WindowsDefenderATP permissions not granted in this app registration. Run using the Maester PowerShell module.';
  const AI_REASON = 'Requires Microsoft 365 Copilot admin APIs not yet fully exposed in Graph. Run using the Maester PowerShell module.';

  const stubs = [
    // ── Intune diagnostics — requires additional Intune scope not worth consent for one test
    stub('MT.1103', 'Intune diagnostics log collection is enabled', 'Intune', 'Low', 'Requires DeviceManagementManagedDevices.Read.All scope and specific diagnostics endpoint. Run using Maester PowerShell module.'),

    // ── Teams admin REST (MT.1037, MT.1042, MT.1045–MT.1048) ────────────────────
    stub('MT.1037', 'External access should be blocked or restricted in Teams', 'Teams', 'High', TEAMS_REASON),
    stub('MT.1042', 'Guest access is configured appropriately in Teams', 'Teams', 'High', TEAMS_REASON),
    stub('MT.1045', 'Anonymous meeting join is disabled in Teams', 'Teams', 'High', TEAMS_REASON),
    stub('MT.1046', 'Teams meeting lobby is configured correctly', 'Teams', 'High', TEAMS_REASON),
    stub('MT.1047', 'Teams recording policy is configured', 'Teams', 'Medium', TEAMS_REASON),
    stub('MT.1048', 'Teams data loss prevention policy is active', 'Teams', 'High', TEAMS_REASON),

    // ── Exchange admin REST config checks ────────────────────────────────────────
    stub('MT.1039', 'Modern Authentication is enabled for Exchange Online', 'Exchange', 'High', EXO_REASON),
    stub('MT.1041', 'Basic Authentication is disabled for Exchange Online clients', 'Exchange', 'High', EXO_REASON),
    stub('MT.1043', 'Malware auto-purge is enabled in Exchange', 'Exchange', 'High', EXO_REASON),
    stub('MT.1044', 'ATP Safe Attachments policy covers Exchange', 'Exchange', 'High', EXO_REASON),
    stub('MT.1062', 'SMTP AUTH is disabled for Exchange Online', 'Exchange', 'High', EXO_REASON),
    stub('MT.1083', 'Audit log is enabled for Exchange Online mailboxes', 'Exchange', 'Medium', EXO_REASON),

    // ── XSPM / Exposure Management (MT.1086–MT.1089) ────────────────────────────
    stub('MT.1086', 'Exposure management critical asset protection score is acceptable', 'XSPM', 'High', DEFENDER_REASON),
    stub('MT.1087', 'Exposure management attack path count is acceptable', 'XSPM', 'High', DEFENDER_REASON),
    stub('MT.1088', 'Exposure management initiative score is acceptable', 'XSPM', 'Medium', DEFENDER_REASON),
    stub('MT.1089', 'No critical exposures found in Defender XSPM', 'XSPM', 'High', DEFENDER_REASON),

    // ── AI Agent / Copilot (MT.1113–MT.1122) ────────────────────────────────────
    stub('MT.1113', 'Copilot is not enabled for users without appropriate licenses', 'AI', 'Medium', AI_REASON),
    stub('MT.1114', 'Copilot plugin access is restricted to approved plugins', 'AI', 'Medium', AI_REASON),
    stub('MT.1115', 'Copilot interaction data is protected appropriately', 'AI', 'Medium', AI_REASON),
    stub('MT.1116', 'Copilot web search grounding is configured per policy', 'AI', 'Medium', AI_REASON),
    stub('MT.1117', 'Copilot audit logging is enabled', 'AI', 'High', AI_REASON),
    stub('MT.1118', 'Copilot oversharing baseline is configured', 'AI', 'High', AI_REASON),
    stub('MT.1119', 'AI Hub data governance policies are configured', 'AI', 'Medium', AI_REASON),
    stub('MT.1120', 'Copilot Studio agents use secure authentication', 'AI', 'High', AI_REASON),
    stub('MT.1121', 'Copilot Studio agents do not allow anonymous access', 'AI', 'High', AI_REASON),
    stub('MT.1122', 'Microsoft 365 Copilot data residency is configured', 'AI', 'Medium', AI_REASON),

    // ── Purview DLP ───────────────────────────────────────────────────────────────
    stub('MT.PUR.1002', 'DLP policies cover sensitive information types', 'Purview', 'High', 'Requires compliance API access (SecurityComplianceCenter) not available via browser PKCE flow.'),

    // ── Drift / Custom ────────────────────────────────────────────────────────────
    stub('MT.DRIFT.1001', 'Configuration drift since last baseline', 'Drift', 'Medium', 'Drift detection requires a stored baseline and comparison logic. Use the Maester PowerShell module baseline feature.'),
  ];

  function buildCatalog() { return stubs; }
  window.TestsStubs = { buildCatalog };
})();

