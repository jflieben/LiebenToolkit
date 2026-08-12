/*
 * Phase 4 — Buckets C/D/E: Microsoft Purview / Exchange Online / SharePoint Online controls.
 *
 * These module tests read configuration through the Security & Compliance, Exchange Online,
 * and SharePoint Online PowerShell sessions (Get-Label, Get-DlpCompliancePolicy,
 * Get-IRMConfiguration, Get-SPOTenant, ...). Those admin surfaces have NO browser-accessible
 * delegated REST/Graph equivalent, so — exactly as the module does when a service is not
 * connected — the browser engine returns a Skipped result carrying the specific reason.
 *
 * The one exception reachable from delegated Graph is sensitivity-label existence (35003).
 */
(() => {
  const { pass, fail, skip, skippedReason, toArray } = window.ZtaLib;

  // Sensitivity labels are readable via delegated Graph (InformationProtectionPolicy.Read).
  async function getSensitivityLabels() {
    for (const path of ['security/informationProtection/sensitivityLabels', 'informationProtection/policy/labels']) {
      try {
        const r = await Api.graph(path, { beta: true });
        if (r) return toArray(r.value);
      } catch (err) {
        if (![400, 403, 404].includes(err.status)) throw err;
      }
    }
    return null;
  }

  const impl = {
    // 35003 — Sensitivity labels are configured (delegated Graph).
    '35003': async (test) => {
      const labels = await getSensitivityLabels();
      if (labels === null) return skip(test, `${skippedReason('NotConnectedSecurityCompliance')} (Sensitivity labels could not be read via Microsoft Graph — InformationProtectionPolicy.Read may not be consented.)`);
      if (!labels.length) return fail(test, 'No sensitivity labels are configured in the tenant.', { labels: 0 });
      return pass(test, `${labels.length} sensitivity label(s) are configured in the tenant.`, { labels: labels.length });
    },
  };

  // Faithful skips for the controls that require an admin PowerShell session with no
  // browser-reachable delegated API. Grouped by the service they would connect to.
  const SECURITY_COMPLIANCE = ['35004', '35009', '35010', '35011', '35012', '35013', '35014', '35015', '35016', '35017', '35018', '35019', '35020', '35021', '35022', '35023', '35028', '35030', '35032', '35033', '35034', '35035', '35036', '35038', '35039', '35040', '35041'];
  const EXCHANGE_ONLINE = ['35024', '35025', '35026', '35027', '35029', '35037'];
  const SHAREPOINT_ONLINE = ['35005', '35006', '35007', '35008'];

  const PORTAL_HINT = {
    compliance: ' Verify in the Microsoft Purview portal (purview.microsoft.com).',
    exchange: ' Verify with Exchange Online PowerShell / the Exchange admin center.',
    sharepoint: ' Verify in the SharePoint admin center.',
  };

  for (const id of SECURITY_COMPLIANCE) impl[id] = async (test) => skip(test, `${skippedReason('NotConnectedSecurityCompliance')}${PORTAL_HINT.compliance}`);
  for (const id of EXCHANGE_ONLINE) impl[id] = async (test) => skip(test, `${skippedReason('NotConnectedExchange')}${PORTAL_HINT.exchange}`);
  for (const id of SHAREPOINT_ONLINE) impl[id] = async (test) => skip(test, `${skippedReason('NotConnectedSharePoint')}${PORTAL_HINT.sharepoint}`);

  ZtaImpl.register(impl);
})();
