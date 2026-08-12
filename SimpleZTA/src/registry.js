(() => {
  let cached = null;
  const ARM_USER_IMPERSONATION_SCOPE = 'https://management.azure.com/user_impersonation';

  const DOC_ROOTS = {
    identity: 'https://learn.microsoft.com/en-us/entra/fundamentals/configure-security?toc=/security/zero-trust/assessment/toc.json&bc=/security/zero-trust/assessment/toc.json',
    devices: 'https://learn.microsoft.com/en-us/intune/device-security/ref-zero-trust-security?toc=%2Fsecurity%2Fzero-trust%2Fassessment%2Ftoc.json&bc=%2Fsecurity%2Fzero-trust%2Fassessment%2Ftoc.json',
    network: 'https://learn.microsoft.com/en-us/azure/networking/security/zero-trust-network-security?toc=/security/zero-trust/assessment/toc.json&bc=/security/zero-trust/assessment/toc.json',
    data: 'https://learn.microsoft.com/en-us/purview/configure-security?toc=/security/zero-trust/assessment/toc.json&bc=/security/zero-trust/assessment/toc.json',
    infrastructure: 'https://learn.microsoft.com/en-us/azure/networking/security/zero-trust-network-security?toc=/security/zero-trust/assessment/toc.json&bc=/security/zero-trust/assessment/toc.json',
    secops: 'https://learn.microsoft.com/en-us/entra/fundamentals/configure-security?toc=/security/zero-trust/assessment/toc.json&bc=/security/zero-trust/assessment/toc.json',
    ai: 'https://learn.microsoft.com/en-us/azure/networking/security/zero-trust-network-security?toc=/security/zero-trust/assessment/toc.json&bc=/security/zero-trust/assessment/toc.json',
    fallback: 'https://learn.microsoft.com/en-us/entra/fundamentals/configure-security?toc=/security/zero-trust/assessment/toc.json&bc=/security/zero-trust/assessment/toc.json',
  };

  const PARTIAL_IMPLEMENTATION_NOTES = {
    '21770': {
      implemented: 'Detects inactive service principals and checks for high-privilege Graph permission exposure using delegated Graph APIs.',
      notImplemented: 'Does not evaluate every consent path (for example, legacy app role grants outside sampled payload fields) and may miss data hidden by tenant permissions.',
    },
    '21777': {
      implemented: 'Verifies multitenant app instance lock posture through available Graph application properties in browser-safe delegated mode.',
      notImplemented: 'Cannot fully validate all edge-case lock semantics for every application type when Graph omits app-instance internals.',
    },
  };

  function toArray(value) {
    if (Array.isArray(value)) return value;
    if (value == null || value === '') return [];
    return [value];
  }

  // Controls whose browser port approximates a complex PowerShell helper (e.g.
  // Find-ZtProfilesLinkedToPolicy), uses a heuristic in place of a DuckDB export join, or
  // relies on a preview Graph/ARM surface. Marked "partial" so the UI is honest about fidelity.
  const PARTIAL_APPROXIMATION_IDS = new Set([
    '24518', '25395', '25396', '25398', '25400', '25408', '25409', '25410', '25411', '25413', '25415',
    '27000', '41003', '51013', '51018',
    '61009', '61011', '61012', '61013', '61014',
    '25419', '25420', '25533', '25535', '25539', '25550',
    '26883', '26884', '26885', '26886', '26887', '26888', '26889',
    '50001', '61016', '61018', '61021', '61022', '61023', '61024',
  ]);
  const GENERIC_PARTIAL_NOTE = {
    implemented: 'Evaluates the control against live delegated Graph / Azure Resource Manager data and reports a pass/fail with evidence.',
    notImplemented: 'Approximates a complex PowerShell helper, a DuckDB export join, or a preview API surface; edge cases may differ from the PowerShell module. Confirm findings in the relevant admin portal.',
  };

  function normalizeMeta(id, m) {
    // Catalog multi-value fields (Pillar/Service/TenantType/licenses) may arrive as
    // arrays or, when the module declared a single value, as scalars (ConvertTo-Json
    // flattens 1-element arrays). Coerce everything so consumers can rely on arrays.
    const pillars = toArray(m.Pillar);
    return {
      id: `${id}`,
      title: m.Title || `Test ${id}`,
      category: m.Category || '',
      pillar: pillars[0] || 'Unknown',
      pillars,
      risk: m.RiskLevel || 'Unknown',
      implementationCost: m.ImplementationCost || '',
      tenantType: toArray(m.TenantType),
      service: toArray(m.Service),
      minimumLicense: toArray(m.MinimumLicense),
      compatibleLicense: toArray(m.CompatibleLicense),
      sfiPillar: m.SfiPillar || '',
      userImpact: m.UserImpact || '',
    };
  }

  function docsRootFor(meta) {
    const pillar = `${meta.pillar || ''}`.toLowerCase();
    if (pillar.includes('identity')) return DOC_ROOTS.identity;
    if (pillar.includes('device')) return DOC_ROOTS.devices;
    if (pillar.includes('network')) return DOC_ROOTS.network;
    if (pillar.includes('data')) return DOC_ROOTS.data;
    if (pillar.includes('infra')) return DOC_ROOTS.infrastructure;
    if (pillar.includes('secops')) return DOC_ROOTS.secops;
    if (pillar.includes('ai')) return DOC_ROOTS.ai;

    const category = `${meta.category || ''}`.toLowerCase();
    if (category.includes('device') || category.includes('intune')) return DOC_ROOTS.devices;
    if (category.includes('network') || category.includes('tls') || category.includes('firewall')) return DOC_ROOTS.network;
    if (category.includes('data') || category.includes('purview')) return DOC_ROOTS.data;
    return DOC_ROOTS.fallback;
  }

  function docsSearchUrl(meta) {
    const terms = `${meta.title} ${meta.pillar || ''} zero trust`;
    return `https://learn.microsoft.com/en-us/search/?terms=${encodeURIComponent(terms)}`;
  }

  function neededScopesForTest(test) {
    const scopes = [...Auth.coreScopes];
    // Controls that call Azure Resource Manager must pre-consent the ARM scope
    // before the run starts to avoid redirect prompts mid-scan. This covers the
    // legacy ARM control (21788) and every control whose module Service is Azure.
    const usesArm = `${test.id}` === '21788' || (Array.isArray(test.service) && test.service.some(s => `${s}`.toLowerCase() === 'azure'));
    if (usesArm) scopes.push(ARM_USER_IMPERSONATION_SCOPE);
    return scopes;
  }

  async function getAll() {
    if (cached) return cached;
    const res = await fetch('./tests/TestMeta.json', { cache: 'no-store' });
    if (!res.ok) throw new Error(`Could not load test metadata (${res.status})`);
    const data = await res.json();

    cached = Object.entries(data)
      .map(([id, m]) => {
        const meta = normalizeMeta(id, m);
        const impl = ZtaImpl.get(meta.id);
        const notes = PARTIAL_IMPLEMENTATION_NOTES[meta.id] || (PARTIAL_APPROXIMATION_IDS.has(meta.id) ? GENERIC_PARTIAL_NOTE : null);
        return {
          ...meta,
          implemented: !!impl,
          run: impl,
          requiredScopes: neededScopesForTest(meta),
          implementationLevel: impl ? (notes ? 'partial' : 'full') : 'none',
          implementationNotes: notes,
          documentationUrl: docsRootFor(meta),
          documentationSearchUrl: docsSearchUrl(meta),
        };
      })
      .sort((a, b) => Number(a.id) - Number(b.id));

    return cached;
  }

  async function getById(id) {
    const all = await getAll();
    return all.find(t => t.id === `${id}`) || null;
  }

  window.Registry = { getAll, getById };
})();
