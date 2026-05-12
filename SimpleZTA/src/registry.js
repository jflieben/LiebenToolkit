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
    '21775': {
      implemented: 'Checks the tenant app management policy endpoint and reports whether baseline policy configuration is present.',
      notImplemented: 'Does not fully evaluate every nested policy rule and exception path used in large enterprise app governance models.',
    },
    '21777': {
      implemented: 'Verifies multitenant app instance lock posture through available Graph application properties in browser-safe delegated mode.',
      notImplemented: 'Cannot fully validate all edge-case lock semantics for every application type when Graph omits app-instance internals.',
    },
  };

  function normalizeMeta(id, m) {
    return {
      id: `${id}`,
      title: m.Title || `Test ${id}`,
      category: m.Category || '',
      pillar: m.Pillar || 'Unknown',
      risk: m.RiskLevel || 'Unknown',
      implementationCost: m.ImplementationCost || '',
      tenantType: Array.isArray(m.TenantType) ? m.TenantType : [],
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
    // Controls that call Azure Resource Manager must pre-consent ARM scope
    // before the run starts to avoid redirect prompts mid-scan.
    if (`${test.id}` === '21788') scopes.push(ARM_USER_IMPERSONATION_SCOPE);
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
        return {
          ...meta,
          implemented: !!impl,
          run: impl,
          requiredScopes: neededScopesForTest(meta),
          implementationLevel: impl ? (PARTIAL_IMPLEMENTATION_NOTES[meta.id] ? 'partial' : 'full') : 'none',
          implementationNotes: PARTIAL_IMPLEMENTATION_NOTES[meta.id] || null,
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
