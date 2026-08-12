(() => {
  // Generates the official Microsoft Zero Trust Assessment HTML report from a SimpleZTA scan.
  // Mirrors Get-HtmlReport.ps1: the assessment-results JSON is spliced into the bundled
  // compiled report template between the `reportData={` and `EndOfJson:"EndOfJson"}` markers.

  const TEMPLATE_URL = './powershell/assets/ReportTemplate.html';
  const MANIFEST_URL = './powershell/ZeroTrustAssessment.psd1';
  const START_MARKER = 'reportData={';
  const END_MARKER = 'EndOfJson:"EndOfJson"}';
  const CORE_PILLARS = ['Identity', 'Devices', 'Network', 'Data'];
  const PREVIEW_PILLARS = ['Infrastructure', 'SecOps', 'AI'];

  let cachedVersion = null;

  async function getModuleVersion() {
    if (cachedVersion) return cachedVersion;
    try {
      const res = await fetch(MANIFEST_URL, { cache: 'force-cache' });
      if (res.ok) {
        const text = await res.text();
        const m = /ModuleVersion\s*=\s*'([^']+)'/.exec(text);
        if (m) cachedVersion = m[1];
      }
    } catch {
      // fall through to default
    }
    if (!cachedVersion) cachedVersion = '0.0.0';
    return cachedVersion;
  }

  // Maps a web-app result status to the module's TestStatus vocabulary.
  // Controls without a browser implementation surface as Planned (module: UnderConstruction).
  function moduleStatus(row) {
    if (row.implementationLevel === 'none') return 'Planned';
    return row.status || 'Skipped';
  }

  function summarize(results) {
    const summary = {};
    const present = new Set(results.map(r => r.pillar).filter(Boolean));
    const pillars = [...CORE_PILLARS, ...PREVIEW_PILLARS.filter(p => present.has(p))];
    for (const pillar of pillars) {
      const inPillar = results.filter(r => r.pillar === pillar);
      const key = pillar === 'AI' ? 'AI' : pillar;
      summary[`${key}Passed`] = inPillar.filter(r => moduleStatus(r) === 'Passed').length;
      summary[`${key}Total`] = inPillar.filter(r => !['Skipped', 'Planned'].includes(moduleStatus(r))).length;
    }
    return summary;
  }

  function minimumLicense(meta) {
    if (!meta) return null;
    if (meta.compatibleLicense && meta.compatibleLicense.length) {
      return meta.compatibleLicense.map(l => `${l}`.split('&').join(' AND '));
    }
    if (meta.minimumLicense && meta.minimumLicense.length) return meta.minimumLicense;
    return null;
  }

  function buildTestEntry(row, meta, md) {
    const status = moduleStatus(row);
    const skipped = status === 'Planned' ? 'UnderConstruction'
      : status === 'Skipped' ? 'NotSupported'
        : '';
    const pillars = meta?.pillars && meta.pillars.length > 1 ? meta.pillars : (row.pillar || meta?.pillar || null);
    return {
      TestId: row.id,
      TestTitle: row.title,
      TestStatus: status,
      TestCategory: row.category || meta?.category || null,
      TestTags: [],
      TestAppliesTo: null,
      TestImpact: meta?.userImpact || null,
      TestRisk: row.risk || meta?.risk || null,
      TestImplementationCost: meta?.implementationCost || null,
      TestSfiPillar: meta?.sfiPillar || null,
      TestPillar: pillars,
      TestMinimumLicense: minimumLicense(meta),
      TestDescription: md?.description || '',
      TestResult: Md.buildResultMarkdown(row, md?.resultTemplate),
      TestSkipped: skipped,
      SkippedReason: skipped ? (row.message || 'Not evaluable in browser mode.') : null,
    };
  }

  async function buildAssessmentJson(scan) {
    const results = scan.results || [];
    const metaById = new Map((await Registry.getAll()).map(t => [t.id, t]));
    const mdById = await Md.loadTestMdMany(results.map(r => r.id));
    const version = await getModuleVersion();
    const account = (typeof Auth !== 'undefined' && Auth.getAccount && Auth.getAccount()) || null;

    return {
      ExecutedAt: scan.startedAt || new Date().toISOString(),
      TenantId: scan.tenantId || '',
      TenantName: scan.tenantName || scan.tenantDomain || '',
      Domain: scan.tenantDomain || '',
      Account: account?.username || scan.account || '',
      CurrentVersion: version,
      LatestVersion: version,
      TestResultSummary: summarize(results),
      Tests: results.map(r => buildTestEntry(r, metaById.get(`${r.id}`), mdById.get(`${r.id}`))),
      TenantInfo: scan.tenantInfo || {},
      EndOfJson: 'EndOfJson',
    };
  }

  async function buildReportHtml(scan) {
    const data = await buildAssessmentJson(scan);

    const res = await fetch(TEMPLATE_URL, { cache: 'force-cache' });
    if (!res.ok) {
      throw new Error(`Could not load the bundled report template (${res.status}). Ensure the powershell folder is deployed alongside the app.`);
    }
    const template = await res.text();

    const start = template.indexOf(START_MARKER);
    const endIdx = template.indexOf(END_MARKER);
    if (start < 0 || endIdx < 0) {
      throw new Error('Report template markers not found. The bundled ReportTemplate.html format has changed; re-sync the powershell folder.');
    }
    const end = endIdx + END_MARKER.length;

    // `</` must be escaped so embedded JSON can never terminate the template's script tag.
    const json = JSON.stringify(data).replace(/<\//g, '<\\/');
    return `${template.substring(0, start)}reportData= ${json}${template.substring(end)}`;
  }

  function download(filename, content, type) {
    const blob = new Blob([content], { type: type || 'text/html' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  async function exportReport(scan) {
    if (!scan) throw new Error('No scan selected.');
    const html = await buildReportHtml(scan);
    const stamp = (scan.startedAt || new Date().toISOString()).slice(0, 10);
    const tenant = (scan.tenantDomain || scan.tenantId || 'tenant').replace(/[^a-z0-9.-]/gi, '_');
    download(`ZeroTrustAssessmentReport-${tenant}-${stamp}.html`, html);
  }

  window.ReportOfficial = {
    buildAssessmentJson,
    buildReportHtml,
    export: exportReport,
  };
})();
