(() => {
  // Dashboard tab: pillar scorecards, tenant KPIs, Zero Trust sankeys, risk heatmap,
  // top-fixes panel and Intune configuration tables. Mirrors the graphics of the official
  // Microsoft report, fed by the scan's results + tenantInfo (see src/tenantinfo.js).

  const PILLARS = ['Identity', 'Devices', 'Network', 'Data', 'Infrastructure', 'SecOps', 'AI'];
  const RISK_ORDER = ['Critical', 'High', 'Medium', 'Low', 'Unranked'];
  const COST_ORDER = { Low: 0, Medium: 1, High: 2 };

  function esc(v) {
    return Md.escapeHtml(v);
  }

  function scored(r) {
    if (r.implementationLevel === 'none') return false; // Planned
    return !['Skipped'].includes(r.status);
  }

  function currentScan() {
    const selected = window.UIResults?.getCurrentScanId?.();
    if (selected) {
      const scan = Store.getScan(selected);
      if (scan) return scan;
    }
    const scans = Store.loadScans().sort((a, b) => b.startedAt.localeCompare(a.startedAt));
    return scans[0] || null;
  }

  function previousScan(scan) {
    const scans = Store.loadScans()
      .filter(s => s.id !== scan.id && s.startedAt < scan.startedAt)
      .filter(s => !scan.tenantId || !s.tenantId || s.tenantId === scan.tenantId)
      .sort((a, b) => b.startedAt.localeCompare(a.startedAt));
    return scans[0] || null;
  }

  function headerCard(scan) {
    const prev = previousScan(scan);
    let deltaChips = '';
    if (prev) {
      const prevById = new Map((prev.results || []).map(r => [r.id, r.status]));
      let regressions = 0;
      let fixes = 0;
      for (const r of scan.results || []) {
        const before = prevById.get(r.id);
        if (before === 'Passed' && r.status === 'Failed') regressions += 1;
        if (before === 'Failed' && r.status === 'Passed') fixes += 1;
      }
      deltaChips = `
        <span class="viz-chip ${regressions ? 'viz-chip-bad' : ''}" title="Controls that passed in the previous scan but fail now">▼ ${regressions} regressed</span>
        <span class="viz-chip ${fixes ? 'viz-chip-good' : ''}" title="Controls that failed in the previous scan but pass now">▲ ${fixes} fixed</span>
        <span class="viz-chip">vs ${new Date(prev.startedAt).toLocaleDateString()}</span>`;
    }
    return `
      <div class="card">
        <div class="row between">
          <div>
            <h2 style="margin:0">${esc(scan.tenantName || scan.tenantDomain || 'Tenant')}</h2>
            <div class="muted">${esc(scan.tenantDomain || '')}${scan.tenantId ? ` · ${esc(scan.tenantId)}` : ''}</div>
          </div>
          <div class="row">${deltaChips}</div>
        </div>
        <div class="muted" style="margin-top:6px">
          Scanned ${new Date(scan.startedAt).toLocaleString()} · ${(scan.results || []).length} controls
          ${scan.tenantInfoMeta?.days || scan.tenantInfo ? '' : ' · tenant insights not collected'}
        </div>
      </div>`;
  }

  function pillarGauges(scan) {
    const results = scan.results || [];
    const present = PILLARS.filter(p => results.some(r => r.pillar === p));
    if (!present.length) return '';
    const gauges = present.map(p => {
      const inPillar = results.filter(r => r.pillar === p);
      const total = inPillar.filter(scored).length;
      const passed = inPillar.filter(r => scored(r) && r.status === 'Passed').length;
      return Charts.gauge({ label: p, value: passed, total });
    }).join('');
    return `<div class="card"><h2>Pillar scores</h2><div class="viz-gauge-row">${gauges}</div></div>`;
  }

  function statusDonut(scan) {
    const counts = {};
    for (const r of scan.results || []) {
      const status = r.implementationLevel === 'none' ? 'Planned' : r.status;
      counts[status] = (counts[status] || 0) + 1;
    }
    const order = ['Passed', 'Failed', 'Error', 'Investigate', 'Skipped', 'Planned'];
    const segments = order.filter(s => counts[s]).map(s => ({ label: s, value: counts[s], color: Charts.statusColor(s) }));
    return `<div class="chart-card"><h3>Result status</h3>${Charts.donut({ title: 'Result status', centerLabel: 'controls', segments })}</div>`;
  }

  function riskHeatmap(scan) {
    const failed = (scan.results || []).filter(r => r.status === 'Failed');
    const pillars = PILLARS.filter(p => (scan.results || []).some(r => r.pillar === p));
    const risks = RISK_ORDER.filter(risk => failed.some(r => (r.risk || 'Unranked') === risk));
    if (!failed.length) return `<div class="chart-card"><h3>Failed controls by pillar × risk</h3><div class="viz-empty">No failed controls. 🎉</div></div>`;
    const html = Charts.heatmap({
      rows: pillars,
      cols: risks.length ? risks : ['Unranked'],
      get: (p, risk) => failed.filter(r => r.pillar === p && (r.risk || 'Unranked') === risk).length,
    });
    return `<div class="chart-card"><h3>Failed controls by pillar × risk</h3>${html}</div>`;
  }

  function topFixes(scan) {
    const failed = (scan.results || []).filter(r => r.status === 'Failed');
    if (!failed.length) return '';
    const metaById = window.__allTests ? new Map(window.__allTests.map(t => [t.id, t])) : new Map();
    const rows = failed
      .map(r => ({ r, meta: metaById.get(`${r.id}`) || {} }))
      .sort((a, b) => {
        const riskDiff = RISK_ORDER.indexOf(a.r.risk || 'Unranked') - RISK_ORDER.indexOf(b.r.risk || 'Unranked');
        if (riskDiff) return riskDiff;
        const costDiff = (COST_ORDER[a.meta.implementationCost] ?? 3) - (COST_ORDER[b.meta.implementationCost] ?? 3);
        if (costDiff) return costDiff;
        return (COST_ORDER[a.meta.userImpact] ?? 3) - (COST_ORDER[b.meta.userImpact] ?? 3);
      })
      .slice(0, 10);

    const body = rows.map(({ r, meta }) => `
      <tr class="viz-fix-row" data-test-id="${esc(r.id)}">
        <td><span class="viz-risk viz-risk-${esc((r.risk || 'Unranked').toLowerCase())}">${esc(r.risk || 'Unranked')}</span></td>
        <td><strong>${esc(r.id)}</strong> ${esc(r.title)}</td>
        <td>${esc(meta.implementationCost || '-')}</td>
        <td>${esc(meta.userImpact || '-')}</td>
      </tr>`).join('');

    return `
      <div class="card">
        <h2>Top fixes — highest risk, lowest effort first</h2>
        <p class="muted">Failed controls ranked by risk, then implementation cost and user impact. Click a row for evidence and remediation guidance.</p>
        <table class="results-table" id="topFixesTable">
          <thead><tr><th>Risk</th><th>Control</th><th>Effort</th><th>User impact</th></tr></thead>
          <tbody>${body}</tbody>
        </table>
      </div>`;
  }

  function tenantKpis(info) {
    const o = info?.TenantOverview;
    if (!o) return '';
    return `
      <div class="card">
        <h2>Tenant overview</h2>
        <div class="summary-grid">
          ${Charts.tile('Users', o.UserCount?.toLocaleString?.() ?? o.UserCount ?? '-')}
          ${Charts.tile('Guests', o.GuestCount?.toLocaleString?.() ?? '-')}
          ${Charts.tile('Groups', o.GroupCount?.toLocaleString?.() ?? '-')}
          ${Charts.tile('Applications', o.ApplicationCount?.toLocaleString?.() ?? '-')}
          ${Charts.tile('Devices', o.DeviceCount?.toLocaleString?.() ?? '-')}
          ${Charts.tile('Managed devices', o.ManagedDeviceCount?.toLocaleString?.() ?? '-')}
        </div>
      </div>`;
  }

  function sankeyCard(title, data, note) {
    const body = data
      ? Charts.sankey({ title, data })
      : `<div class="viz-empty">${esc(note || 'Not collected for this scan.')}</div>`;
    return `<div class="chart-card"><h3>${esc(title)}</h3>${body}</div>`;
  }

  function ownershipDonut(info) {
    const d = info?.DeviceOverview?.DeviceOwnership;
    if (!d) return '';
    const segments = [
      { label: 'Corporate', value: d.corporateCount || 0, color: 'var(--viz-s1)' },
      { label: 'Personal', value: d.personalCount || 0, color: 'var(--viz-s3)' },
    ];
    return `<div class="chart-card"><h3>Managed device ownership</h3>${Charts.donut({ title: 'Device ownership', centerLabel: 'devices', segments })}</div>`;
  }

  function configTable(title, rows, columns) {
    if (!rows || !rows.length) return '';
    const cols = columns || Object.keys(rows[0]);
    const body = rows.map(r => `<tr>${cols.map(c => `<td>${esc(r[c] ?? '')}</td>`).join('')}</tr>`).join('');
    return `
      <details class="card viz-config">
        <summary><strong>${esc(title)}</strong> <span class="muted">(${rows.length})</span></summary>
        <div class="md-table-wrap">
          <table class="results-table">
            <thead><tr>${cols.map(c => `<th>${esc(c.replace(/([a-z])([A-Z])/g, '$1 $2'))}</th>`).join('')}</tr></thead>
            <tbody>${body}</tbody>
          </table>
        </div>
      </details>`;
  }

  function render(scan) {
    const root = document.getElementById('dashboardRoot');
    if (!scan) {
      root.innerHTML = '<div class="card"><h2>Dashboard</h2><p class="muted">No scans yet. Run an assessment first.</p></div>';
      return;
    }

    const info = scan.tenantInfo || null;
    const notes = scan.tenantInfoNotes || {};
    const note = key => notes[key] || (info ? 'No data available (check license/permissions).' : "Tenant insights were not collected. Enable 'Collect tenant insight graphics' on the Run tab.");

    const compliance = info?.ConfigDeviceCompliancePolicies;
    const appProtection = info?.ConfigDeviceAppProtectionPolicies;
    const enrollment = info?.ConfigDeviceEnrollmentRestriction;
    const winEnrollment = info?.ConfigWindowsEnrollment;

    root.innerHTML = `
      ${headerCard(scan)}
      ${pillarGauges(scan)}
      <div class="results-charts">
        ${statusDonut(scan)}
        ${riskHeatmap(scan)}
      </div>
      ${topFixes(scan)}
      ${tenantKpis(info)}
      <div class="card">
        <h2>Identity — sign-in protection &amp; authentication methods</h2>
        <div class="results-charts">
          ${sankeyCard('Conditional access → MFA coverage', info?.OverviewCaMfaAllUsers, note('OverviewCaMfaAllUsers'))}
          ${sankeyCard('Sign-ins by device management state', info?.OverviewCaDevicesAllUsers, note('OverviewCaDevicesAllUsers'))}
          ${sankeyCard('Auth methods — all users', info?.OverviewAuthMethodsAllUsers, note('OverviewAuthMethodsAllUsers'))}
          ${sankeyCard('Auth methods — privileged users', info?.OverviewAuthMethodsPrivilegedUsers, note('OverviewAuthMethodsPrivilegedUsers'))}
        </div>
      </div>
      <div class="card">
        <h2>Devices — join type, compliance &amp; ownership</h2>
        <div class="results-charts">
          ${sankeyCard('Desktop devices', info?.DeviceOverview?.DesktopDevicesSummary, note('DeviceOverview'))}
          ${sankeyCard('Mobile devices', info?.DeviceOverview?.MobileSummary, note('DeviceOverview'))}
          ${ownershipDonut(info)}
        </div>
      </div>
      ${configTable('Device compliance policies', compliance)}
      ${configTable('App protection policies', appProtection)}
      ${configTable('Enrollment platform restrictions', enrollment)}
      ${configTable('Windows automatic enrollment', winEnrollment)}
    `;

    Charts.attachTooltips(root);

    document.querySelectorAll('#topFixesTable .viz-fix-row').forEach(tr => {
      tr.addEventListener('click', () => {
        const id = tr.dataset.testId;
        document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x.dataset.tab === 'results'));
        document.querySelectorAll('.panel').forEach(p => p.classList.toggle('active', p.id === 'panel-results'));
        window.UIResults?.openDetail?.(scan.id, id);
      });
    });
  }

  function refresh() {
    render(currentScan());
  }

  window.UIDashboard = { refresh, render };
})();
