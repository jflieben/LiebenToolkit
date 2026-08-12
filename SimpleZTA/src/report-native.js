(() => {
  // SimpleZTA-branded standalone HTML report (single self-contained file, no external
  // resources) and the Zero Trust Workshop JSON export (port of
  // Convert-ZtAssessmentToWorkshop.ps1 using the module's ztw-task-mapping.json).

  const PILLARS = ['Identity', 'Devices', 'Network', 'Data', 'Infrastructure', 'SecOps', 'AI'];
  const RISK_ORDER = ['Critical', 'High', 'Medium', 'Low', 'Unranked'];
  const COST_ORDER = { Low: 0, Medium: 1, High: 2 };

  function esc(v) {
    return Md.escapeHtml(v);
  }

  function scored(r) {
    return r.implementationLevel !== 'none' && r.status !== 'Skipped';
  }

  // Inline stylesheet for the exported document: same viz tokens as the app (light + dark
  // via prefers-color-scheme) plus a print-friendly layout. Kept self-contained by design.
  const REPORT_CSS = `
  :root {
    --bg:#f9f9f7; --card:#fcfcfb; --text:#0b0b0b; --muted:#52514e; --border:#e1e0d9; --accent:#2a78d6;
    --viz-s1:#2a78d6; --viz-s2:#1baf7a; --viz-s3:#eda100; --viz-s4:#008300; --viz-s5:#4a3aa7;
    --viz-s6:#e34948; --viz-s7:#e87ba4; --viz-s8:#eb6834;
    --viz-good:#0ca30c; --viz-warning:#fab219; --viz-serious:#ec835a; --viz-critical:#d03b3b;
    --viz-neutral:#898781; --viz-grid:#e1e0d9; --viz-axis:#c3c2b7;
    --viz-seq-1:#cde2fb; --viz-seq-2:#9ec5f4; --viz-seq-3:#6da7ec; --viz-seq-4:#2a78d6; --viz-seq-5:#1c5cab;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg:#0d0d0d; --card:#1a1a19; --text:#ffffff; --muted:#c3c2b7; --border:#2c2c2a; --accent:#3987e5;
      --viz-s1:#3987e5; --viz-s2:#199e70; --viz-s3:#c98500; --viz-s5:#9085e9; --viz-s6:#e66767;
      --viz-s7:#d55181; --viz-s8:#d95926;
      --viz-grid:#2c2c2a; --viz-axis:#383835;
      --viz-seq-1:#184f95; --viz-seq-2:#256abf; --viz-seq-3:#3987e5; --viz-seq-4:#6da7ec; --viz-seq-5:#9ec5f4;
    }
  }
  * { box-sizing: border-box; }
  body { margin:0; font-family: system-ui, -apple-system, "Segoe UI", sans-serif; background: var(--bg); color: var(--text); }
  main { max-width: 1080px; margin: 0 auto; padding: 24px 16px 48px; }
  h1 { font-size: 26px; margin: 0 0 4px; }
  h2 { font-size: 19px; margin: 28px 0 10px; border-bottom: 1px solid var(--border); padding-bottom: 6px; }
  h3 { font-size: 14px; margin: 0 0 8px; }
  .muted { color: var(--muted); }
  .card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 14px; margin: 12px 0; break-inside: avoid; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 12px; }
  .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 8px; }
  .kpi { border: 1px solid var(--border); border-radius: 8px; padding: 10px; }
  .kpi .label { color: var(--muted); font-size: 12px; }
  .kpi .value { font-size: 22px; font-weight: 700; }
  .chart-card { background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 10px; break-inside: avoid; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th, td { border-bottom: 1px solid var(--border); padding: 7px 8px; text-align: left; vertical-align: top; }
  th { color: var(--muted); font-weight: 600; }
  .md-table-wrap { overflow-x: auto; }
  .viz-gauge-row { display: flex; flex-wrap: wrap; gap: 16px; }
  .viz-gauge { text-align: center; }
  .viz-gauge-value { font-size: 20px; font-weight: 700; fill: var(--text); }
  .viz-gauge-sub { font-size: 10px; fill: var(--muted); }
  .viz-gauge-label { font-size: 13px; font-weight: 600; margin-top: 2px; }
  .viz-donut { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
  .viz-legend { display: flex; flex-direction: column; gap: 4px; }
  .viz-legend-row { display: flex; align-items: center; gap: 6px; font-size: 12px; }
  .viz-swatch { width: 10px; height: 10px; border-radius: 3px; display: inline-block; flex: none; }
  .viz-legend-value { color: var(--muted); margin-left: auto; padding-left: 10px; }
  .viz-sankey svg, .viz-trend svg { width: 100%; height: auto; }
  .viz-node-label { font-size: 11px; fill: var(--text); paint-order: stroke; stroke: var(--card); stroke-width: 3px; stroke-linejoin: round; }
  .viz-node-value { fill: var(--muted); }
  .viz-caption { margin: 6px 0 0; font-size: 12px; color: var(--muted); }
  .viz-axis-label { font-size: 10px; fill: var(--muted); }
  .viz-empty { color: var(--muted); font-size: 13px; padding: 18px 8px; text-align: center; border: 1px dashed var(--border); border-radius: 8px; }
  .viz-heatmap { border-collapse: collapse; font-size: 12px; }
  .viz-heat-cell { padding: 8px 10px; text-align: center; border: 2px solid var(--card); border-radius: 6px; min-width: 44px; color: #0b0b0b; font-weight: 600; }
  .viz-heat-dark { color: #fff; }
  .viz-heat-zero { color: var(--muted); background: transparent; font-weight: 400; }
  .viz-risk { font-size: 11px; font-weight: 600; border-radius: 999px; padding: 2px 8px; border: 1px solid var(--border); white-space: nowrap; }
  .viz-risk-critical { border-color: var(--viz-critical); color: var(--viz-critical); }
  .viz-risk-high { border-color: var(--viz-serious); color: var(--viz-serious); }
  .viz-risk-medium { border-color: var(--viz-warning); color: var(--muted); }
  .viz-risk-low, .viz-risk-unranked { color: var(--muted); }
  .status-Passed { color: var(--viz-good); font-weight: 600; }
  .status-Failed { color: var(--viz-critical); font-weight: 600; }
  .status-Error { color: var(--viz-serious); font-weight: 600; }
  .status-Skipped, .status-Planned { color: var(--viz-neutral); font-weight: 600; }
  .md-body { line-height: 1.5; font-size: 13px; }
  .md-body ul, .md-body ol { margin: 6px 0; padding-left: 20px; }
  .md-body a { color: var(--accent); }
  .md-body pre { background: rgba(127,127,127,0.1); border: 1px solid var(--border); border-radius: 8px; padding: 10px; overflow: auto; }
  .finding { border-left: 3px solid var(--viz-critical); padding-left: 12px; margin: 16px 0; break-inside: avoid; }
  footer { margin-top: 32px; font-size: 12px; color: var(--muted); border-top: 1px solid var(--border); padding-top: 12px; }
  @media print {
    body { background: #fff; }
    .card, .chart-card { border-color: #ccc; box-shadow: none; }
    a { color: inherit; }
  }`;

  function pillarGauges(results) {
    const present = PILLARS.filter(p => results.some(r => r.pillar === p));
    if (!present.length) return '';
    return `<div class="card"><h3>Pillar scores</h3><div class="viz-gauge-row">${present.map(p => {
      const inPillar = results.filter(r => r.pillar === p);
      const total = inPillar.filter(scored).length;
      const passed = inPillar.filter(r => scored(r) && r.status === 'Passed').length;
      return Charts.gauge({ label: p, value: passed, total });
    }).join('')}</div></div>`;
  }

  function statusDonut(results) {
    const counts = {};
    for (const r of results) {
      const s = r.implementationLevel === 'none' ? 'Planned' : r.status;
      counts[s] = (counts[s] || 0) + 1;
    }
    const segments = ['Passed', 'Failed', 'Error', 'Investigate', 'Skipped', 'Planned']
      .filter(s => counts[s]).map(s => ({ label: s, value: counts[s], color: Charts.statusColor(s) }));
    return `<div class="chart-card"><h3>Result status</h3>${Charts.donut({ centerLabel: 'controls', segments })}</div>`;
  }

  function riskHeatmap(results) {
    const failed = results.filter(r => r.status === 'Failed');
    if (!failed.length) return '';
    const pillars = PILLARS.filter(p => results.some(r => r.pillar === p));
    const risks = RISK_ORDER.filter(risk => failed.some(r => (r.risk || 'Unranked') === risk));
    return `<div class="chart-card"><h3>Failed controls by pillar × risk</h3>${Charts.heatmap({
      rows: pillars,
      cols: risks,
      get: (p, risk) => failed.filter(r => r.pillar === p && (r.risk || 'Unranked') === risk).length,
    })}</div>`;
  }

  function tenantKpis(info) {
    const o = info?.TenantOverview;
    if (!o) return '';
    const t = (label, v) => `<div class="kpi"><div class="label">${esc(label)}</div><div class="value">${esc(v?.toLocaleString?.() ?? v ?? '-')}</div></div>`;
    return `<div class="card"><h3>Tenant overview</h3><div class="kpis">
      ${t('Users', o.UserCount)}${t('Guests', o.GuestCount)}${t('Groups', o.GroupCount)}
      ${t('Applications', o.ApplicationCount)}${t('Devices', o.DeviceCount)}${t('Managed devices', o.ManagedDeviceCount)}
    </div></div>`;
  }

  function sankeys(info) {
    if (!info) return '';
    const card = (title, data) => data
      ? `<div class="chart-card"><h3>${esc(title)}</h3>${Charts.sankey({ title, data })}</div>`
      : '';
    const cards = [
      card('Conditional access → MFA coverage', info.OverviewCaMfaAllUsers),
      card('Sign-ins by device management state', info.OverviewCaDevicesAllUsers),
      card('Auth methods — all users', info.OverviewAuthMethodsAllUsers),
      card('Auth methods — privileged users', info.OverviewAuthMethodsPrivilegedUsers),
      card('Desktop devices', info.DeviceOverview?.DesktopDevicesSummary),
      card('Mobile devices', info.DeviceOverview?.MobileSummary),
    ].filter(Boolean).join('');
    return cards ? `<h2>Tenant posture graphics</h2><div class="grid">${cards}</div>` : '';
  }

  function rankFailed(results) {
    const metaById = new Map((window.__allTests || []).map(t => [t.id, t]));
    return results.filter(r => r.status === 'Failed')
      .map(r => ({ r, meta: metaById.get(`${r.id}`) || {} }))
      .sort((a, b) => {
        const risk = RISK_ORDER.indexOf(a.r.risk || 'Unranked') - RISK_ORDER.indexOf(b.r.risk || 'Unranked');
        if (risk) return risk;
        return (COST_ORDER[a.meta.implementationCost] ?? 3) - (COST_ORDER[b.meta.implementationCost] ?? 3);
      });
  }

  function findings(results, mdById) {
    const ranked = rankFailed(results);
    if (!ranked.length) return '<p class="muted">No failed controls — nothing to remediate. 🎉</p>';
    return ranked.map(({ r, meta }) => {
      const md = mdById.get(`${r.id}`);
      const guidance = md?.description ? `<div class="md-body">${Md.mdToHtml(md.description)}</div>` : '';
      return `
        <div class="finding">
          <div><span class="viz-risk viz-risk-${esc((r.risk || 'Unranked').toLowerCase())}">${esc(r.risk || 'Unranked')}</span>
            <strong> ${esc(r.id)} ${esc(r.title)}</strong>
            <span class="muted"> · ${esc(r.pillar || '')}${meta.implementationCost ? ` · effort ${esc(meta.implementationCost)}` : ''}</span></div>
          <p>${esc(r.message || '')}</p>
          ${guidance}
        </div>`;
    }).join('');
  }

  function appendixTable(results) {
    const rows = [...results].sort((a, b) => Number(a.id) - Number(b.id)).map(r => `
      <tr>
        <td class="status-${esc(r.implementationLevel === 'none' ? 'Planned' : r.status)}">${esc(r.implementationLevel === 'none' ? 'Planned' : r.status)}</td>
        <td>${esc(r.id)}</td><td>${esc(r.title)}</td><td>${esc(r.pillar || '')}</td><td>${esc(r.risk || '')}</td>
      </tr>`).join('');
    return `<div class="md-table-wrap"><table>
      <thead><tr><th>Status</th><th>ID</th><th>Title</th><th>Pillar</th><th>Risk</th></tr></thead>
      <tbody>${rows}</tbody></table></div>`;
  }

  async function buildHtml(scan) {
    const results = scan.results || [];
    const failedIds = results.filter(r => r.status === 'Failed').map(r => r.id);
    const mdById = await Md.loadTestMdMany(failedIds);
    const total = results.filter(scored).length;
    const passed = results.filter(r => scored(r) && r.status === 'Passed').length;
    const score = total ? Math.round((100 * passed) / total) : 0;

    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Zero Trust Assessment — ${esc(scan.tenantName || scan.tenantDomain || 'tenant')}</title>
<style>${REPORT_CSS}</style>
</head>
<body>
<main>
  <h1>Zero Trust Assessment</h1>
  <div class="muted">${esc(scan.tenantName || '')} · ${esc(scan.tenantDomain || '')} · ${esc(scan.tenantId || '')}</div>
  <div class="muted">Scanned ${new Date(scan.startedAt).toLocaleString()} · ${results.length} controls evaluated · overall pass score <strong>${score}%</strong></div>

  <h2>Executive summary</h2>
  ${pillarGauges(results)}
  <div class="grid">
    ${statusDonut(results)}
    ${riskHeatmap(results)}
  </div>
  ${tenantKpis(scan.tenantInfo)}
  ${sankeys(scan.tenantInfo)}

  <h2>Findings &amp; remediation (${failedIds.length} failed controls, highest risk first)</h2>
  ${findings(results, mdById)}

  <h2>Appendix — all results</h2>
  ${appendixTable(results)}

  <footer>
    Generated by SimpleZTA (Lieben Consultancy) — a browser-native companion to the Microsoft ZeroTrustAssessment
    PowerShell module (© Microsoft Corporation). Control metadata and remediation guidance are sourced from the module.
    This report contains sensitive tenant configuration data — share only with authorized personnel.
  </footer>
</main>
</body>
</html>`;
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
    const html = await buildHtml(scan);
    const stamp = (scan.startedAt || new Date().toISOString()).slice(0, 10);
    const tenant = (scan.tenantDomain || scan.tenantId || 'tenant').replace(/[^a-z0-9.-]/gi, '_');
    download(`SimpleZTA-report-${tenant}-${stamp}.html`, html);
  }

  // ---- Zero Trust Workshop export (port of Convert-ZtAssessmentToWorkshop.ps1) ------------

  const KNOWN_WORKSHOP_PILLARS = ['identity', 'devices', 'data', 'network', 'infrastructure', 'security-ops', 'ai'];

  async function exportWorkshop(scan) {
    if (!scan) throw new Error('No scan selected.');

    let mapping = null;
    try {
      const res = await fetch('./powershell/assets/ztw-task-mapping.json', { cache: 'force-cache' });
      if (res.ok) mapping = await res.json();
    } catch {
      mapping = null;
    }

    const pillars = {};
    for (const p of KNOWN_WORKSHOP_PILLARS) pillars[p] = { taskOverrides: {} };

    const collectedNotes = new Map(); // "pillar|taskId" -> Set of findings
    let modifiedCount = 0;

    for (const r of scan.results || []) {
      if (r.status === 'Skipped' || r.implementationLevel === 'none') continue;
      const pillarNames = [r.pillar].filter(Boolean);
      if (!pillarNames.length) continue;

      const headline = `${r.message || ''}`.split('\n').map(l => l.trim()).find(l => l.length) || '';

      for (const pillarName of pillarNames) {
        const pillarKey = `${pillarName}`.toLowerCase();
        let overrideIds;
        if (mapping) {
          const pillarMap = mapping[pillarKey] || mapping[pillarName] || null;
          const mapped = pillarMap ? pillarMap[`${r.id}`] : null;
          if (!mapped) continue; // not mapped to a Workshop task in this pillar
          overrideIds = Array.from(new Set([].concat(mapped).map(x => `${x}`).filter(Boolean)));
        } else {
          overrideIds = [`${r.id}`];
        }

        if (!pillars[pillarKey]) pillars[pillarKey] = { taskOverrides: {} };
        for (const overrideId of overrideIds) {
          if (headline) {
            const key = `${pillarKey}|${overrideId}`;
            if (!collectedNotes.has(key)) collectedNotes.set(key, new Set());
            collectedNotes.get(key).add(headline);
          }
          if (!pillars[pillarKey].taskOverrides[overrideId]) {
            pillars[pillarKey].taskOverrides[overrideId] = { status: 'not-reviewed', notes: '' };
            modifiedCount += 1;
          }
        }
      }
    }

    for (const [key, notes] of collectedNotes.entries()) {
      const [pKey, oKey] = key.split('|');
      pillars[pKey].taskOverrides[oKey].notes = `ZT Assessment result:\n${Array.from(notes).join('\n')}\n`;
    }

    const timestamp = new Date().toISOString().replace(/\.\d+Z$/, m => `.${m.slice(1, 4)}Z`);
    const output = {
      metadata: {
        version: '1.0.0',
        formatVersion: '1.0',
        exportedAt: timestamp,
        applicationVersion: '1.0.0',
        exportType: 'full-configuration',
        scope: 'all',
        description: 'Zero Trust Assessment Result Export',
      },
      configuration: {
        applicationState: { currentPillar: 'identity', lastModified: timestamp },
        pillars,
        globalSettings: { preferences: { autoSave: true, confirmationDialogs: true } },
      },
      statistics: {
        totalTasks: (scan.results || []).length,
        modifiedTasks: modifiedCount,
        completedTasks: 0,
        inProgressTasks: 0,
        plannedTasks: 0,
        pillarsWithChanges: Object.keys(pillars).filter(p => Object.keys(pillars[p].taskOverrides).length),
      },
    };

    const stamp = (scan.startedAt || new Date().toISOString()).slice(0, 10);
    download(`ZeroTrustWorkshop-${stamp}.json`, JSON.stringify(output, null, 2), 'application/json');
  }

  window.ReportNative = {
    export: exportReport,
    buildHtml,
    exportWorkshop,
  };
})();
