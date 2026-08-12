(() => {
  // History tab: pass-score trend lines (overall + per pillar), posture metrics over time
  // (from stored tenantInfo aggregates), per-test movement, and a scan-to-scan diff view.

  const PILLARS = ['Identity', 'Devices', 'Network', 'Data', 'Infrastructure', 'SecOps', 'AI'];
  const PILLAR_COLORS = ['var(--viz-s1)', 'var(--viz-s2)', 'var(--viz-s3)', 'var(--viz-s5)', 'var(--viz-s6)', 'var(--viz-s7)', 'var(--viz-s8)'];

  function esc(v) {
    return Md.escapeHtml(v);
  }

  function scoreOf(results, pillar) {
    const set = (results || []).filter(r => (!pillar || r.pillar === pillar) && r.implementationLevel !== 'none' && r.status !== 'Skipped');
    if (!set.length) return null;
    const passed = set.filter(r => r.status === 'Passed').length;
    return Math.round((100 * passed) / set.length);
  }

  // Extracts a "% of flow that reached the good terminal node" metric from a stored sankey.
  function sankeyPercent(data, goodTarget, totalSources) {
    if (!data?.nodes) return null;
    const good = data.nodes.filter(n => n.target === goodTarget).reduce((a, n) => a + (Number(n.value) || 0), 0);
    const total = data.nodes.filter(n => totalSources.includes(n.source)).reduce((a, n) => a + (Number(n.value) || 0), 0);
    if (!total) return null;
    return Math.round((100 * good) / total);
  }

  function buildTrend(scans) {
    if (!scans.length) return '<p class="muted">No history yet.</p>';
    const ordered = [...scans].sort((a, b) => a.startedAt.localeCompare(b.startedAt));
    const labels = ordered.map(s => new Date(s.startedAt).toLocaleDateString());

    const series = [];
    const overall = ordered.map((s, i) => ({ label: labels[i], value: scoreOf(s.results) ?? 0 }));
    series.push({ name: 'Overall', color: 'var(--text)', points: overall });
    PILLARS.forEach((p, idx) => {
      const points = ordered.map((s, i) => ({ label: labels[i], value: scoreOf(s.results, p) }))
        .map(pt => ({ ...pt, value: pt.value == null ? 0 : pt.value }));
      if (ordered.some(s => (s.results || []).some(r => r.pillar === p))) {
        series.push({ name: p, color: PILLAR_COLORS[idx], points });
      }
    });

    const scoreChart = `
      <div class="chart-card">
        <h3>Pass score over time</h3>
        ${Charts.trend({ title: 'Pass score over time', series, yMax: 100, unit: '%' })}
      </div>`;

    // Posture metrics from tenant insights, when at least two scans collected them.
    const postureSeries = [];
    const withInfo = ordered.filter(s => s.tenantInfo);
    if (withInfo.length >= 2) {
      const metric = (name, color, get) => {
        const points = withInfo
          .map(s => ({ label: new Date(s.startedAt).toLocaleDateString(), value: get(s.tenantInfo) }))
          .filter(p => p.value != null);
        if (points.length >= 2) postureSeries.push({ name, color, points });
      };
      metric('MFA-protected sign-ins', 'var(--viz-s1)', info => sankeyPercent(info.OverviewCaMfaAllUsers, 'MFA', ['User sign in']));
      metric('Compliant-device sign-ins', 'var(--viz-s2)', info => sankeyPercent(info.OverviewCaDevicesAllUsers, 'Compliant', ['User sign in']));
      metric('Phish-resistant users', 'var(--viz-s5)', info => {
        const nodes = info.OverviewAuthMethodsAllUsers?.nodes;
        if (!nodes) return null;
        const val = t => nodes.filter(n => n.source === 'Users' && n.target === t).reduce((a, n) => a + (Number(n.value) || 0), 0);
        const total = val('Single factor') + val('Phishable') + val('Phish resistant');
        return total ? Math.round((100 * val('Phish resistant')) / total) : null;
      });
    }
    const postureChart = postureSeries.length
      ? `<div class="chart-card">
          <h3>Posture metrics over time</h3>
          ${Charts.trend({ title: 'Posture metrics over time', series: postureSeries, yMax: 100, unit: '%' })}
          <p class="viz-caption">Derived from the tenant insight aggregates stored with each scan.</p>
        </div>`
      : '';

    const rows = ordered.map((s, i) => `
      <tr>
        <td>${new Date(s.startedAt).toLocaleString()}</td>
        <td>${esc(s.tenantDomain || s.tenantId || '')}</td>
        <td>${scoreOf(s.results) ?? '-'}%</td>
        <td>${s.summary.passed}/${s.summary.total}</td>
        <td>${s.summary.failed}</td>
      </tr>`).join('');

    return `
      <div class="results-charts">${scoreChart}${postureChart}</div>
      <table class="results-table"><thead><tr><th>Run</th><th>Tenant</th><th>Pass score</th><th>Passed</th><th>Failed</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  function statusDot(status, at) {
    const color = Charts.statusColor(status);
    return `<span class="hist-dot" style="background:${color}" title="${esc(status)} — ${new Date(at).toLocaleString()}"></span>`;
  }

  function renderPerTest(scans) {
    const ordered = [...scans].sort((a, b) => a.startedAt.localeCompare(b.startedAt));
    const map = new Map();
    for (const scan of ordered) {
      for (const r of scan.results || []) {
        if (!map.has(r.id)) map.set(r.id, { id: r.id, title: r.title, items: [] });
        map.get(r.id).items.push({ status: r.status, at: scan.startedAt });
      }
    }

    const q = (document.getElementById('historySearch').value || '').toLowerCase();
    const rows = Array.from(map.values())
      .filter(x => `${x.id} ${x.title}`.toLowerCase().includes(q))
      .sort((a, b) => Number(a.id) - Number(b.id))
      .map(x => {
        const total = x.items.length;
        const passed = x.items.filter(i => i.status === 'Passed').length;
        const last = x.items[x.items.length - 1]?.status || 'Unknown';
        const recent = x.items.slice(-12).map(i => statusDot(i.status, i.at)).join('');
        const rate = total ? Math.round((100 * passed) / total) : 0;
        const spark = Charts.sparkline(x.items.map(i => (i.status === 'Passed' ? 1 : 0)));
        return `<tr><td>${esc(x.id)} ${esc(x.title)}</td><td class="status-${esc(last)}">${esc(last)}</td><td>${rate}% ${spark}</td><td>${total}</td><td class="hist-dots">${recent}</td></tr>`;
      }).join('');

    document.querySelector('#historyTable tbody').innerHTML = rows;
  }

  // ---- Scan diff -------------------------------------------------------------------------

  function scanLabel(s) {
    return `${new Date(s.startedAt).toLocaleString()} | ${s.tenantDomain || s.tenantId || 'tenant'} | ${s.summary.failed}F/${s.summary.passed}P`;
  }

  function renderDiffPickers(scans) {
    const ordered = [...scans].sort((a, b) => b.startedAt.localeCompare(a.startedAt));
    const from = document.getElementById('diffFromSelect');
    const to = document.getElementById('diffToSelect');
    if (!ordered.length) {
      from.innerHTML = to.innerHTML = '<option value="">(no scans)</option>';
      document.getElementById('diffRoot').innerHTML = '<p class="muted">Run at least two scans to compare.</p>';
      return;
    }
    const options = ordered.map(s => `<option value="${s.id}">${esc(scanLabel(s))}</option>`).join('');
    from.innerHTML = options;
    to.innerHTML = options;
    to.value = ordered[0].id;
    from.value = (ordered[1] || ordered[0]).id;
    renderDiff();
  }

  function renderDiff() {
    const fromId = document.getElementById('diffFromSelect').value;
    const toId = document.getElementById('diffToSelect').value;
    const root = document.getElementById('diffRoot');
    const a = Store.getScan(fromId);
    const b = Store.getScan(toId);
    if (!a || !b || a.id === b.id) {
      root.innerHTML = '<p class="muted">Select two different scans to compare.</p>';
      return;
    }

    const aById = new Map((a.results || []).map(r => [r.id, r]));
    const bById = new Map((b.results || []).map(r => [r.id, r]));
    const regressions = [];
    const fixes = [];
    const changed = [];
    const added = [];
    const removed = [];

    for (const r of b.results || []) {
      const before = aById.get(r.id);
      if (!before) { added.push(r); continue; }
      if (before.status === r.status) continue;
      if (before.status === 'Passed' && ['Failed', 'Error'].includes(r.status)) regressions.push({ before, after: r });
      else if (['Failed', 'Error'].includes(before.status) && r.status === 'Passed') fixes.push({ before, after: r });
      else changed.push({ before, after: r });
    }
    for (const r of a.results || []) {
      if (!bById.has(r.id)) removed.push(r);
    }

    const section = (title, cssClass, rows, render) => rows.length
      ? `<h3 class="${cssClass}">${esc(title)} (${rows.length})</h3>
         <table class="results-table"><tbody>${rows.map(render).join('')}</tbody></table>`
      : '';

    const pair = ({ before, after }) => `
      <tr>
        <td><strong>${esc(after.id)}</strong> ${esc(after.title)}</td>
        <td class="status-${esc(before.status)}">${esc(before.status)}</td>
        <td>→</td>
        <td class="status-${esc(after.status)}">${esc(after.status)}</td>
      </tr>`;
    const single = r => `<tr><td><strong>${esc(r.id)}</strong> ${esc(r.title)}</td><td class="status-${esc(r.status)}">${esc(r.status)}</td></tr>`;

    const scoreA = scoreOf(a.results);
    const scoreB = scoreOf(b.results);
    const delta = scoreA != null && scoreB != null ? scoreB - scoreA : null;
    const deltaChip = delta == null ? '' : `<span class="viz-chip ${delta > 0 ? 'viz-chip-good' : delta < 0 ? 'viz-chip-bad' : ''}">Pass score ${scoreA}% → ${scoreB}% (${delta >= 0 ? '+' : ''}${delta})</span>`;

    root.innerHTML = `
      <div class="row">${deltaChip}
        <span class="viz-chip ${regressions.length ? 'viz-chip-bad' : ''}">▼ ${regressions.length} regressed</span>
        <span class="viz-chip ${fixes.length ? 'viz-chip-good' : ''}">▲ ${fixes.length} fixed</span>
      </div>
      ${section('Regressions', 'status-Failed', regressions, pair)}
      ${section('Fixed', 'status-Passed', fixes, pair)}
      ${section('Other status changes', 'muted', changed, pair)}
      ${section('Newly scanned controls', 'muted', added, single)}
      ${section('No longer scanned', 'muted', removed, single)}
      ${!regressions.length && !fixes.length && !changed.length && !added.length && !removed.length ? '<p class="muted">No differences between these scans.</p>' : ''}
    `;
  }

  function refresh() {
    const scans = Store.loadScans();
    document.getElementById('historySummary').innerHTML = buildTrend(scans);
    Charts.attachTooltips(document.getElementById('historySummary'));
    renderPerTest(scans);
    renderDiffPickers(scans);
  }

  function init() {
    document.getElementById('historySearch').addEventListener('input', () => renderPerTest(Store.loadScans()));
    document.getElementById('diffFromSelect').addEventListener('change', renderDiff);
    document.getElementById('diffToSelect').addEventListener('change', renderDiff);
    document.getElementById('clearHistoryBtn').addEventListener('click', () => {
      if (!confirm('Delete all local scan history from this browser?')) return;
      Store.clearScans();
      refresh();
      UIResults.refreshPicker();
    });
  }

  window.UIHistory = { init, refresh };
})();
