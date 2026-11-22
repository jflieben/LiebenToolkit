// History UI: stacked bar chart of pass/fail/skip/error per scan and a per-test
// breakdown with pass rate and the last 10 results.
(() => {
  const { esc, statusBadge, sparkHtml } = SMUi;
  let _chart = null;
  let _lastScans = [];
  const $ = (id) => document.getElementById(id);

  async function refresh() {
    const all = await History.getAllScans();
    const tenantSel = $('historyTenantFilter');
    const tenants = Array.from(new Set(all.map((s) => s.tenantDomain || s.tenantId))).sort();
    const cur = tenantSel.value;
    tenantSel.replaceChildren(new Option('All tenants', ''), ...tenants.map((t) => new Option(t, t)));
    if (cur && tenants.includes(cur)) tenantSel.value = cur;

    const filtered = tenantSel.value ? all.filter((s) => (s.tenantDomain || s.tenantId) === tenantSel.value) : all;
    filtered.sort((a, b) => a.startedAt.localeCompare(b.startedAt));
    _lastScans = filtered;

    $('historyEmpty').hidden = filtered.length > 0;
    $('historyChartWrap').hidden = !filtered.length;
    $('clearHistoryBtn').disabled = !all.length;
    renderChart(filtered);
    renderPerTest(filtered);
    await renderStorage();
  }

  async function renderStorage() {
    let stats = null;
    try { stats = await History.getStorageStats(); } catch { /* unavailable */ }
    const el = $('storageDetail');
    if (!stats) { el.textContent = ''; return; }
    const parts = [`${stats.scanCount} scan${stats.scanCount === 1 ? '' : 's'} stored in this browser (${stats.backend})`];
    if (stats.quotaBytes > 0) parts.push(`${(stats.usedBytes / 1048576).toFixed(1)} MB used by this site of about ${Math.round(stats.quotaBytes / 1048576)} MB`);
    el.textContent = parts.join(', ');
  }

  function cssVar(name, fallback) {
    const v = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return v || fallback;
  }

  function renderChart(scans) {
    const ctx = $('historyChart');
    if (_chart) { _chart.destroy(); _chart = null; }
    if (!scans.length) return;
    if (!window.Chart) { $('chartMissing').hidden = false; return; }
    const ink = cssVar('--ink-500', '#566878');
    const line = cssVar('--line', '#dce5ee');
    _chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: scans.map((s) => new Date(s.startedAt).toLocaleString()),
        datasets: [
          { label: 'Passed', data: scans.map((s) => s.summary.passed), backgroundColor: cssVar('--ok', '#2f7d2e') },
          { label: 'Failed', data: scans.map((s) => s.summary.failed), backgroundColor: cssVar('--danger', '#c0352b') },
          { label: 'Skipped', data: scans.map((s) => s.summary.skipped), backgroundColor: cssVar('--ink-300', '#a9b6c3') },
          { label: 'Error', data: scans.map((s) => s.summary.error), backgroundColor: cssVar('--warn', '#b26a00') },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { stacked: true, ticks: { color: ink }, grid: { color: line } },
          y: { stacked: true, beginAtZero: true, ticks: { color: ink, precision: 0 }, grid: { color: line } },
        },
        plugins: { legend: { position: 'bottom', labels: { color: ink } } },
      },
    });
  }

  function renderPerTest(scans) {
    const tbody = document.querySelector('#historyTestTable tbody');
    tbody.replaceChildren();
    if (!scans.length) return;

    const byId = new Map();
    for (const s of scans) {
      for (const r of s.results) {
        let row = byId.get(r.id);
        if (!row) { row = { id: r.id, title: r.title || r.id, severity: r.severity, tag: r.tag, category: r.category, history: [] }; byId.set(r.id, row); }
        row.history.push({ at: s.startedAt, status: r.status, reason: r.reason, actual: r.actual, expectation: r.expectation, graphCalls: r.graphCalls, tenantId: s.tenantId });
      }
    }

    const q = ($('historyTestSearch').value || '').toLowerCase();
    const rows = [...byId.values()].filter((r) => !q || r.id.toLowerCase().includes(q) || (r.title || '').toLowerCase().includes(q));
    rows.sort((a, b) => a.id.localeCompare(b.id));

    for (const r of rows) {
      const last = r.history[r.history.length - 1];
      const passes = r.history.filter((h) => h.status === 'Passed').length;
      const evals = r.history.filter((h) => h.status === 'Passed' || h.status === 'Failed').length;
      const rate = evals ? `${Math.round((100 * passes) / evals)}%` : '-';
      const tr = document.createElement('tr');
      tr.className = 'clickable-row';
      tr.tabIndex = 0;
      tr.innerHTML = `<td><span class="mono">${esc(r.id)}</span><br><span class="muted small">${esc(r.title)}</span></td>
        <td>${statusBadge(last.status)}</td>
        <td class="num">${esc(rate)}</td><td class="num">${r.history.length}</td><td>${sparkHtml(r.history.slice(-10))}</td>`;
      const open = () => UIResults.openDetail({
        id: r.id, title: r.title, severity: r.severity, tag: r.tag, category: r.category,
        status: last.status, reason: last.reason, actual: last.actual, expectation: last.expectation,
        graphCalls: last.graphCalls, tenantId: last.tenantId,
      });
      tr.addEventListener('click', open);
      tr.addEventListener('keydown', (e) => { if (e.key === 'Enter') open(); });
      tbody.appendChild(tr);
    }
  }

  async function clearHistory() {
    const all = await History.getAllScans();
    if (!all.length) { Toolkit.toast('There is no stored history.', 'info'); return; }
    const tenants = Array.from(new Set(all.map((s) => s.tenantDomain || s.tenantId)));
    const ok = await Toolkit.confirm({
      title: 'Clear scan history',
      body: 'Removes every stored scan from this browser. This cannot be undone. Your tenant is not touched.',
      items: [`${all.length} scan(s)`, ...tenants.map((t) => `Tenant: ${t}`)],
      okLabel: 'Clear history',
      danger: true,
    });
    if (!ok) return;
    await History.clearAll();
    Log.info('Cleared scan history');
    Toolkit.toast('Scan history cleared', 'ok');
    UIResults.hide();
    document.dispatchEvent(new CustomEvent('simplemaester:history-cleared'));
    refresh();
  }

  function init() {
    $('historyTenantFilter').addEventListener('change', refresh);
    $('historyTestSearch').addEventListener('input', () => renderPerTest(_lastScans));
    $('clearHistoryBtn').addEventListener('click', (e) => Toolkit.busy(e.currentTarget, clearHistory));
    document.addEventListener('toolkit:theme', () => { if (_chart) renderChart(_lastScans); });
    window.addEventListener('SimpleMaester:storageWarning', (e) => {
      if (e.detail && e.detail.message) Toolkit.toast(e.detail.message, 'warn');
    });
  }

  window.UIHistory = { refresh, init };
})();
