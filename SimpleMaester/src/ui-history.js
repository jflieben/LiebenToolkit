// History UI: renders a stacked-bar trend chart of pass/fail/skip/error per scan.
// Also renders a per-test breakdown: pass rate and a sparkline of the last 10 results.
(() => {
  let _chart = null;

  async function refresh() {
    const all = await History.getAllScans();
    const tenantSel = document.getElementById('historyTenantFilter');
    if (!tenantSel) return;
    const tenants = Array.from(new Set(all.map(s => s.tenantDomain || s.tenantId))).sort();
    const cur = tenantSel.value;
    tenantSel.innerHTML = '<option value="">All tenants</option>' + tenants.map(t => `<option value="${t}">${t}</option>`).join('');
    if (cur) tenantSel.value = cur;

    const filtered = tenantSel.value ? all.filter(s => (s.tenantDomain || s.tenantId) === tenantSel.value) : all;
    filtered.sort((a, b) => a.startedAt.localeCompare(b.startedAt));

    const historyEmpty = document.getElementById('historyEmpty');
    if (historyEmpty) historyEmpty.style.display = filtered.length ? 'none' : '';
    renderChart(filtered);
    renderPerTest(filtered);
    await renderStorageCard();
  }

  async function renderStorageCard() {
    const card = document.getElementById('storageCard');
    const label = document.getElementById('storageLabel');
    const fill = document.getElementById('storageBarFill');
    const detail = document.getElementById('storageDetail');
    if (!card) return;

    let stats;
    try { stats = await History.getStorageStats(); } catch { card.style.display = 'none'; return; }
    if (!stats || (!stats.quotaBytes && !stats.scanCount)) { card.style.display = 'none'; return; }

    card.style.display = '';
    const pct = Math.round((stats.usedPct || 0) * 100);
    const level = pct >= 80 ? 'danger' : pct >= 60 ? 'warn' : '';
    card.className = 'storage-card' + (level ? ' ' + level : '');
    fill.className = 'storage-bar-fill' + (level ? ' ' + level : '');
    fill.style.width = Math.max(1, pct) + '%';
    label.textContent = `Storage - ${pct}% used`;

    const usedMB = (stats.usedBytes / (1024 * 1024)).toFixed(1);
    const quotaMB = stats.quotaBytes > 0 ? (stats.quotaBytes / (1024 * 1024)).toFixed(0) : '?';
    const scanInfo = `${stats.scanCount} scan${stats.scanCount !== 1 ? 's' : ''}`;
    detail.textContent = stats.quotaBytes > 0
      ? `${usedMB} MB of ~${quotaMB} MB used (origin total) - ${scanInfo} stored - ${stats.backend}`
      : `${scanInfo} stored - ${stats.backend}`;
  }

  function renderChart(scans) {
    const ctx = document.getElementById('historyChart');
    if (!ctx) return;
    if (!window.Chart) {
      ctx.parentElement.insertBefore(Object.assign(document.createElement('p'), {
        textContent: 'Chart.js failed to load - chart unavailable.',
        className: 'muted',
      }), ctx);
      return;
    }

    if (_chart) _chart.destroy();

    const labels = scans.map(s => new Date(s.startedAt).toLocaleString());
    const passed = scans.map(s => s.summary.passed);
    const failed = scans.map(s => s.summary.failed);
    const skipped = scans.map(s => s.summary.skipped);
    const errored = scans.map(s => s.summary.error);

    _chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'Passed', data: passed, backgroundColor: '#128d4f' },
          { label: 'Failed', data: failed, backgroundColor: '#c0392b' },
          { label: 'Skipped', data: skipped, backgroundColor: '#8a9bb5' },
          { label: 'Errored', data: errored, backgroundColor: '#8b1a1a' },
        ],
      },
      options: {
        responsive: true,
        scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } },
        plugins: { legend: { position: 'bottom' } },
      },
    });
  }

  function renderPerTest(scans) {
    const tbody = document.querySelector('#historyTestTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    if (!scans.length) return;

    const byId = new Map();
    for (const s of scans) {
      for (const r of s.results) {
        let row = byId.get(r.id);
        if (!row) {
          row = { id: r.id, title: r.title || r.id, severity: r.severity, history: [] };
          byId.set(r.id, row);
        }
        row.history.push({ at: s.startedAt, status: r.status, reason: r.reason });
      }
    }

    const q = (document.getElementById('historyTestSearch').value || '').toLowerCase();
    const rows = [...byId.values()].filter(r => !q || r.id.toLowerCase().includes(q) || (r.title || '').toLowerCase().includes(q));
    rows.sort((a, b) => a.id.localeCompare(b.id));

    for (const r of rows) {
      const last = r.history[r.history.length - 1];
      const passes = r.history.filter(h => h.status === 'Passed').length;
      const evals = r.history.filter(h => h.status === 'Passed' || h.status === 'Failed').length;
      const rate = evals ? Math.round((100 * passes) / evals) + '%' : '-';
      const last10 = r.history.slice(-10);
      const spark = `<div class="spark">${last10.map(h => `<span class="${(h.status || '').toLowerCase()}" style="height:${h.status === 'Passed' ? 14 : h.status === 'Failed' ? 14 : 8}px" title="${h.at}: ${h.status}"></span>`).join('')}</div>`;
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${r.id}<br><span class="muted" style="font-size:11px">${escape(r.title)}</span></td>
        <td><span class="badge ${(last.status || '').toLowerCase()}">${last.status}</span></td>
        <td>${rate}</td><td>${r.history.length}</td><td>${spark}</td>`;
      tr.addEventListener('click', () => UIResults.openDetail({
        id: r.id,
        title: r.title,
        severity: r.severity,
        status: last.status,
        reason: last.reason,
      }));
      tbody.appendChild(tr);
    }
  }

  function escape(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function init() {
    const tenantFilter = document.getElementById('historyTenantFilter');
    const testSearch = document.getElementById('historyTestSearch');
    const clearHistoryBtn = document.getElementById('clearHistoryBtn');
    const pruneHistoryBtn = document.getElementById('pruneHistoryBtn');

    if (tenantFilter) tenantFilter.addEventListener('change', refresh);
    if (testSearch) testSearch.addEventListener('input', refresh);

    if (clearHistoryBtn) clearHistoryBtn.addEventListener('click', async () => {
      if (!confirm('Delete all stored scan history? This cannot be undone.')) return;
      await History.clearAll();
      Log.info('Cleared scan history');
      refresh();
    });

    if (pruneHistoryBtn) pruneHistoryBtn.addEventListener('click', async () => {
      const stats = await History.getStorageStats().catch(() => null);
      const pct = stats ? Math.round((stats.usedPct || 0) * 100) : '?';
      if (!confirm(`Remove the oldest scans to free up storage space? (currently at ${pct}%)`)) return;
      const pruned = await History.pruneOldScans();
      Log.info(`Pruned ${pruned} old scan(s) to free storage space`);
      refresh();
    });

    window.addEventListener('SimpleMaester:storageWarning', () => {
      const topHistory = document.getElementById('top-history');
      if (topHistory && topHistory.classList.contains('active')) {
        renderStorageCard();
      }
    });
  }

  window.UIHistory = { refresh, init };
})();
