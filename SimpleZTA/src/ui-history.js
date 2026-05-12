(() => {
  function buildTrend(scans) {
    if (!scans.length) return '<p class="muted">No history yet.</p>';
    const ordered = scans.sort((a, b) => a.startedAt.localeCompare(b.startedAt));
    const data = ordered.map(s => ({
      at: s.startedAt,
      tenant: s.tenantDomain || s.tenantId || '',
      score: s.summary.total ? Math.round((100 * s.summary.passed) / s.summary.total) : 0,
      passed: s.summary.passed,
      total: s.summary.total,
      failed: s.summary.failed,
    }));

    const maxScore = Math.max(1, ...data.map(d => d.score));
    const chart = `
      <div class="chart-card">
        <h3>Pass score trend</h3>
        ${data.map(d => `
          <div class="bar-row">
            <span class="bar-label">${new Date(d.at).toLocaleDateString()}</span>
            <span class="bar-track"><span class="bar-fill passed" style="width:${Math.round((d.score / maxScore) * 100)}%"></span></span>
            <span class="bar-value">${d.score}%</span>
          </div>
        `).join('')}
      </div>
    `;

    const rows = data.map(d => `
      <tr><td>${new Date(d.at).toLocaleString()}</td><td>${d.tenant}</td><td>${d.score}%</td><td>${d.passed}/${d.total}</td><td>${d.failed}</td></tr>
    `).join('');

    return `${chart}<table class="results-table"><thead><tr><th>Run</th><th>Tenant</th><th>Pass score</th><th>Passed</th><th>Failed</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  function statusChar(status) {
    if (status === 'Passed') return 'P';
    if (status === 'Failed') return 'F';
    if (status === 'Skipped') return 'S';
    return 'E';
  }

  function renderPerTest(scans) {
    const map = new Map();
    for (const scan of scans) {
      for (const r of scan.results) {
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
        const recent = x.items.slice(-10).map(i => statusChar(i.status)).join(' ');
        const rate = total ? Math.round((100 * passed) / total) : 0;
        return `<tr><td>${x.id} ${x.title}</td><td class="status-${last}">${last}</td><td>${rate}%</td><td>${total}</td><td class="spark">${recent}</td></tr>`;
      }).join('');

    document.querySelector('#historyTable tbody').innerHTML = rows;
  }

  function refresh() {
    const scans = Store.loadScans();
    document.getElementById('historySummary').innerHTML = buildTrend(scans);
    renderPerTest(scans);
  }

  function init() {
    document.getElementById('historySearch').addEventListener('input', refresh);
    document.getElementById('clearHistoryBtn').addEventListener('click', () => {
      if (!confirm('Delete all local scan history from this browser?')) return;
      Store.clearScans();
      refresh();
      UIResults.refreshPicker();
    });
  }

  window.UIHistory = { init, refresh };
})();
