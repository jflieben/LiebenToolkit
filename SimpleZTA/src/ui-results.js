(() => {
  // Results tab: scan picker, summary, filterable table, detail dialog and exports.
  const { esc, toast, busy } = Toolkit;
  let currentScanId = null;
  let selectedResultId = null;

  const STATUS_BADGE = { Passed: 'ok', Failed: 'danger', Error: 'danger', Investigate: 'warn', Skipped: 'neutral', Planned: 'neutral' };
  function statusBadge(status) {
    return `<span class="badge dot ${STATUS_BADGE[status] || 'neutral'}">${esc(status || 'Unknown')}</span>`;
  }

  function openModal(row) {
    Toolkit.modal({ title: `${row.id} ${row.title}`, html: renderDetail(row), wide: true });
    // The dialog is in the DOM now; load the module guidance into it without blocking.
    const target = document.getElementById('mdGuidance');
    if (target && typeof Md !== 'undefined') {
      Md.loadTestMd(row.id).then(md => {
        if (!target.isConnected) return;
        target.innerHTML = md?.description
          ? `<div class="md-body">${Md.mdToHtml(md.description)}</div>`
          : '<div class="muted">No module guidance available for this control.</div>';
      }).catch(() => {
        target.innerHTML = '<div class="muted">Guidance could not be loaded.</div>';
      });
    }
  }

  function jsonBlock(value) {
    const text = typeof value === 'string' ? value : JSON.stringify(value, null, 2);
    return `<pre class="detail-pre">${esc(text || '')}</pre>`;
  }

  function renderDetail(row) {
    const evidence = row.evidence == null ? '<div class="muted">No evidence captured.</div>' : jsonBlock(row.evidence);
    const trace = Array.isArray(row.apiTrace) && row.apiTrace.length
      ? row.apiTrace.map((entry, index) => `
        <div class="trace-entry">
          <div><strong>Call ${index + 1}</strong> ${esc(entry.method || 'GET')} ${esc(entry.status ?? '')}</div>
          <div class="trace-url">${esc(entry.url || '')}</div>
          ${entry.requestPreview ? `<div class="trace-block"><div class="muted small">Request</div>${jsonBlock(entry.requestPreview)}</div>` : ''}
          ${entry.responsePreview ? `<div class="trace-block"><div class="muted small">Response</div>${jsonBlock(entry.responsePreview)}</div>` : ''}
        </div>
      `).join('')
      : '<div class="muted">No API calls were recorded for this result.</div>';

    const impl = row.implementationLevel === 'partial' && row.implementationNotes
      ? `
      <div class="detail-section">
        <h3>Partial implementation</h3>
        <div><strong>Implemented:</strong> ${esc(row.implementationNotes.implemented || '-')}</div>
        <div><strong>Not implemented:</strong> ${esc(row.implementationNotes.notImplemented || '-')}</div>
      </div>`
      : '';

    const docs = row.documentationUrl
      ? `
      <div class="detail-section">
        <h3>Documentation</h3>
        <div><a href="${esc(Toolkit.safeUrl(row.documentationUrl))}" target="_blank" rel="noopener noreferrer">Guidance root</a></div>
        <div><a href="${esc(Toolkit.safeUrl(row.documentationSearchUrl))}" target="_blank" rel="noopener noreferrer">Search for this control</a></div>
      </div>`
      : '';

    const meta = (window.__allTests || []).find(t => `${t.id}` === `${row.id}`) || {};
    const chips = [
      row.pillar && `Pillar: ${row.pillar}`,
      row.risk && `Risk: ${row.risk}`,
      meta.userImpact && `User impact: ${meta.userImpact}`,
      meta.implementationCost && `Effort: ${meta.implementationCost}`,
      (meta.minimumLicense || []).length && `License: ${(meta.minimumLicense || []).join(', ')}`,
      meta.sfiPillar && `SFI: ${meta.sfiPillar}`,
    ].filter(Boolean).map(c => `<span class="viz-chip">${esc(c)}</span>`).join(' ');

    return `
      <div class="detail-section">
        <div class="row tight">${statusBadge(row.status)} <span class="muted">${esc(fmtMs(row.durationMs))}</span></div>
        <div class="row chips-row">${chips}</div>
      </div>
      <div class="detail-section">
        <h3>Summary</h3>
        <div>${esc(row.message || '')}</div>
      </div>
      <div class="detail-section">
        <h3>What was checked and how to fix it</h3>
        <div id="mdGuidance"><div class="muted">Loading guidance...</div></div>
      </div>
      ${impl}
      ${docs}
      <div class="detail-section">
        <h3>Evidence</h3>
        ${evidence}
      </div>
      <div class="detail-section">
        <h3>API calls</h3>
        ${trace}
      </div>
    `;
  }

  function fmtMs(ms) {
    if (ms == null) return '';
    if (ms < 1000) return `${ms} ms`;
    return `${(ms / 1000).toFixed(1)} s`;
  }

  function kpi(label, value, accent) {
    return `<div class="stat${accent ? ` accent-${accent}` : ''}"><span class="label">${esc(label)}</span><span class="value">${esc(value)}</span></div>`;
  }

  function setSummary(scan) {
    const s = scan.summary || { total: 0, passed: 0, failed: 0, skipped: 0, error: 0 };
    document.getElementById('scanMeta').textContent = `${scan.tenantDomain || scan.tenantId || 'Unknown tenant'} | ${new Date(scan.startedAt).toLocaleString()} | ${scan.results.length} controls`;
    document.getElementById('summaryGrid').innerHTML = [
      kpi('Total', s.total, 'primary'),
      kpi('Passed', s.passed, 'ok'),
      kpi('Failed', s.failed, 'danger'),
      kpi('Skipped', s.skipped),
      kpi('Error', s.error, 'warn'),
      kpi('Implemented', `${scan.implementationCoverage?.implemented || 0}/${scan.implementationCoverage?.totalCatalog || 0}`),
    ].join('');

    const counts = { Passed: 0, Failed: 0, Skipped: 0, Error: 0 };
    for (const r of (scan.results || [])) {
      if (counts[r.status] != null) counts[r.status] += 1;
    }
    const max = Math.max(1, counts.Passed, counts.Failed, counts.Skipped, counts.Error);
    const root = document.getElementById('resultsCharts');
    root.innerHTML = `
      <div class="chart-card">
        <h3>Status distribution</h3>
        ${Object.entries(counts).map(([name, val]) => `
          <div class="bar-row">
            <span class="bar-label">${esc(name)}</span>
            <span class="bar-track"><span class="bar-fill ${esc(name.toLowerCase())}" data-pct="${Math.round((val / max) * 100)}"></span></span>
            <span class="bar-value">${val}</span>
          </div>
        `).join('')}
      </div>
    `;
    root.querySelectorAll('.bar-fill[data-pct]').forEach(el => { el.style.width = `${el.dataset.pct}%`; });
  }

  const RISK_SORT = { Critical: 0, High: 1, Medium: 2, Low: 3, Unranked: 4 };
  const STATUS_SORT = { Failed: 0, Error: 1, Investigate: 2, Passed: 3, Skipped: 4 };
  let sortKey = null;
  let sortDir = 1;

  function filtered(scan) {
    const q = (document.getElementById('resultSearch').value || '').toLowerCase();
    const sf = document.getElementById('statusFilter').value;
    const rf = document.getElementById('riskFilter').value;
    const pf = document.getElementById('pillarFilter').value;
    return scan.results.filter(r => {
      if (sf && r.status !== sf) return false;
      if (rf && r.risk !== rf) return false;
      if (pf && r.pillar !== pf) return false;
      if (!q) return true;
      return `${r.id} ${r.title} ${r.message || ''} ${r.category || ''}`.toLowerCase().includes(q);
    });
  }

  function sortRows(rows) {
    if (!sortKey) return rows;
    const val = r => {
      if (sortKey === 'risk') return RISK_SORT[r.risk] ?? 9;
      if (sortKey === 'status') return STATUS_SORT[r.status] ?? 9;
      if (sortKey === 'durationMs') return Number(r.durationMs) || 0;
      if (sortKey === 'id') return Number(r.id) || 0;
      return `${r[sortKey] || ''}`.toLowerCase();
    };
    return [...rows].sort((a, b) => {
      const av = val(a); const bv = val(b);
      if (av < bv) return -sortDir;
      if (av > bv) return sortDir;
      return 0;
    });
  }

  function renderTable(scan) {
    const rows = sortRows(filtered(scan));
    document.querySelector('#resultsTable tbody').innerHTML = rows.map(r => {
      const selected = `${selectedResultId}` === `${r.id}` ? ' selected' : '';
      const docs = r.documentationUrl
        ? `<a href="${esc(Toolkit.safeUrl(r.documentationUrl))}" target="_blank" rel="noopener noreferrer">Docs</a>`
        : '';
      return `<tr data-id="${esc(r.id)}" class="clickable-row${selected}"><td>${statusBadge(r.status)}</td><td>${esc(r.id)}</td><td>${esc(r.title)}</td><td>${esc(r.pillar || '')}</td><td>${esc(r.risk || '')}</td><td class="num nowrap">${esc(fmtMs(r.durationMs))}</td><td>${docs}</td></tr>`;
    }).join('');
  }

  function onRowClick(e) {
    if (e.target.closest('a')) return;
    const tr = e.target.closest('tr[data-id]');
    const scan = Store.getScan(currentScanId);
    if (!tr || !scan) return;
    const row = scan.results.find(x => `${x.id}` === tr.dataset.id);
    if (!row) return;
    selectedResultId = row.id;
    openModal(row);
    renderTable(scan);
  }

  function bindSorting() {
    document.querySelectorAll('#resultsTable th[data-sort]').forEach(th => {
      th.addEventListener('click', () => {
        const key = th.dataset.sort;
        if (sortKey === key) sortDir = -sortDir;
        else { sortKey = key; sortDir = 1; }
        document.querySelectorAll('#resultsTable th[data-sort]').forEach(x => x.classList.remove('sort-asc', 'sort-desc'));
        th.classList.add(sortDir > 0 ? 'sort-asc' : 'sort-desc');
        const scan = Store.getScan(currentScanId);
        if (scan) renderTable(scan);
      });
    });
  }

  function bindFilters() {
    ['resultSearch', 'statusFilter', 'riskFilter', 'pillarFilter'].forEach(id => {
      document.getElementById(id).addEventListener('input', () => {
        const scan = Store.getScan(currentScanId);
        if (scan) renderTable(scan);
      });
    });
  }

  const EXPORT_BUTTONS = ['exportOfficialBtn', 'exportHtmlBtn', 'exportXlsxBtn', 'exportCsvBtn', 'exportJsonBtn', 'exportWorkshopBtn'];
  function setExportsEnabled(on) {
    EXPORT_BUTTONS.forEach(id => { const b = document.getElementById(id); if (b) b.disabled = !on; });
  }

  function refreshPicker(preselect) {
    const scans = Store.loadScans().sort((a, b) => b.startedAt.localeCompare(a.startedAt));
    const picker = document.getElementById('scanPicker');
    if (!scans.length) {
      picker.innerHTML = '<option value="">(no scans)</option>';
      currentScanId = null;
      document.querySelector('#resultsTable tbody').innerHTML = '';
      document.getElementById('summaryGrid').innerHTML = '';
      document.getElementById('resultsCharts').innerHTML = '';
      document.getElementById('scanMeta').textContent = 'No scans stored in this browser yet. Run an assessment first.';
      setExportsEnabled(false);
      return;
    }

    setExportsEnabled(true);
    picker.innerHTML = scans.map(s => `<option value="${esc(s.id)}">${esc(new Date(s.startedAt).toLocaleString())} | ${esc(s.tenantDomain || s.tenantId || 'tenant')} | ${Number(s.summary?.failed) || 0}F/${Number(s.summary?.passed) || 0}P</option>`).join('');
    picker.value = preselect || scans[0].id;
    currentScanId = picker.value;
    show(currentScanId);
  }

  function show(scanId) {
    const scan = Store.getScan(scanId);
    if (!scan) return;
    currentScanId = scan.id;
    selectedResultId = null;
    setSummary(scan);
    renderTable(scan);
  }

  const EXPORT_COLUMNS = [
    { key: 'scanId', label: 'Scan' },
    { key: 'startedAt', label: 'Started' },
    { key: 'tenantId', label: 'Tenant id' },
    { key: 'tenantDomain', label: 'Tenant domain' },
    { key: 'id', label: 'Control' },
    { key: 'title', label: 'Title' },
    { key: 'status', label: 'Status' },
    { key: 'risk', label: 'Risk' },
    { key: 'pillar', label: 'Pillar' },
    { key: 'category', label: 'Category' },
    { key: 'implementationLevel', label: 'Implementation' },
    { key: 'durationMs', label: 'Duration (ms)' },
    { key: 'message', label: 'Summary', format: v => `${v || ''}`.replace(/\s*\n\s*/g, ' ') },
  ];
  function exportRows(scan) {
    return (scan.results || []).map(r => ({ ...r, scanId: scan.id, startedAt: scan.startedAt, tenantId: scan.tenantId || '', tenantDomain: scan.tenantDomain || '' }));
  }
  const fileBase = scan => `SimpleZTA-${`${scan.tenantDomain || scan.tenantId || 'tenant'}`.replace(/[^a-z0-9.-]/gi, '_')}-${`${scan.startedAt || ''}`.slice(0, 10)}`;

  function exportHandler(id, fn, label) {
    const btn = document.getElementById(id);
    btn.addEventListener('click', () => busy(btn, async () => {
      const scan = Store.getScan(currentScanId);
      if (!scan) return;
      try { await fn(scan); }
      catch (e) { Log.err(`${label} export failed:`, e); toast(`${label} export failed: ${e.message || e}`, 'err'); }
    }));
  }

  function init() {
    bindFilters();
    bindSorting();
    document.querySelector('#resultsTable tbody').addEventListener('click', onRowClick);
    document.getElementById('scanPicker').addEventListener('change', e => show(e.target.value));

    exportHandler('exportOfficialBtn', scan => ReportOfficial.export(scan), 'Official report');
    exportHandler('exportHtmlBtn', scan => ReportNative.export(scan), 'HTML report');
    exportHandler('exportWorkshopBtn', scan => ReportNative.exportWorkshop(scan), 'Workshop');
    exportHandler('exportXlsxBtn', scan => Export.xlsx([{ name: 'Results', rows: exportRows(scan), columns: EXPORT_COLUMNS }], `${fileBase(scan)}.xlsx`), 'Excel');
    exportHandler('exportCsvBtn', async scan => { Export.csv(exportRows(scan), { columns: EXPORT_COLUMNS, filename: `${fileBase(scan)}.csv` }); }, 'CSV');
    exportHandler('exportJsonBtn', async scan => { Export.save(new Blob([JSON.stringify(scan, null, 2)], { type: 'application/json' }), `${fileBase(scan)}.json`); }, 'JSON');
  }

  // Opens the detail dialog for a control in a scan (dashboard top fixes). Shows that
  // scan in the picker and table first.
  function openDetail(scanId, testId) {
    const scan = Store.getScan(scanId);
    if (!scan) return;
    if (currentScanId !== scanId) {
      const picker = document.getElementById('scanPicker');
      if (picker && Array.from(picker.options).some(o => o.value === scanId)) picker.value = scanId;
      show(scanId);
    }
    const row = scan.results.find(x => `${x.id}` === `${testId}`);
    if (!row) return;
    selectedResultId = row.id;
    openModal(row);
    renderTable(scan);
  }

  window.UIResults = {
    init,
    refreshPicker,
    show,
    openDetail,
    getCurrentScanId: () => currentScanId,
    exportColumns: EXPORT_COLUMNS,
    exportRows,
  };
})();
