(() => {
  let currentScanId = null;
  let selectedResultId = null;

  function modal() {
    return document.getElementById('resultModal');
  }

  function closeModal() {
    const el = modal();
    el.classList.add('hidden');
    el.setAttribute('aria-hidden', 'true');
  }

  function openModal(row) {
    const el = modal();
    document.getElementById('detailView').innerHTML = renderDetail(row);
    document.getElementById('resultModalTitle').textContent = `${row.id} ${row.title}`;
    el.classList.remove('hidden');
    el.setAttribute('aria-hidden', 'false');
  }

  function escapeHtml(value) {
    return `${value ?? ''}`
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function jsonBlock(value) {
    const text = typeof value === 'string' ? value : JSON.stringify(value, null, 2);
    return `<pre>${escapeHtml(text || '')}</pre>`;
  }

  function renderDetail(row) {
    const evidence = row.evidence == null ? '<div class="muted">No evidence captured.</div>' : jsonBlock(row.evidence);
    const trace = Array.isArray(row.apiTrace) && row.apiTrace.length
      ? row.apiTrace.map((entry, index) => `
        <div class="trace-entry">
          <div><strong>Call ${index + 1}</strong> ${escapeHtml(entry.method || 'GET')} ${escapeHtml(entry.status || '')}</div>
          <div class="trace-url">${escapeHtml(entry.url || '')}</div>
          ${entry.requestPreview ? `<div class="trace-block"><div class="muted">Request</div>${jsonBlock(entry.requestPreview)}</div>` : ''}
          <div class="trace-block"><div class="muted">Response</div>${jsonBlock(entry.responsePreview || '')}</div>
        </div>
      `).join('')
      : '<div class="muted">No API calls were recorded for this result.</div>';

    const impl = row.implementationLevel === 'partial' && row.implementationNotes
      ? `
      <div class="detail-section">
        <div><strong>Partial implementation</strong></div>
        <div><strong>Implemented:</strong> ${escapeHtml(row.implementationNotes.implemented || '-')}</div>
        <div><strong>Not implemented:</strong> ${escapeHtml(row.implementationNotes.notImplemented || '-')}</div>
      </div>`
      : '';

    const docs = row.documentationUrl
      ? `
      <div class="detail-section">
        <div><strong>Supporting documentation</strong></div>
        <div><a href="${escapeHtml(row.documentationUrl)}" target="_blank" rel="noopener noreferrer">Guidance root</a></div>
        <div><a href="${escapeHtml(row.documentationSearchUrl || '')}" target="_blank" rel="noopener noreferrer">Search by this control</a></div>
      </div>`
      : '';

    return `
      <div class="detail-section">
        <div><strong>${escapeHtml(row.id)}</strong> ${escapeHtml(row.title)}</div>
        <div class="muted">${escapeHtml(row.status)} | ${escapeHtml(row.pillar || '')} | ${escapeHtml(row.risk || '')} | ${fmtMs(row.durationMs)}</div>
      </div>
      <div class="detail-section">
        <div><strong>Summary</strong></div>
        <div>${escapeHtml(row.message || '')}</div>
      </div>
      ${impl}
      ${docs}
      <div class="detail-section">
        <div><strong>Evidence</strong></div>
        ${evidence}
      </div>
      <div class="detail-section">
        <div><strong>API Trace</strong></div>
        ${trace}
      </div>
    `;
  }

  function fmtMs(ms) {
    if (ms == null) return '';
    if (ms < 1000) return `${ms} ms`;
    return `${(ms / 1000).toFixed(1)} s`;
  }

  function kpi(label, value) {
    return `<div class="kpi"><div class="label">${label}</div><div class="value">${value}</div></div>`;
  }

  function setSummary(scan) {
    const s = scan.summary || { total: 0, passed: 0, failed: 0, skipped: 0, error: 0 };
    document.getElementById('scanMeta').textContent = `${scan.tenantDomain || scan.tenantId || 'Unknown tenant'} | ${new Date(scan.startedAt).toLocaleString()} | ${scan.results.length} controls`;
    document.getElementById('summaryGrid').innerHTML = [
      kpi('Total', s.total),
      kpi('Passed', s.passed),
      kpi('Failed', s.failed),
      kpi('Skipped', s.skipped),
      kpi('Error', s.error),
      kpi('Implemented', `${scan.implementationCoverage?.implemented || 0}/${scan.implementationCoverage?.totalCatalog || 0}`),
    ].join('');

    const counts = { Passed: 0, Failed: 0, Skipped: 0, Error: 0 };
    for (const r of (scan.results || [])) {
      if (counts[r.status] != null) counts[r.status] += 1;
    }
    const max = Math.max(1, counts.Passed, counts.Failed, counts.Skipped, counts.Error);
    document.getElementById('resultsCharts').innerHTML = `
      <div class="chart-card">
        <h3>Status distribution</h3>
        ${Object.entries(counts).map(([name, val]) => `
          <div class="bar-row">
            <span class="bar-label">${name}</span>
            <span class="bar-track"><span class="bar-fill ${name.toLowerCase()}" style="width:${Math.round((val / max) * 100)}%"></span></span>
            <span class="bar-value">${val}</span>
          </div>
        `).join('')}
      </div>
    `;
  }

  function filtered(scan) {
    const q = (document.getElementById('resultSearch').value || '').toLowerCase();
    const sf = document.getElementById('statusFilter').value;
    const rf = document.getElementById('riskFilter').value;
    return scan.results.filter(r => {
      if (sf && r.status !== sf) return false;
      if (rf && r.risk !== rf) return false;
      if (!q) return true;
      return `${r.id} ${r.title} ${r.message || ''} ${r.category || ''}`.toLowerCase().includes(q);
    });
  }

  function renderTable(scan) {
    const rows = filtered(scan);
    const tbody = document.querySelector('#resultsTable tbody');
    tbody.innerHTML = rows.map(r => {
      const selected = selectedResultId === r.id ? ' class="selected"' : '';
      const docs = r.documentationUrl
        ? `<a href="${r.documentationUrl}" target="_blank" rel="noopener noreferrer">Docs</a>`
        : '';
      return `<tr data-id="${r.id}"${selected}><td class="status-${r.status}">${r.status}</td><td>${r.id}</td><td>${r.title}</td><td>${r.pillar || ''}</td><td>${r.risk || ''}</td><td>${fmtMs(r.durationMs)}</td><td>${docs}</td></tr>`;
    }).join('');

    tbody.querySelectorAll('tr').forEach(tr => {
      tr.addEventListener('click', () => {
        selectedResultId = tr.dataset.id;
        const row = scan.results.find(x => x.id === selectedResultId);
        if (!row) return;
        openModal(row);
        renderTable(scan);
      });
    });
  }

  function bindFilters() {
    ['resultSearch', 'statusFilter', 'riskFilter'].forEach(id => {
      document.getElementById(id).addEventListener('input', () => {
        const scan = Store.getScan(currentScanId);
        if (scan) renderTable(scan);
      });
      document.getElementById(id).addEventListener('change', () => {
        const scan = Store.getScan(currentScanId);
        if (scan) renderTable(scan);
      });
    });
  }

  function refreshPicker(preselect) {
    const scans = Store.loadScans().sort((a, b) => b.startedAt.localeCompare(a.startedAt));
    const picker = document.getElementById('scanPicker');
    if (!scans.length) {
      picker.innerHTML = '<option value="">(no scans)</option>';
      document.querySelector('#resultsTable tbody').innerHTML = '';
      closeModal();
      document.getElementById('summaryGrid').innerHTML = '';
      document.getElementById('scanMeta').textContent = '';
      return;
    }

    picker.innerHTML = scans.map(s => `<option value="${s.id}">${new Date(s.startedAt).toLocaleString()} | ${(s.tenantDomain || s.tenantId || 'tenant')} | ${s.summary.failed}F/${s.summary.passed}P</option>`).join('');
    picker.value = preselect || scans[0].id;
    currentScanId = picker.value;
    show(currentScanId);
  }

  function show(scanId) {
    const scan = Store.getScan(scanId);
    if (!scan) return;
    currentScanId = scan.id;
    selectedResultId = null;
    closeModal();
    setSummary(scan);
    renderTable(scan);
  }

  function init() {
    bindFilters();
    closeModal();
    document.getElementById('scanPicker').addEventListener('change', e => show(e.target.value));
    document.getElementById('closeResultModalBtn').addEventListener('click', closeModal);
    document.getElementById('resultModalBackdrop').addEventListener('click', closeModal);
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape') closeModal();
    });

    document.getElementById('exportJsonBtn').addEventListener('click', () => {
      const scan = Store.getScan(currentScanId);
      if (!scan) return;
      const blob = new Blob([JSON.stringify(scan, null, 2)], { type: 'application/json' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `${scan.id}.json`;
      a.click();
      URL.revokeObjectURL(a.href);
    });

    document.getElementById('exportCsvBtn').addEventListener('click', () => {
      const scan = Store.getScan(currentScanId);
      if (!scan) return;
      const blob = new Blob([Store.toCsv(scan)], { type: 'text/csv' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `${scan.id}.csv`;
      a.click();
      URL.revokeObjectURL(a.href);
    });
  }

  window.UIResults = { init, refreshPicker, show };
})();
