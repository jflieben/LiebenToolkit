// Results UI: one scan with summary tiles and a sortable table. A row opens the detail drawer.
(() => {
  const { esc, statusBadge, severityBadge, tagBadge, badge, sparkHtml } = SMUi;
  let _scan = null;
  let _filtered = [];
  let _sortKey = 'severity';
  let _sortAsc = false;

  const sevW = { Critical: 5, High: 4, Medium: 3, Low: 2, Info: 1 };
  const statusW = { Failed: 4, Error: 3, Skipped: 2, Passed: 1 };
  const $ = (id) => document.getElementById(id);

  function show(scan) {
    _scan = scan;
    $('noResultsCard').hidden = true;
    $('resultsWorkspace').hidden = false;
    $('scanMeta').textContent = `${scan.tenantDomain || scan.tenantId} - ${new Date(scan.startedAt).toLocaleString()} - ${scan.results.length} tests in ${(scan.durationMs / 1000).toFixed(1)}s`;
    renderSummary(scan.summary);
    populateTags(scan.results);
    applyFilters();
  }

  function renderSummary(s) {
    const rate = s.total ? `${Math.round(100 * s.passed / Math.max(1, s.passed + s.failed))}%` : '-';
    const tiles = [
      { lbl: 'Total', num: s.total, cls: 'accent-primary' },
      { lbl: 'Passed', num: s.passed, cls: 'accent-ok' },
      { lbl: 'Failed', num: s.failed, cls: 'accent-danger' },
      { lbl: 'Skipped', num: s.skipped, cls: '' },
      { lbl: 'Errors', num: s.error, cls: 'accent-warn' },
      { lbl: 'Pass rate', num: rate, cls: 'accent-ok', foot: 'of passed and failed' },
    ];
    $('summaryGrid').replaceChildren(...tiles.map((t) => Toolkit.h('div.stat', { class: t.cls }, [
      Toolkit.h('span.label', { text: t.lbl }),
      Toolkit.h('span.value', { text: String(t.num) }),
      t.foot ? Toolkit.h('span.foot', { text: t.foot }) : null,
    ])));
  }

  function populateTags(results) {
    const sel = $('filterTag');
    const tags = Array.from(new Set(results.map((r) => r.tag || r.category || 'Other'))).sort();
    sel.replaceChildren(new Option('All categories', ''), ...tags.map((t) => new Option(t, t)));
  }

  function applyFilters() {
    if (!_scan) return;
    const q = ($('resultsSearch').value || '').toLowerCase();
    const sev = $('filterSeverity').value;
    const st = $('filterStatus').value;
    const tag = $('filterTag').value;
    _filtered = _scan.results.filter((r) => {
      if (sev && r.severity !== sev) return false;
      if (st && r.status !== st) return false;
      if (tag && (r.tag || r.category) !== tag) return false;
      if (q && !`${r.id} ${r.title || ''} ${r.tag || ''} ${r.reason || ''}`.toLowerCase().includes(q)) return false;
      return true;
    });
    sort();
    render();
  }

  function sort() {
    _filtered.sort((a, b) => {
      let va, vb;
      switch (_sortKey) {
        case 'severity': va = sevW[a.severity] || 0; vb = sevW[b.severity] || 0; break;
        case 'status': va = statusW[a.status] || 0; vb = statusW[b.status] || 0; break;
        case 'durationMs': va = a.durationMs || 0; vb = b.durationMs || 0; break;
        default: va = String(a[_sortKey] || '').toLowerCase(); vb = String(b[_sortKey] || '').toLowerCase();
      }
      if (va < vb) return _sortAsc ? -1 : 1;
      if (va > vb) return _sortAsc ? 1 : -1;
      return 0;
    });
    document.querySelectorAll('#resultsTable thead th[data-sort]').forEach((th) => {
      th.classList.toggle('sort-asc', th.dataset.sort === _sortKey && _sortAsc);
      th.classList.toggle('sort-desc', th.dataset.sort === _sortKey && !_sortAsc);
    });
  }

  function render() {
    const tbody = document.querySelector('#resultsTable tbody');
    tbody.replaceChildren();
    $('resultsCount').textContent = `${_filtered.length} of ${_scan.results.length} shown`;
    for (const r of _filtered) {
      const tr = document.createElement('tr');
      tr.className = 'clickable-row';
      tr.tabIndex = 0;
      tr.innerHTML = `
        <td>${statusBadge(r.status)}</td>
        <td>${severityBadge(r.severity)}</td>
        <td class="mono nowrap">${esc(r.id)}</td>
        <td>${esc(r.title || '')}</td>
        <td>${tagBadge(r.tag || r.category)}</td>
        <td class="num nowrap">${Number(r.durationMs) || 0} ms</td>`;
      tr.addEventListener('click', () => openDetail(r));
      tr.addEventListener('keydown', (e) => { if (e.key === 'Enter') openDetail(r); });
      tbody.appendChild(tr);
    }
  }

  function closeDetail() { $('detailDrawer').hidden = true; }

  // r can come from a stored scan or from the history table. Description and markdown
  // always come from the bundled catalog, never from stored data.
  async function openDetail(r) {
    const entry = await Registry.getById(r.id).catch(() => null);
    const detailMd = entry && entry.detailMd;
    const description = (entry && entry.description) || r.description;
    const docUrl = r.docUrl || (entry && entry.docUrl);
    $('detailTitle').textContent = `${r.id} - ${r.title || (entry && entry.title) || ''}`;
    const md = detailMd ? Markdown.render(detailMd) : (description ? `<p>${esc(description)}</p>` : '');
    const body = $('detailBody');
    body.innerHTML = `
      <div class="row tight">${severityBadge(r.severity)} ${statusBadge(r.status)} ${tagBadge(r.tag || r.category)}</div>
      <h3>Result</h3>
      <p>${esc(r.reason || '-')}</p>
      ${r.actual !== undefined && r.actual !== null ? `<p><strong>Observed value:</strong> <code>${esc(JSON.stringify(r.actual))}</code></p>` : ''}
      ${r.expectation ? `<p><strong>Expected:</strong> ${esc(r.expectation)}</p>` : ''}
      ${r.relativeUri ? `<p><strong>Graph endpoint:</strong> <code>${esc(r.apiVersion || 'v1.0')}/${esc(r.relativeUri)}</code></p>` : ''}
      ${renderCalls(r.graphCalls)}
      ${md ? `<h3>Description</h3><div class="md">${md}</div>` : ''}
      <h3>Trend across your scans of this tenant</h3>
      <div id="detailHistory" class="muted">Loading</div>
      ${docUrl ? `<p><a href="${esc(Toolkit.safeUrl(docUrl))}" target="_blank" rel="noopener noreferrer">View on maester.dev ${Toolkit.iconHtml('i-external', 'sm')}</a></p>` : ''}`;
    $('detailDrawer').hidden = false;
    $('closeDetail').focus();

    const tenantId = r.tenantId || (_scan && _scan.tenantId);
    History.getAllScans().then((all) => {
      const points = all.filter((s) => !tenantId || s.tenantId === tenantId)
        .sort((a, b) => a.startedAt.localeCompare(b.startedAt))
        .map((s) => { const x = s.results.find((y) => y.id === r.id); return x ? { at: s.startedAt, status: x.status } : null; })
        .filter(Boolean);
      const el = $('detailHistory');
      if (!el) return;
      el.innerHTML = points.length ? `${sparkHtml(points)}<p class="muted small">${points.length} run(s) of this test.</p>` : '<p class="muted">No history yet.</p>';
    }).catch(() => {});
  }

  function renderCalls(calls) {
    if (!Array.isArray(calls) || !calls.length) return '';
    const rows = calls.map((c, i) => {
      if (c.method === 'INFO') return `<p class="muted small">${esc(c.url)}</p>`;
      const cls = c.status >= 200 && c.status < 300 ? 'ok' : c.status === 0 ? 'neutral' : 'danger';
      const body = c.response ? `<pre class="gcall-body">${esc(prettyJson(c.response))}</pre>` : '<p class="muted small">No response body.</p>';
      return `<details class="gcall"${i === 0 ? ' open' : ''}>
        <summary>${badge(cls, c.status || '?')} <code>${esc(c.method)} ${esc(shortenUrl(c.url))}</code>
          <span class="muted small">${Number(c.durationMs) | 0} ms</span>${c.fromCache ? ` ${badge('info', 'cached')}` : ''}${c.truncated ? ` ${badge('neutral', 'truncated')}` : ''}
        </summary>
        ${body}
      </details>`;
    }).join('');
    return `<h3>API calls (${calls.filter((c) => c.method !== 'INFO').length})</h3>${rows}`;
  }

  function prettyJson(s) {
    try { return JSON.stringify(JSON.parse(s), null, 2); } catch { return s; }
  }
  function shortenUrl(u) { return String(u || '').replace(/^https:\/\/graph\.microsoft\.com\//, ''); }

  function exportJson() {
    if (!_scan) return;
    Export.save(new Blob([JSON.stringify(_scan, null, 2)], { type: 'application/json' }), `SimpleMaester-${_scan.id}.json`);
  }
  function exportCsv() {
    if (!_scan) return;
    Export.csv(_scan.results, {
      filename: `SimpleMaester-${_scan.id}.csv`,
      columns: [
        { key: 'id', label: 'Id' },
        { key: 'title', label: 'Title' },
        { key: 'category', label: 'Category', format: (v, r) => v || r.tag || '' },
        { key: 'severity', label: 'Severity' },
        { key: 'status', label: 'Status' },
        { key: 'reason', label: 'Reason' },
        { key: 'actual', label: 'Observed value' },
        { key: 'expectation', label: 'Expected' },
        { key: 'durationMs', label: 'Duration (ms)' },
      ],
    });
  }

  function init() {
    $('resultsSearch').addEventListener('input', applyFilters);
    $('filterSeverity').addEventListener('change', applyFilters);
    $('filterStatus').addEventListener('change', applyFilters);
    $('filterTag').addEventListener('change', applyFilters);
    document.querySelectorAll('#resultsTable thead th[data-sort]').forEach((th) => {
      th.addEventListener('click', () => {
        const k = th.dataset.sort;
        if (_sortKey === k) _sortAsc = !_sortAsc; else { _sortKey = k; _sortAsc = false; }
        sort(); render();
      });
    });
    $('exportJsonBtn').addEventListener('click', exportJson);
    $('exportCsvBtn').addEventListener('click', exportCsv);
    $('closeDetail').addEventListener('click', closeDetail);
    $('detailBackdrop').addEventListener('click', closeDetail);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('detailDrawer').hidden) closeDetail(); });
  }

  function hide() {
    _scan = null;
    $('noResultsCard').hidden = false;
    $('resultsWorkspace').hidden = true;
  }

  window.UIResults = { show, hide, init, openDetail };
})();
