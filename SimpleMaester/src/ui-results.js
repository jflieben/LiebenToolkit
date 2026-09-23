// Results UI: shows a single scan in detail with summary tiles and a sortable
// table. Clicking a row opens the detail drawer.
(() => {
  let _scan = null;
  let _filtered = [];
  let _sortKey = 'severity';
  let _sortAsc = false;

  // Severity sort weight (higher = worse).
  const sevW = { Critical: 5, High: 4, Medium: 3, Low: 2, Info: 1 };
  const statusW = { Failed: 4, Error: 3, Skipped: 2, Passed: 1 };

  async function show(scan) {
    _scan = scan;
    document.getElementById('noResultsCard').classList.add('hidden');
    document.getElementById('resultsWorkspace').classList.remove('hidden');

    const meta = document.getElementById('scanMeta');
    meta.textContent = `${scan.tenantDomain || scan.tenantId} - ${new Date(scan.startedAt).toLocaleString()} - ${scan.results.length} tests in ${(scan.durationMs/1000).toFixed(1)}s`;

    renderSummary(scan.summary);
    populateTags(scan.results);
    applyFilters();
  }

  function renderSummary(s) {
    const grid = document.getElementById('summaryGrid');
    grid.innerHTML = '';
    const tiles = [
      { lbl: 'Total', num: s.total },
      { lbl: 'Passed', num: s.passed, cls: 'passed' },
      { lbl: 'Failed', num: s.failed, cls: 'failed' },
      { lbl: 'Skipped', num: s.skipped, cls: 'skipped' },
      { lbl: 'Errors', num: s.error, cls: 'error' },
      { lbl: 'Pass rate', num: s.total ? Math.round(100 * s.passed / Math.max(1, s.passed + s.failed)) + '%' : '-' },
    ];
    for (const t of tiles) {
      const div = document.createElement('div');
      div.className = 'summary-tile' + (t.cls ? ' ' + t.cls : '');
      div.innerHTML = `<div class="num">${t.num}</div><div class="lbl">${t.lbl}</div>`;
      grid.appendChild(div);
    }
  }

  function populateTags(results) {
    const sel = document.getElementById('filterTag');
    const tags = Array.from(new Set(results.map(r => r.tag || r.category || 'Other'))).sort();
    sel.innerHTML = '<option value="">All categories</option>' + tags.map(t => `<option value="${t}">${t}</option>`).join('');
  }

  function applyFilters() {
    if (!_scan) return;
    const q = (document.getElementById('resultsSearch').value || '').toLowerCase();
    const sev = document.getElementById('filterSeverity').value;
    const st  = document.getElementById('filterStatus').value;
    const tag = document.getElementById('filterTag').value;
    _filtered = _scan.results.filter(r => {
      if (sev && r.severity !== sev) return false;
      if (st  && r.status   !== st)  return false;
      if (tag && (r.tag||r.category) !== tag) return false;
      if (q) {
        const blob = `${r.id} ${r.title||''} ${r.tag||''} ${r.reason||''}`.toLowerCase();
        if (!blob.includes(q)) return false;
      }
      return true;
    });
    sort();
    render();
  }

  function sort() {
    _filtered.sort((a, b) => {
      let va, vb;
      switch (_sortKey) {
        case 'severity': va = sevW[a.severity]||0; vb = sevW[b.severity]||0; break;
        case 'status':   va = statusW[a.status]||0; vb = statusW[b.status]||0; break;
        case 'durationMs': va = a.durationMs||0; vb = b.durationMs||0; break;
        default: va = (a[_sortKey]||'').toString().toLowerCase(); vb = (b[_sortKey]||'').toString().toLowerCase();
      }
      if (va < vb) return _sortAsc ? -1 : 1;
      if (va > vb) return _sortAsc ?  1 : -1;
      return 0;
    });
  }

  function render() {
    const tbody = document.querySelector('#resultsTable tbody');
    tbody.innerHTML = '';
    for (const r of _filtered) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><span class="badge ${r.status?.toLowerCase()}">${r.status||'-'}</span></td>
        <td><span class="badge ${(r.severity||'info').toLowerCase()}">${r.severity||'-'}</span></td>
        <td>${escape(r.id)}</td>
        <td>${escape(r.title||'')}</td>
        <td><span class="badge tag">${escape(r.tag||r.category||'-')}</span></td>
        <td>${r.durationMs||0}ms</td>`;
      tr.addEventListener('click', () => openDetail(r));
      tbody.appendChild(tr);
    }
  }

  function openDetail(r) {
    document.getElementById('detailTitle').textContent = `${r.id} - ${r.title || ''}`;
    const body = document.getElementById('detailBody');
    const md = window.marked && r.detailMd ? window.marked.parse(r.detailMd) : (r.description ? `<p>${escape(r.description)}</p>` : '');
    let history = '';
    History.getAllScans().then(all => {
      const tenantScans = all.filter(s => s.tenantId === _scan.tenantId).sort((a,b) => a.startedAt.localeCompare(b.startedAt));
      const points = tenantScans.map(s => {
        const tr = s.results.find(x => x.id === r.id);
        return { at: s.startedAt, status: tr?.status || '-' };
      });
      const histEl = document.getElementById('detail-history');
      if (histEl) histEl.innerHTML = renderHistorySpark(points);
    });
    body.innerHTML = `
      <div>
        <span class="badge ${(r.severity||'info').toLowerCase()}">${r.severity||'-'}</span>
        <span class="badge ${r.status?.toLowerCase()}">${r.status||'-'}</span>
        <span class="badge tag">${escape(r.tag||r.category||'-')}</span>
      </div>
      <h3>Result</h3>
      <p>${escape(r.reason || '-')}</p>
      ${r.actual !== undefined && r.actual !== null ? `<p><strong>Observed value:</strong> <code>${escape(JSON.stringify(r.actual))}</code></p>` : ''}
      ${r.expectation ? `<p><strong>Expected:</strong> ${escape(r.expectation)}</p>` : ''}
      ${r.relativeUri ? `<p><strong>Graph endpoint:</strong> <code>${escape(r.apiVersion || 'v1.0')}/${escape(r.relativeUri)}</code></p>` : ''}
      ${renderGraphCalls(r.graphCalls)}
      ${md ? `<h3>Description</h3>${md}` : ''}
      <h3>Trend across all your scans</h3>
      <div id="detail-history">Loading...</div>
      ${r.docUrl ? `<p><a href="${r.docUrl}" target="_blank" rel="noopener">View on maester.dev &rarr;</a></p>` : ''}`;
    document.getElementById('detailDrawer').classList.remove('hidden');
  }

  function renderGraphCalls(calls) {
    if (!Array.isArray(calls) || calls.length === 0) return '';
    const rows = calls.map((c, i) => {
      if (c.method === 'INFO') {
        return `<details class="gcall"><summary><span class="badge skipped">info</span> ${escape(c.url)}</summary></details>`;
      }
      const statusCls = c.status >= 200 && c.status < 300 ? 'passed' : c.status === 0 ? 'skipped' : 'failed';
      const cacheTag = c.fromCache ? ' <span class="badge tag">cached</span>' : '';
      const truncTag = c.truncated ? ' <span class="badge skipped">truncated</span>' : '';
      const body = c.response
        ? `<pre class="gcall-body">${escape(prettyJson(c.response))}</pre>`
        : '<p class="muted">(no response body)</p>';
      return `<details class="gcall"${i === 0 ? ' open' : ''}>
        <summary>
          <span class="badge ${statusCls}">${c.status || '?'}</span>
          <code>${escape(c.method)} ${escape(shortenUrl(c.url))}</code>
          <span class="muted">${c.durationMs|0}ms</span>${cacheTag}${truncTag}
        </summary>
        ${body}
      </details>`;
    }).join('');
    return `<h3>Graph calls (${calls.filter(c => c.method !== 'INFO').length})</h3>${rows}`;
  }

  function prettyJson(s) {
    if (!s) return '';
    try { return JSON.stringify(JSON.parse(s), null, 2); }
    catch { return s; }
  }
  function shortenUrl(u) {
    if (!u) return '';
    return u.replace(/^https:\/\/graph\.microsoft\.com\//, '');
  }

  function renderHistorySpark(points) {
    if (!points.length) return '<p class="muted">No history yet.</p>';
    const bars = points.map(p => `<span class="${p.status.toLowerCase()}" style="height:${p.status === 'Passed' ? 14 : p.status === 'Failed' ? 14 : 8}px" title="${p.at}: ${p.status}"></span>`).join('');
    return `<div class="spark">${bars}</div><p class="muted" style="margin-top:8px">${points.length} historical run(s).</p>`;
  }

  function escape(s) { return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

  function exportJson() {
    if (!_scan) return;
    download(`SimpleMaester-${_scan.id}.json`, JSON.stringify(_scan, null, 2), 'application/json');
  }
  function exportCsv() {
    if (!_scan) return;
    const cols = ['id','title','category','severity','status','reason','actual','expectation','durationMs'];
    const lines = [cols.join(',')];
    for (const r of _scan.results) {
      lines.push(cols.map(c => csv(r[c])).join(','));
    }
    download(`SimpleMaester-${_scan.id}.csv`, lines.join('\n'), 'text/csv');
  }
  function csv(v) {
    if (v == null) return '';
    const s = typeof v === 'object' ? JSON.stringify(v) : String(v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  function download(name, data, type) {
    const blob = new Blob([data], { type });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 200);
  }

  function init() {
    document.getElementById('resultsSearch').addEventListener('input', applyFilters);
    document.getElementById('filterSeverity').addEventListener('change', applyFilters);
    document.getElementById('filterStatus').addEventListener('change', applyFilters);
    document.getElementById('filterTag').addEventListener('change', applyFilters);
    document.querySelectorAll('#resultsTable thead th[data-sort]').forEach(th => {
      th.addEventListener('click', () => {
        const k = th.dataset.sort;
        if (_sortKey === k) _sortAsc = !_sortAsc; else { _sortKey = k; _sortAsc = false; }
        sort(); render();
      });
    });
    document.getElementById('exportJsonBtn').addEventListener('click', exportJson);
    document.getElementById('exportCsvBtn').addEventListener('click', exportCsv);
    document.getElementById('closeDetail').addEventListener('click', () => document.getElementById('detailDrawer').classList.add('hidden'));
    document.querySelector('.drawer-backdrop').addEventListener('click', () => document.getElementById('detailDrawer').classList.add('hidden'));
  }

  window.UIResults = { show, init, openDetail };
})();
