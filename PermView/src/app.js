// PermView - SPA wiring.
// Tabs are workload providers from providers.js. Each tab renders a small
// form, runs the provider, then shows the results in a uniform table.
(() => {

  const state = {
    activeProviderId: null,
    lastResult: null,
    lastFormValues: null,
  };

  // --------- THEME ---------
  function initTheme() {
    const saved = localStorage.getItem('permview-theme') || 'light';
    document.documentElement.setAttribute('data-theme', saved);
  }
  function toggleTheme() {
    const cur = document.documentElement.getAttribute('data-theme');
    const nxt = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', nxt);
    localStorage.setItem('permview-theme', nxt);
  }

  // --------- TOAST ---------
  function toast(msg, kind = 'info') {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.className = `toast toast-${kind}`;
    el.classList.remove('hidden');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.add('hidden'), 4000);
  }

  // --------- AUTH UI ---------
  async function refreshAuthUI() {
    const acc = Auth.getAccount();
    const userBox = document.getElementById('userBox');
    const userName = document.getElementById('userName');
    const connectCard = document.getElementById('connectCard');
    const workspace = document.getElementById('workspace');
    if (acc) {
      userBox.classList.remove('hidden');
      userName.textContent = acc.username || acc.name || 'Signed in';
      connectCard.classList.add('hidden');
      workspace.classList.remove('hidden');
      if (!state.activeProviderId) selectProvider(Providers.list[0].id);
    } else {
      userBox.classList.add('hidden');
      connectCard.classList.remove('hidden');
      workspace.classList.add('hidden');
    }
  }
  async function onSignIn() {
    try {
      await Auth.signIn();
      await refreshAuthUI();
    } catch (e) { Log.err('Sign-in failed:', e); toast('Sign-in failed: ' + e.message, 'err'); }
  }
  async function onSignOut() {
    await Auth.signOut();
    state.activeProviderId = null;
    state.lastResult = null;
    state.lastFormValues = null;
    await refreshAuthUI();
  }

  // --------- PROVIDER TABS ---------

  function renderProviderTabs() {
    const tabsEl = document.getElementById('providerTabs');
    tabsEl.innerHTML = '';
    for (const p of Providers.list) {
      const btn = document.createElement('button');
      btn.className = 'tab' + (p.id === state.activeProviderId ? ' active' : '');
      btn.dataset.provider = p.id;
      btn.innerHTML = `<span class="tab-icon">${p.icon}</span><span>${p.name}</span>`;
      btn.addEventListener('click', () => selectProvider(p.id));
      tabsEl.appendChild(btn);
    }
  }

  async function selectProvider(id) {
    state.activeProviderId = id;
    state.lastResult = null;
    renderProviderTabs();
    renderProviderPanel();
  }

  function renderProviderPanel() {
    const provider = Providers.byId(state.activeProviderId);
    const root = document.getElementById('providerPanel');
    root.innerHTML = '';
    if (!provider) return;

    const card = document.createElement('div');
    card.className = 'card';

    const title = document.createElement('h2');
    title.textContent = provider.name;
    card.appendChild(title);

    const desc = document.createElement('p');
    desc.textContent = provider.description;
    card.appendChild(desc);

    if (provider.notice) {
      const notice = document.createElement('div');
      notice.className = 'notice';
      notice.innerHTML = `<strong>Quick view limitation:</strong> ${provider.notice}`;
      card.appendChild(notice);
    }

    // Form
    const form = document.createElement('form');
    form.className = 'provider-form';
    const inputEls = {};
    for (const f of provider.form) {
      const wrap = document.createElement('label');
      wrap.className = 'field';
      const lbl = document.createElement('span');
      lbl.textContent = f.label;
      wrap.appendChild(lbl);
      let input;
      if (f.kind === 'select') {
        input = document.createElement('select');
        input.innerHTML = '<option value="">Loading...</option>';
        input.disabled = true;
        // Fire and forget the loader.
        (async () => {
          try {
            const opts = await f.loadOptions();
            input.innerHTML = '<option value="">-- choose --</option>' +
              opts.map(o => `<option value="${escapeHtml(o.value)}">${escapeHtml(o.label)}</option>`).join('');
            input.disabled = false;
          } catch (e) {
            Log.err('loadOptions failed:', e);
            input.innerHTML = `<option value="">(failed: ${escapeHtml(e.message)})</option>`;
          }
        })();
      } else {
        input = document.createElement('input');
        input.type = 'text';
        if (f.placeholder) input.placeholder = f.placeholder;
      }
      input.name = f.name;
      wrap.appendChild(input);
      form.appendChild(wrap);
      inputEls[f.name] = input;
    }

    const actions = document.createElement('div');
    actions.className = 'row';
    const runBtn = document.createElement('button');
    runBtn.type = 'submit';
    runBtn.className = 'primary-btn';
    runBtn.textContent = 'Show permissions';
    actions.appendChild(runBtn);

    const status = document.createElement('span');
    status.className = 'muted run-status';
    actions.appendChild(status);
    form.appendChild(actions);

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      const values = {};
      for (const k of Object.keys(inputEls)) values[k] = inputEls[k].value;
      state.lastFormValues = values;
      runBtn.disabled = true;
      status.textContent = 'Starting...';

      // Streaming UI: progress bar + live-updating table.
      const stream = startStreamingResults(provider);
      try {
        const ctx = {
          onRow: (r) => stream.addRow(r),
          onRows: (rs) => rs.forEach(stream.addRow),
          onProgress: (p) => {
            // p = { label, current?, total? }
            stream.setProgress(p);
            if (p && p.label) status.textContent = p.label;
          },
          setEntity: (entity) => stream.setEntity(entity),
        };
        const res = await provider.run(values, ctx);
        // Provider may also return final shape; merge any extra rows.
        if (res) {
          if (res.entity) stream.setEntity(res.entity);
          if (Array.isArray(res.rows)) {
            for (const r of res.rows) stream.addRowIfMissing(r);
          }
        }
        stream.finish();
        state.lastResult = stream.snapshot();
        status.textContent = `${state.lastResult.rows.length} row(s)`;
      } catch (err) {
        Log.err('Provider run failed:', err);
        toast('Failed: ' + err.message, 'err');
        status.textContent = 'Failed';
        stream.fail(err.message);
      } finally {
        runBtn.disabled = false;
      }
    });

    card.appendChild(form);
    root.appendChild(card);

    // Results placeholder
    const resultsCard = document.createElement('div');
    resultsCard.className = 'card hidden';
    resultsCard.id = 'resultsCard';
    root.appendChild(resultsCard);
  }

  // --------- RESULTS (streaming) ---------

  // Builds the progress card + live-updating table, returns helpers to push
  // rows / progress as the provider runs.
  function startStreamingResults(provider) {
    const card = document.getElementById('resultsCard');
    card.classList.remove('hidden');
    card.innerHTML = '';

    const entity = { name: provider.name, sub: '' };
    const columns = Array.isArray(provider.columns) && provider.columns.length
      ? provider.columns
      : [
          { key: 'principal', label: 'Principal' },
          { key: 'principalType', label: 'Type' },
          { key: 'role', label: 'Role / Permission' },
          { key: 'scope', label: 'Scope' },
          { key: 'source', label: 'Source' },
        ];

    // Header + actions
    const head = document.createElement('div');
    head.className = 'row between';
    const titleWrap = document.createElement('div');
    const title = document.createElement('h2');
    title.textContent = entity.name;
    title.style.marginBottom = '4px';
    titleWrap.appendChild(title);
    const sub = document.createElement('div');
    sub.className = 'muted';
    titleWrap.appendChild(sub);
    head.appendChild(titleWrap);

    const exportBtn = document.createElement('button');
    exportBtn.className = 'ghost-btn';
    exportBtn.textContent = 'Download CSV';
    exportBtn.disabled = true;
    head.appendChild(exportBtn);
    card.appendChild(head);

    // Plug
    const plug = document.createElement('div');
    plug.className = 'plug';
    plug.innerHTML = 'This is a quick view. For comprehensive permission audits across every workload, every scope and every point in time, see <a href="https://www.m365permissions.com" target="_blank" rel="noopener">m365permissions.com</a>.';
    card.appendChild(plug);

    // Progress block
    const progressBox = document.createElement('div');
    progressBox.className = 'progress-box';
    progressBox.innerHTML = `
      <div class="progress-header">
        <span class="progress-label">Working...</span>
        <span class="progress-count"></span>
      </div>
      <div class="progress-bar"><div class="progress-fill"></div></div>`;
    card.appendChild(progressBox);
    const progressLabel = progressBox.querySelector('.progress-label');
    const progressCount = progressBox.querySelector('.progress-count');
    const progressFill = progressBox.querySelector('.progress-fill');
    progressFill.style.width = '0%';
    // Indeterminate state until we get a total.
    progressBox.classList.add('indeterminate');

    // Filter row (live, applies to current rows)
    const filterRow = document.createElement('div');
    filterRow.className = 'row';
    const filter = document.createElement('input');
    filter.type = 'search';
    filter.placeholder = 'Filter rows...';
    filterRow.appendChild(filter);
    const liveCount = document.createElement('span');
    liveCount.className = 'muted';
    filterRow.appendChild(liveCount);
    card.appendChild(filterRow);

    let activeQuickFilter = null;
    if (Array.isArray(provider.quickFilters) && provider.quickFilters.length) {
      const quickRow = document.createElement('div');
      quickRow.className = 'row quick-filter-row';

      const allBtn = document.createElement('button');
      allBtn.type = 'button';
      allBtn.className = 'chip-btn active';
      allBtn.textContent = 'All';
      quickRow.appendChild(allBtn);

      const quickBtns = [{ id: 'all', btn: allBtn, predicate: null }];
      for (const q of provider.quickFilters) {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'chip-btn';
        b.textContent = q.label;
        quickBtns.push({ id: q.id, btn: b, predicate: q.predicate });
        quickRow.appendChild(b);
      }

      function applyQuick(id, predicate) {
        activeQuickFilter = predicate;
        for (const q of quickBtns) q.btn.classList.toggle('active', q.id === id);
        repaint();
      }

      for (const q of quickBtns) {
        q.btn.addEventListener('click', () => applyQuick(q.id, q.predicate));
      }
      card.appendChild(quickRow);
    }

    // Table
    const wrap = document.createElement('div');
    wrap.className = 'table-wrap';
    const table = document.createElement('table');
    table.className = 'data-table';
    table.innerHTML = `<thead><tr>${columns.map(c => {
      const key = escapeHtml(c.key);
      const label = escapeHtml(c.label);
      return `<th>
        <button class="th-btn" data-col-key="${key}" type="button" title="Sort and filter">
          <span>${label}</span><span class="th-caret">▾</span>
        </button>
        <div class="th-menu hidden" data-col-menu="${key}"></div>
      </th>`;
    }).join('')}</tr></thead><tbody></tbody>`;
    const tbody = table.querySelector('tbody');
    wrap.appendChild(table);
    card.appendChild(wrap);

    // Empty state placeholder (only shown after finish())
    const emptyEl = document.createElement('p');
    emptyEl.className = 'muted hidden';
    emptyEl.textContent = 'No permissions found.';
    card.appendChild(emptyEl);

    const rows = [];
    const seen = new Set();
    const columnFilters = new Map(); // key -> Set(values) | null (all)
    let sortState = { key: null, dir: null }; // dir: 'asc' | 'desc' | null
    let activeMenu = null;
    const menuByKey = new Map();

    function closeActiveMenu() {
      if (!activeMenu) return;
      activeMenu.classList.add('hidden');
      activeMenu.style.visibility = '';
      activeMenu.style.left = '';
      activeMenu.style.top = '';
      activeMenu = null;
    }

    function positionMenuForButton(btn, menu) {
      if (!btn || !menu) return;

      menu.classList.remove('hidden');
      menu.style.visibility = 'hidden';
      menu.style.left = '0px';
      menu.style.top = '0px';

      const btnRect = btn.getBoundingClientRect();
      const menuRect = menu.getBoundingClientRect();
      const pad = 8;
      const gap = 6;

      let left = btnRect.left;
      if (left + menuRect.width > window.innerWidth - pad) {
        left = window.innerWidth - menuRect.width - pad;
      }
      left = Math.max(pad, left);

      let top = btnRect.bottom + gap;
      if (top + menuRect.height > window.innerHeight - pad) {
        top = btnRect.top - menuRect.height - gap;
      }
      top = Math.max(pad, top);

      menu.style.left = `${Math.round(left)}px`;
      menu.style.top = `${Math.round(top)}px`;
      menu.style.visibility = 'visible';
    }

    function clearAllColumnFilters() {
      for (const c of columns) columnFilters.set(c.key, null);
    }

    clearAllColumnFilters();

    for (const c of columns) {
      const menu = table.querySelector(`.th-menu[data-col-menu="${c.key}"]`);
      if (menu) menuByKey.set(c.key, menu);
    }

    document.addEventListener('click', (ev) => {
      if (!activeMenu) return;
      if (activeMenu.contains(ev.target)) return;
      if (ev.target && ev.target.closest && ev.target.closest('.th-btn')) return;
      closeActiveMenu();
    });

    document.addEventListener('keydown', (ev) => {
      if (ev.key === 'Escape') closeActiveMenu();
    });

    table.querySelectorAll('.th-btn').forEach((btn) => {
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        const key = btn.dataset.colKey;
        const menu = menuByKey.get(key);
        if (!menu) return;
        const wasOpen = !menu.classList.contains('hidden');
        closeActiveMenu();
        if (wasOpen) return;
        menu._anchorBtn = btn;
        renderColumnMenu(key, menu);
        positionMenuForButton(btn, menu);
        activeMenu = menu;
      });
    });

    window.addEventListener('resize', () => {
      if (activeMenu && activeMenu._anchorBtn) positionMenuForButton(activeMenu._anchorBtn, activeMenu);
    });

    window.addEventListener('scroll', () => {
      if (activeMenu && activeMenu._anchorBtn) positionMenuForButton(activeMenu._anchorBtn, activeMenu);
    }, true);

    function rowKey(r) {
      return columns.map(c => String(r[c.key] || '')).join('|');
    }

    function rowMatchesFilter(r, ft) {
      if (activeQuickFilter && !activeQuickFilter(r)) return false;
      for (const c of columns) {
        const selected = columnFilters.get(c.key);
        if (selected && selected.size) {
          const val = String(r[c.key] == null ? '' : r[c.key]);
          if (!selected.has(val)) return false;
        }
      }
      if (!ft) return true;
      const text = columns.map(c => String(r[c.key] || '')).join(' ').toLowerCase();
      return text.includes(ft);
    }

    function compareValues(a, b) {
      const na = Number(a);
      const nb = Number(b);
      const aNum = Number.isFinite(na);
      const bNum = Number.isFinite(nb);
      if (aNum && bNum) return na - nb;
      return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
    }

    function getVisibleRows() {
      const ft = filter.value.trim().toLowerCase();
      const vis = rows.filter(r => rowMatchesFilter(r, ft));
      if (sortState.key && sortState.dir) {
        vis.sort((ra, rb) => {
          const av = String(ra[sortState.key] == null ? '' : ra[sortState.key]);
          const bv = String(rb[sortState.key] == null ? '' : rb[sortState.key]);
          const cmp = compareValues(av, bv);
          return sortState.dir === 'asc' ? cmp : -cmp;
        });
      }
      return vis;
    }

    function updateHeaderState() {
      table.querySelectorAll('.th-btn').forEach((btn) => {
        const key = btn.dataset.colKey;
        const hasFilter = !!(columnFilters.get(key) && columnFilters.get(key).size);
        const isSort = sortState.key === key && sortState.dir;
        btn.classList.toggle('filtered', hasFilter);
        btn.classList.toggle('sorted', !!isSort);
        const caret = btn.querySelector('.th-caret');
        if (!caret) return;
        if (isSort && sortState.dir === 'asc') caret.textContent = '↑';
        else if (isSort && sortState.dir === 'desc') caret.textContent = '↓';
        else caret.textContent = '▾';
      });
    }

    function renderColumnMenu(colKey, menu) {
      const values = Array.from(new Set(rows.map(r => String(r[colKey] == null ? '' : r[colKey]))))
        .sort((a, b) => compareValues(a, b));
      const selected = columnFilters.get(colKey);
      const allSelected = !selected;

      menu.innerHTML = `
        <div class="th-menu-actions">
          <button type="button" class="th-menu-btn" data-action="sort-asc">Sort A -> Z</button>
          <button type="button" class="th-menu-btn" data-action="sort-desc">Sort Z -> A</button>
          <button type="button" class="th-menu-btn" data-action="clear-sort">Clear sort</button>
        </div>
        <div class="th-menu-actions">
          <button type="button" class="th-menu-btn" data-action="select-all">Select all</button>
          <button type="button" class="th-menu-btn" data-action="clear-filter">Clear filter</button>
        </div>
        <input type="search" class="th-menu-search" placeholder="Find value..." />
        <div class="th-menu-values"></div>`;

      const valuesWrap = menu.querySelector('.th-menu-values');
      const search = menu.querySelector('.th-menu-search');

      function repaintValues() {
        const q = (search.value || '').trim().toLowerCase();
        const html = values
          .filter(v => !q || v.toLowerCase().includes(q))
          .map(v => {
            const checked = allSelected || (selected && selected.has(v));
            const txt = v || '(blank)';
            return `<label class="th-menu-check"><input type="checkbox" data-value="${escapeHtml(v)}" ${checked ? 'checked' : ''}/> <span>${escapeHtml(txt)}</span></label>`;
          })
          .join('');
        valuesWrap.innerHTML = html || '<div class="muted">No values</div>';

        valuesWrap.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
          cb.addEventListener('change', () => {
            const cur = columnFilters.get(colKey);
            let next;
            if (!cur) next = new Set(values);
            else next = new Set(cur);
            const val = cb.dataset.value || '';
            if (cb.checked) next.add(val);
            else next.delete(val);
            if (next.size >= values.length) columnFilters.set(colKey, null);
            else columnFilters.set(colKey, next);
            repaint();
            updateHeaderState();
          });
        });
      }

      repaintValues();
      search.addEventListener('input', repaintValues);

      menu.querySelectorAll('.th-menu-btn').forEach((b) => {
        b.addEventListener('click', () => {
          const a = b.dataset.action;
          if (a === 'sort-asc') sortState = { key: colKey, dir: 'asc' };
          else if (a === 'sort-desc') sortState = { key: colKey, dir: 'desc' };
          else if (a === 'clear-sort' && sortState.key === colKey) sortState = { key: null, dir: null };
          else if (a === 'select-all') columnFilters.set(colKey, null);
          else if (a === 'clear-filter') columnFilters.set(colKey, new Set());
          repaint();
          updateHeaderState();
          renderColumnMenu(colKey, menu);
        });
      });

      if (menu._anchorBtn) positionMenuForButton(menu._anchorBtn, menu);
    }

    function badgeClassForCell(key, value, r) {
      if (key === 'riskLevel') {
        const v = String(value || '').toLowerCase();
        if (v === 'high') return 'badge-risk-high';
        if (v === 'medium') return 'badge-risk-medium';
        if (v === 'low') return 'badge-risk-low';
        return 'badge-risk-info';
      }
      if (key === 'credentialState') {
        const s = (r.credentialSeverity || '').toLowerCase();
        if (s === 'critical') return 'badge-risk-high';
        if (s === 'high') return 'badge-risk-medium';
        if (s === 'ok') return 'badge-risk-low';
        return 'badge-risk-info';
      }
      if (key === 'permissionMode') {
        const m = String(value || '').toLowerCase();
        if (m === 'application') return 'badge-mode-app';
        if (m === 'delegated') return 'badge-mode-delegated';
      }
      return '';
    }

    function appendDom(r) {
      const ft = filter.value.trim().toLowerCase();
      if (!rowMatchesFilter(r, ft)) return;
      const tr = document.createElement('tr');
      tr.innerHTML = columns.map(c => {
        const val = r[c.key] == null ? '' : String(r[c.key]);
        const badgeClass = badgeClassForCell(c.key, val, r);
        if (badgeClass && val) return `<td><span class="badge ${badgeClass}">${escapeHtml(val)}</span></td>`;
        return `<td>${escapeHtml(val)}</td>`;
      }).join('');
      tbody.appendChild(tr);
    }

    function repaint() {
      tbody.innerHTML = '';
      const vis = getVisibleRows();
      for (const r of vis) appendDom(r);
      updateCount();
    }

    function updateCount() {
      const visible = tbody.children.length;
      liveCount.textContent = rows.length ? `${visible} / ${rows.length} row(s)` : '';
    }

    filter.addEventListener('input', repaint);

    function setEntity(e) {
      if (!e) return;
      Object.assign(entity, e);
      title.textContent = entity.name || provider.name;
      sub.textContent = entity.sub || '';
    }

    function addRow(r) {
      if (!r) return;
      const k = rowKey(r);
      if (seen.has(k)) return;
      seen.add(k);
      rows.push(r);
      repaint();
      updateHeaderState();
    }
    function addRowIfMissing(r) { addRow(r); }

    function setProgress(p) {
      if (!p) return;
      if (p.label) progressLabel.textContent = p.label;
      if (typeof p.total === 'number' && p.total > 0) {
        progressBox.classList.remove('indeterminate');
        const cur = Math.min(p.current || 0, p.total);
        const pct = Math.round((cur / p.total) * 100);
        progressFill.style.width = pct + '%';
        progressCount.textContent = `${cur} / ${p.total}`;
      } else if (typeof p.current === 'number') {
        progressCount.textContent = `${p.current}`;
      }
    }

    function finish() {
      progressBox.classList.remove('indeterminate');
      progressFill.style.width = '100%';
      progressLabel.textContent = `Done. ${rows.length} row(s).`;
      progressCount.textContent = '';
      // Hide progress block after a short delay.
      setTimeout(() => progressBox.classList.add('hidden'), 600);
      exportBtn.disabled = false;
      const snap = snapshot();
      exportBtn.onclick = () => exportCsv(snap);
      if (!rows.length) emptyEl.classList.remove('hidden');
    }

    function fail(msg) {
      progressBox.classList.remove('indeterminate');
      progressFill.style.width = '100%';
      progressLabel.textContent = 'Failed: ' + (msg || 'unknown error');
      progressLabel.style.color = 'var(--danger)';
    }

    function snapshot() {
      return { entity: { ...entity }, rows: rows.slice(), columns: columns.slice() };
    }

    return { setEntity, addRow, addRowIfMissing, setProgress, finish, fail, snapshot };
  }

  function exportCsv(res) {
    const columns = Array.isArray(res.columns) && res.columns.length
      ? res.columns
      : [
          { key: 'principal', label: 'Principal' },
          { key: 'principalType', label: 'Type' },
          { key: 'role', label: 'Role / Permission' },
          { key: 'scope', label: 'Scope' },
          { key: 'source', label: 'Source' },
        ];
    const headers = columns.map(c => c.label);
    const lines = [headers.join(',')];
    for (const r of res.rows) {
      lines.push(columns.map(c => csvEscape(r[c.key])).join(','));
    }
    const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `permview-${(res.entity.name || 'export').replace(/[^a-z0-9]+/gi, '_')}.csv`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 200);
  }

  function csvEscape(v) {
    const s = (v == null ? '' : String(v));
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // --------- BOOT ---------

  async function boot() {
    initTheme();
    document.getElementById('darkToggle').addEventListener('click', toggleTheme);
    document.getElementById('signInBtn').addEventListener('click', onSignIn);
    document.getElementById('signOutBtn').addEventListener('click', onSignOut);
    document.getElementById('verboseLog').addEventListener('change', (e) => Log.setVerbose(e.target.checked));
    document.getElementById('clearLogBtn').addEventListener('click', () => Log.clear());

    // Top-level tabs (Workloads / Debug / About).
    document.querySelectorAll('.top-tab').forEach(t => {
      t.addEventListener('click', () => {
        document.querySelectorAll('.top-tab').forEach(x => x.classList.toggle('active', x === t));
        document.querySelectorAll('.top-panel').forEach(p => p.classList.toggle('active', p.id === `top-${t.dataset.top}`));
      });
    });

    try {
      await Auth.init();
    } catch (e) {
      Log.err('Auth init failed:', e);
      toast('Auth init failed: ' + e.message, 'err');
    }
    await refreshAuthUI();
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
