// PermView UI. Workloads come from providers.js; each renders a small form, runs, and
// streams rows into one table with filters, sorting and export.
(() => {
  const { h, esc, icon, iconHtml, toast, busy, setUser } = Toolkit;
  const $ = (id) => document.getElementById(id);
  const MAX_ROWS = 2000;
  const PLUG_URL = 'https://www.m365permissions.com';

  const DEFAULT_COLUMNS = [
    { key: 'principal', label: 'Principal' },
    { key: 'principalType', label: 'Type' },
    { key: 'role', label: 'Role / permission' },
    { key: 'scope', label: 'Scope' },
    { key: 'source', label: 'Source' },
  ];

  const state = { providerId: null, running: null, closeMenu: null };

  // Text with every "m365permissions.com" turned into a link.
  function withPlugLinks(text) {
    const parts = String(text).split('m365permissions.com');
    const out = [];
    parts.forEach((p, i) => {
      if (p) out.push(p);
      if (i < parts.length - 1) out.push(h('a', { href: PLUG_URL, target: '_blank', rel: 'noopener', text: 'm365permissions.com' }));
    });
    return out;
  }

  // ---------- auth ----------
  function refreshAuthUI() {
    const signedIn = !!Auth.getAccount();
    setUser(Auth.userInfo(onSignOut));
    $('connectCard').hidden = signedIn;
    $('workspace').hidden = !signedIn;
    if (signedIn) {
      showTenant();
      if (!state.providerId) selectProvider(Providers.list[0].id);
    }
  }
  async function onSignIn() {
    try { await Auth.signIn('read'); }
    catch (e) { Log.err('Sign-in failed:', e); toast(`Sign-in failed: ${e.message}`, 'err'); }
  }
  async function onSignOut() {
    if (state.running) state.running.cancel();
    state.providerId = null;
    $('providerPanel').replaceChildren();
    await Auth.signOut();
  }

  // Name and id of the tenant every result comes from, next to the account.
  async function showTenant() {
    const t = Auth.tenantInfo();
    const line = $('tenantLine');
    line.textContent = `Signed in as ${t.upn}, tenant ${t.tenantId}`;
    try {
      const org = await Api.Graph('read').json('/organization?$select=id,displayName');
      const o = org && org.value && org.value[0];
      if (o) line.textContent = `Signed in as ${t.upn}, tenant ${o.displayName} (${o.id})`;
    } catch (e) { Log.warn(`Could not read the tenant name: ${e.message}`); }
  }

  // ---------- workload picker ----------
  function renderProviderTabs() {
    const nav = $('providerTabs');
    nav.replaceChildren(...Providers.list.map((p) => {
      const active = p.id === state.providerId;
      return h('button.chip-btn', {
        type: 'button', class: active ? 'active' : '', 'aria-pressed': active ? 'true' : 'false',
        on: { click: () => selectProvider(p.id) },
      }, [icon(p.icon, 'sm'), p.name]);
    }));
  }

  function selectProvider(id) {
    if (state.running) state.running.cancel();
    if (state.closeMenu) state.closeMenu();
    state.providerId = id;
    renderProviderTabs();
    renderProviderPanel();
  }

  function consentLine(provider) {
    if (provider.atSignIn) {
      return h('p.muted.small', null, ['Uses the permission granted at sign-in: ', ...provider.scopes.map((s) => h('code', { text: s }))]);
    }
    const codes = [];
    provider.scopes.forEach((s, i) => { if (i) codes.push(', '); codes.push(h('code', { text: s })); });
    return h('p.muted.small', null, ['Asks for ', ...codes, ' the first time you run this workload (consent popup). Read only.']);
  }

  function buildSelect(f) {
    const select = h('select', { name: f.name, id: `f-${f.name}` });
    const reload = h('button.btn.btn-icon', { type: 'button', title: `Reload ${f.label.toLowerCase()} list`, 'aria-label': `Reload ${f.label.toLowerCase()} list` }, icon('i-refresh'));
    async function load() {
      select.disabled = true;
      select.replaceChildren(h('option', { value: '', text: 'Loading...' }));
      await busy(reload, async () => {
        try {
          const opts = await f.loadOptions();
          select.replaceChildren(
            h('option', { value: '', text: opts.length ? 'Choose' : 'Nothing found' }),
            ...opts.map((o) => h('option', { value: o.value, text: o.label })),
          );
          select.disabled = false;
        } catch (e) {
          Log.err(`Loading ${f.label} failed:`, e);
          select.replaceChildren(h('option', { value: '', text: `Failed: ${e.message}` }));
        }
      });
    }
    reload.addEventListener('click', load);
    load();
    return h('div.select-row', null, [select, reload]);
  }

  function renderProviderPanel() {
    const provider = Providers.byId(state.providerId);
    const root = $('providerPanel');
    root.replaceChildren();
    if (!provider) return;

    const inputs = {};
    const fields = provider.form.map((f) => {
      const control = f.kind === 'select'
        ? buildSelect(f)
        : h('input', { type: 'text', name: f.name, id: `f-${f.name}`, placeholder: f.placeholder || null, autocomplete: 'off', spellcheck: 'false' });
      inputs[f.name] = () => root.querySelector(`[name="${f.name}"]`).value;
      return h('div.field', null, [h('label', { for: `f-${f.name}`, text: f.label }), control]);
    });

    const runBtn = h('button.btn.btn-primary', { type: 'submit' }, [icon('i-play'), 'Show permissions']);
    const status = h('span.muted.small');
    const form = h('form.provider-form', null, [...fields, h('div.row.tight', null, [runBtn, status])]);

    const card = h('div.card', null, [
      h('div.card-head', null, h('div', null, [
        h('h2', null, [icon(provider.icon), ` ${provider.name}`]),
        h('p.muted', { text: provider.description }),
      ])),
      consentLine(provider),
      h('div.notice.notice-warn', null, [icon('i-alert'), h('div', null, [h('strong', { text: 'Quick view: ' }), ...withPlugLinks(provider.notice)])]),
      form,
    ]);
    const results = h('div.card#resultsCard', { hidden: true });
    root.append(card, results);

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      if (state.running) return;
      const values = {};
      for (const [k, get] of Object.entries(inputs)) values[k] = get().trim();
      const cancelToken = new Concurrency.CancelToken();
      state.running = cancelToken;
      const stream = startResults(provider, results, cancelToken);
      status.textContent = '';
      await busy(runBtn, async () => {
        try {
          await provider.run(values, {
            cancelToken,
            onRow: (r) => { cancelToken.throwIfCancelled(); stream.addRow(r); },
            onProgress: (p) => { cancelToken.throwIfCancelled(); stream.setProgress(p); },
            setEntity: (en) => stream.setEntity(en),
          });
          cancelToken.throwIfCancelled();
          stream.finish();
          status.textContent = `${stream.count()} row(s)`;
        } catch (err) {
          if (cancelToken.cancelled || err.message === 'Cancelled') {
            stream.fail('Stopped. The table shows what was read until then.', true);
            status.textContent = 'Stopped';
          } else {
            Log.err(`${provider.name} failed:`, err);
            toast(`${provider.name} failed: ${err.message}`, 'err');
            stream.fail(err.message);
            status.textContent = 'Failed, see the Debug tab';
          }
        } finally {
          if (state.running === cancelToken) state.running = null;
        }
      });
    });
  }

  // ---------- results ----------
  const RISK_BADGE = { high: 'danger', medium: 'warn', low: 'ok' };
  const CRED_BADGE = { critical: 'danger', high: 'warn', ok: 'ok' };
  const MODE_BADGE = { application: 'info', delegated: 'violet' };

  function badgeFor(key, value, r) {
    const v = String(value || '').toLowerCase();
    if (key === 'riskLevel') return RISK_BADGE[v] || 'neutral';
    if (key === 'credentialState') return CRED_BADGE[String(r.credentialSeverity || '').toLowerCase()] || 'neutral';
    if (key === 'permissionMode') return MODE_BADGE[v] || '';
    return '';
  }

  function compareValues(a, b) {
    const na = Number(a), nb = Number(b);
    if (a !== '' && b !== '' && Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
    return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
  }
  const cell = (r, key) => (r[key] == null ? '' : String(r[key]));

  function startResults(provider, card, cancelToken) {
    card.hidden = false;
    const columns = provider.columns && provider.columns.length ? provider.columns : DEFAULT_COLUMNS;
    const entity = { name: provider.name, sub: '' };
    const rows = [];
    const seen = new Set();
    const columnFilters = new Map(columns.map((c) => [c.key, null]));
    let sort = { key: null, dir: null };
    let quick = null;
    let paintQueued = false;
    const startedAt = Date.now();

    const title = h('h2', { text: entity.name });
    const sub = h('p.muted.small');
    const xlsxBtn = h('button.btn.btn-small', { type: 'button', disabled: true }, [icon('i-table'), 'Excel']);
    const csvBtn = h('button.btn.btn-small', { type: 'button', disabled: true }, [icon('i-download'), 'CSV']);

    const progressLabel = h('span', { text: 'Starting' });
    const progressEta = h('span.muted.small');
    const fill = h('div');
    const bar = h('div.progress-bar.indeterminate', null, fill);
    const progressCount = h('span');
    const cancelBtn = h('button.btn.btn-small', { type: 'button', on: { click: () => cancelToken.cancel() } }, [icon('i-stop'), 'Stop']);
    const progressBox = h('div.progress-box', null, [
      h('div.progress-header', null, [progressLabel, progressEta]),
      bar,
      h('div.progress-footer', null, [progressCount, cancelBtn]),
    ]);

    const filter = h('input', { type: 'search', placeholder: 'Filter rows', 'aria-label': 'Filter rows' });
    const liveCount = h('span.muted.small');
    const filterRow = h('div.row', null, [filter, liveCount]);

    let quickRow = null;
    if (provider.quickFilters && provider.quickFilters.length) {
      const chips = [{ id: 'all', label: 'All', predicate: null }, ...provider.quickFilters].map((q) => {
        const b = h('button.chip-btn', { type: 'button', class: q.id === 'all' ? 'active' : '', text: q.label });
        b.addEventListener('click', () => {
          quick = q.predicate;
          chips.forEach((x) => x.classList.toggle('active', x === b));
          repaint();
        });
        return b;
      });
      quickRow = h('div.row.quick-filters', null, chips);
    }

    const table = h('table.data-table.pv-table');
    table.innerHTML = `<thead><tr>${columns.map((c) => `<th>
        <button class="th-btn" data-col-key="${esc(c.key)}" type="button" title="Sort and filter">
          <span>${esc(c.label)}</span><span class="th-sort">${iconHtml('i-sort', 'sm')}</span>
        </button></th>`).join('')}</tr></thead><tbody></tbody>`;
    const tbody = table.querySelector('tbody');
    const hint = h('p.table-hint', { hidden: true });
    const empty = h('div.empty-state', { hidden: true }, [icon('i-search'), h('h3', { text: 'No permissions found' }), h('p', { text: 'Nothing came back for this entity, or you cannot read it.' })]);

    card.replaceChildren(
      h('div.card-head', null, [h('div', null, [title, sub]), h('div.row.tight', null, [xlsxBtn, csvBtn])]),
      h('div.plug-banner', null, [icon('i-sparkle'), h('div', null, [
        'Quick view, top level only. For audits across every workload, every scope and over time, ',
        h('a', { href: PLUG_URL, target: '_blank', rel: 'noopener', text: 'm365permissions.com' }),
        ' goes all the way down.',
      ])]),
      ...[progressBox, filterRow, quickRow, h('div.table-wrap', null, table), hint, empty].filter(Boolean),
    );

    // ---------- column menu ----------
    const menu = h('div.th-menu', { hidden: true });
    document.body.appendChild(menu);
    let menuKey = null;
    let menuAnchor = null;
    function closeMenu() { menu.hidden = true; menuKey = null; menuAnchor = null; }
    function removeMenu() { closeMenu(); menu.remove(); document.removeEventListener('click', onDocClick); document.removeEventListener('keydown', onKey); window.removeEventListener('resize', place); window.removeEventListener('scroll', place, true); }
    if (state.closeMenu) state.closeMenu();
    state.closeMenu = removeMenu;

    function place() {
      if (!menuAnchor || menu.hidden) return;
      const b = menuAnchor.getBoundingClientRect();
      const m = menu.getBoundingClientRect();
      const pad = 8;
      let left = Math.max(pad, Math.min(b.left, window.innerWidth - m.width - pad));
      let top = b.bottom + 6;
      if (top + m.height > window.innerHeight - pad) top = Math.max(pad, b.top - m.height - 6);
      menu.style.left = `${Math.round(left)}px`;
      menu.style.top = `${Math.round(top)}px`;
    }
    function onDocClick(ev) {
      if (menu.hidden || menu.contains(ev.target) || (ev.target.closest && ev.target.closest('.th-btn'))) return;
      closeMenu();
    }
    function onKey(ev) { if (ev.key === 'Escape' && !menu.hidden) closeMenu(); }
    document.addEventListener('click', onDocClick);
    document.addEventListener('keydown', onKey);
    window.addEventListener('resize', place);
    window.addEventListener('scroll', place, true);

    table.querySelector('thead').addEventListener('click', (ev) => {
      const btn = ev.target.closest('.th-btn');
      if (!btn) return;
      ev.stopPropagation();
      const key = btn.dataset.colKey;
      if (menuKey === key) { closeMenu(); return; }
      menuKey = key;
      menuAnchor = btn;
      renderMenu();
      menu.hidden = false;
      place();
    });

    function renderMenu() {
      const key = menuKey;
      const values = [...new Set(rows.map((r) => cell(r, key)))].sort(compareValues);
      const selected = columnFilters.get(key);
      const action = (act, label) => `<button type="button" class="btn btn-small" data-act="${act}">${esc(label)}</button>`;
      menu.innerHTML = `
        <div class="th-menu-actions">${action('sort-asc', 'Sort A to Z')}${action('sort-desc', 'Sort Z to A')}${action('clear-sort', 'No sort')}</div>
        <div class="th-menu-actions">${action('select-all', 'Select all')}${action('clear-filter', 'Select none')}</div>
        <input type="search" class="th-menu-search" placeholder="Find value" aria-label="Find value" />
        <div class="th-menu-values"></div>`;
      const list = menu.querySelector('.th-menu-values');
      const search = menu.querySelector('.th-menu-search');
      const paintValues = () => {
        const q = search.value.trim().toLowerCase();
        const shown = values.filter((v) => !q || v.toLowerCase().includes(q));
        list.innerHTML = shown.map((v) => `<label class="checkbox"><input type="checkbox" data-value="${esc(v)}" ${!selected || selected.has(v) ? 'checked' : ''}/> <span>${esc(v || '(blank)')}</span></label>`).join('')
          || '<p class="muted small">No values</p>';
      };
      paintValues();
      search.addEventListener('input', paintValues);
      list.addEventListener('change', (ev) => {
        const cb = ev.target.closest('input[data-value]');
        if (!cb) return;
        const next = new Set(columnFilters.get(key) || values);
        if (cb.checked) next.add(cb.dataset.value); else next.delete(cb.dataset.value);
        columnFilters.set(key, next.size >= values.length ? null : next);
        repaint();
      });
      menu.querySelectorAll('[data-act]').forEach((b) => b.addEventListener('click', () => {
        const a = b.dataset.act;
        if (a === 'sort-asc') sort = { key, dir: 'asc' };
        else if (a === 'sort-desc') sort = { key, dir: 'desc' };
        else if (a === 'clear-sort' && sort.key === key) sort = { key: null, dir: null };
        else if (a === 'select-all') columnFilters.set(key, null);
        else if (a === 'clear-filter') columnFilters.set(key, new Set());
        repaint();
        renderMenu();
        place();
      }));
    }

    // ---------- table ----------
    function matches(r, ft) {
      if (quick && !quick(r)) return false;
      for (const c of columns) {
        const sel = columnFilters.get(c.key);
        if (sel && !sel.has(cell(r, c.key))) return false;
      }
      return !ft || columns.map((c) => cell(r, c.key)).join(' ').toLowerCase().includes(ft);
    }
    function visibleRows() {
      const ft = filter.value.trim().toLowerCase();
      const vis = rows.filter((r) => matches(r, ft));
      if (sort.key) {
        const m = sort.dir === 'asc' ? 1 : -1;
        vis.sort((a, b) => m * compareValues(cell(a, sort.key), cell(b, sort.key)));
      }
      return vis;
    }
    function rowHtml(r) {
      return `<tr>${columns.map((c) => {
        const v = cell(r, c.key);
        const b = v ? badgeFor(c.key, v, r) : '';
        return b ? `<td><span class="badge ${b}">${esc(v)}</span></td>` : `<td>${esc(v)}</td>`;
      }).join('')}</tr>`;
    }
    function paint() {
      paintQueued = false;
      const vis = visibleRows();
      const shown = vis.slice(0, MAX_ROWS);
      tbody.innerHTML = shown.map(rowHtml).join('');
      liveCount.textContent = rows.length ? `${vis.length} of ${rows.length} row(s)` : '';
      hint.hidden = vis.length <= MAX_ROWS;
      hint.textContent = `Showing the first ${MAX_ROWS} of ${vis.length}. Filter to narrow it down; the export has every row.`;
      table.querySelectorAll('.th-btn').forEach((btn) => {
        const key = btn.dataset.colKey;
        const sel = columnFilters.get(key);
        btn.classList.toggle('filtered', !!sel);
        btn.classList.toggle('sorted', sort.key === key);
        const ico = sort.key === key ? (sort.dir === 'asc' ? 'i-arrow-up' : 'i-arrow-down') : 'i-sort';
        const slot = btn.querySelector('.th-sort');
        if (slot.dataset.icon !== ico) { slot.dataset.icon = ico; slot.innerHTML = iconHtml(ico, 'sm'); }
      });
    }
    // Rows stream in fast; paint at most once per frame.
    function repaint() {
      if (paintQueued) return;
      paintQueued = true;
      requestAnimationFrame(paint);
    }
    filter.addEventListener('input', repaint);

    function exportRows(kind) {
      const name = String(entity.name || provider.name).replace(/[^a-z0-9]+/gi, '_').replace(/^_|_$/g, '');
      const base = `PermView-${name}-${Export.stamp()}`;
      const cols = columns.map((c) => ({ key: c.key, label: c.label }));
      if (kind === 'csv') Export.csv(rows, { columns: cols, filename: `${base}.csv` });
      else return Export.xlsx([{ name: provider.name, rows, columns: cols }], `${base}.xlsx`);
      return null;
    }
    xlsxBtn.addEventListener('click', () => busy(xlsxBtn, async () => {
      try { await exportRows('xlsx'); } catch (e) { Log.err('Excel export failed:', e); toast(`Excel export failed: ${e.message}`, 'err'); }
    }));
    csvBtn.addEventListener('click', () => exportRows('csv'));

    function endRun() {
      cancelBtn.hidden = true;
      bar.classList.remove('indeterminate');
      fill.style.width = '100%';
      progressEta.textContent = '';
      xlsxBtn.disabled = csvBtn.disabled = rows.length === 0;
      paint();
    }

    return {
      count: () => rows.length,
      setEntity(en) {
        if (!en) return;
        Object.assign(entity, en);
        title.textContent = entity.name || provider.name;
        sub.textContent = entity.sub || '';
      },
      addRow(r) {
        if (!r) return;
        const k = columns.map((c) => cell(r, c.key)).join('|');
        if (seen.has(k)) return;
        seen.add(k);
        rows.push(r);
        repaint();
      },
      setProgress(p) {
        if (!p) return;
        if (p.label) progressLabel.textContent = p.label;
        if (typeof p.total === 'number' && p.total > 0) {
          bar.classList.remove('indeterminate');
          const cur = Math.min(p.current || 0, p.total);
          fill.style.width = `${Math.round((cur / p.total) * 100)}%`;
          progressCount.textContent = `${cur} of ${p.total}`;
          progressEta.textContent = Concurrency.eta(startedAt, cur, p.total);
        } else {
          bar.classList.add('indeterminate');
          progressCount.textContent = typeof p.current === 'number' ? `${p.current} done` : '';
          progressEta.textContent = '';
        }
      },
      finish() {
        endRun();
        progressLabel.textContent = `Done, ${rows.length} row(s) in ${((Date.now() - startedAt) / 1000).toFixed(1)}s.`;
        progressCount.textContent = '';
        setTimeout(() => { progressBox.hidden = true; }, 800);
        empty.hidden = rows.length > 0;
        Log.info(`${provider.name}: ${rows.length} row(s) for ${entity.sub || entity.name}`);
      },
      fail(msg, stopped) {
        endRun();
        progressBox.classList.add(stopped ? 'is-stopped' : 'is-failed');
        progressLabel.textContent = stopped ? msg : `Failed: ${msg || 'unknown error'}`;
      },
    };
  }

  // ---------- boot ----------
  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    Toolkit.tabs('#tabs');
    $('signInBtn').addEventListener('click', onSignIn);
    try { await Auth.init(); }
    catch (e) { Log.err('Auth init failed:', e); toast(`Sign-in is not available: ${e.message}`, 'err'); }
    refreshAuthUI();
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
