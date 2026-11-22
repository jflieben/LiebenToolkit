// InactiveGuests UI.
(() => {
  const { esc, toast, confirm, busy, setUser } = Toolkit;

  window.Auth = ToolkitAuth.create({
    scopes: {
      read: ['User.Read', 'User.Read.All', 'AuditLog.Read.All'],
      write: ['User.ReadWrite.All'],
    },
  });

  const state = {
    guests: [],
    filtered: [],
    selected: new Set(),
    sort: { col: 'inactiveDays', dir: 'desc' },
    deleted: [],
    cancelToken: null,
  };
  const $ = (id) => document.getElementById(id);
  const MAX_ROWS = 1000;

  // ---------- auth ----------
  function refreshAuthUI() {
    const signedIn = !!Auth.getAccount();
    setUser(Auth.userInfo(onSignOut));
    $('connectCard').hidden = signedIn;
    $('browseCard').hidden = !signedIn;
  }
  async function onSignIn() {
    try { await Auth.signIn('read'); }
    catch (e) { Log.err('Sign-in failed:', e); toast(`Sign-in failed: ${e.message}`, 'err'); }
  }
  async function onSignOut() {
    state.guests = []; state.filtered = []; state.selected.clear();
    await Auth.signOut();
  }

  // ---------- load ----------
  async function loadGuests() {
    const box = $('progressBox');
    box.hidden = false;
    $('progressCount').textContent = '0 loaded';
    $('progressLabel').textContent = 'Loading guests';
    state.cancelToken = new Concurrency.CancelToken();
    $('cancelBtn').onclick = () => state.cancelToken.cancel();
    const t0 = Date.now();
    await busy($('loadGuestsBtn'), async () => {
      try {
        const raw = await Graph.listAllGuests((loaded, page) => {
          $('progressCount').textContent = `${loaded} loaded, page ${page}`;
        }, state.cancelToken.signal);
        const now = new Date();
        state.guests = raw.map((g) => Guests.enrich(g, now));
        const dt = ((Date.now() - t0) / 1000).toFixed(1);
        $('loadStatus').textContent = `${state.guests.length} guests loaded in ${dt}s. Nothing changes until you pick an action and confirm it.`;
        Log.info(`Loaded ${state.guests.length} guests in ${dt}s`);
        const withSignIn = state.guests.filter((g) => g.lastSignIn).length;
        $('licenseNotice').hidden = !(state.guests.length > 0 && withSignIn / state.guests.length < 0.1);
        $('filterRow').hidden = false;
        renderStats();
        applyFilter();
      } catch (e) {
        if (e.message === 'Cancelled') { toast('Loading cancelled', 'warn'); return; }
        Log.err('Load failed:', e);
        toast(`Loading failed: ${e.message}`, 'err');
      } finally {
        box.hidden = true;
      }
    });
  }

  function renderStats() {
    const g = state.guests;
    $('statTiles').hidden = !g.length;
    $('statTotal').textContent = g.length;
    $('statInactive').textContent = g.filter((x) => x.inactiveDays != null && x.inactiveDays >= 90).length;
    $('statPending').textContent = g.filter((x) => x.redemptionState === 'PendingAcceptance').length;
    $('statDisabled').textContent = g.filter((x) => !x.accountEnabled).length;
  }

  function readFilter() {
    return {
      search: $('search').value.trim(),
      minInactive: parseInt($('minInactive').value, 10) || 0,
      minAge: parseInt($('minAge').value, 10) || 0,
      redemption: $('redemption').value,
      enabled: $('enabled').value,
      neverSignedIn: $('neverSignedIn').value,
    };
  }
  function applyFilter() {
    const f = readFilter();
    state.filtered = state.guests.filter((g) => Guests.passesFilter(g, f));
    // Only what is visible stays selected, so an action never hits hidden rows.
    const visible = new Set(state.filtered.map((g) => g.id));
    for (const id of [...state.selected]) if (!visible.has(id)) state.selected.delete(id);
    renderTable();
  }

  // ---------- table ----------
  function sortKey(g, col) {
    switch (col) {
      case 'displayName': return (g.displayName || '').toLowerCase();
      case 'userPrincipalName': return (g.userPrincipalName || '').toLowerCase();
      case 'redemptionState': return g.redemptionState || '';
      case 'accountEnabled': return g.accountEnabled ? 1 : 0;
      case 'createdDateTime': return g.createdDateTime || '';
      case 'lastSignIn': return g.lastSignIn ? g.lastSignIn.getTime() : -1;
      case 'inactiveDays': return g.inactiveDays == null ? -1 : g.inactiveDays;
      default: return '';
    }
  }
  function sorted(rows) {
    const { col, dir } = state.sort;
    const m = dir === 'asc' ? 1 : -1;
    return rows.slice().sort((a, b) => {
      const ka = sortKey(a, col), kb = sortKey(b, col);
      return ka > kb ? m : ka < kb ? -m : 0;
    });
  }

  const REDEMPTION_BADGE = { Accepted: 'ok', PendingAcceptance: 'warn', Unknown: 'neutral' };
  const REDEMPTION_LABEL = { Accepted: 'Accepted', PendingAcceptance: 'Pending', Unknown: 'Unknown' };

  function renderTable() {
    const has = state.guests.length > 0;
    $('tableWrap').hidden = !has;
    $('tableHint').hidden = !has;
    $('actionBar').hidden = !has;
    const tbody = document.querySelector('#guestsTable tbody');
    if (!has) { tbody.replaceChildren(); return; }
    const rows = sorted(state.filtered).slice(0, MAX_ROWS);
    $('tableHint').textContent = `${state.filtered.length} match${state.filtered.length === 1 ? '' : 'es'}` +
      (state.filtered.length > rows.length ? `, showing the first ${rows.length}` : '');
    tbody.innerHTML = rows.map((g) => `
      <tr data-id="${esc(g.id)}">
        <td class="col-check"><input type="checkbox" data-act="sel" aria-label="Select ${esc(g.displayName || g.userPrincipalName)}" ${state.selected.has(g.id) ? 'checked' : ''}/></td>
        <td><strong>${esc(g.displayName || '(no name)')}</strong>${g.companyName ? `<br><span class="muted small">${esc(g.companyName)}</span>` : ''}</td>
        <td><code>${esc(g.userPrincipalName || '')}</code>${g.mail && g.mail !== g.userPrincipalName ? `<br><span class="muted small">${esc(g.mail)}</span>` : ''}</td>
        <td><span class="badge dot ${REDEMPTION_BADGE[g.redemptionState] || 'neutral'}">${esc(REDEMPTION_LABEL[g.redemptionState] || g.redemptionState)}</span></td>
        <td>${g.accountEnabled ? '<span class="badge ok">Enabled</span>' : '<span class="badge neutral">Disabled</span>'}</td>
        <td class="nowrap">${esc(Guests.fmtDate(g.createdDateTime))}<br><span class="muted small">${g.accountAgeDays != null ? `${g.accountAgeDays} days ago` : ''}</span></td>
        <td class="nowrap">${g.lastSignIn ? esc(Guests.fmtDate(g.lastSignIn)) : '<span class="muted">Never</span>'}</td>
        <td class="num nowrap">${g.inactiveDays == null ? '' : `${g.inactiveDays} d`}</td>
      </tr>`).join('');
    document.querySelectorAll('#guestsTable thead th').forEach((th) => {
      th.classList.remove('sort-asc', 'sort-desc');
      if (th.dataset.sort) {
        th.classList.add('sortable-col');
        if (th.dataset.sort === state.sort.col) th.classList.add(state.sort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
      }
    });
    updateSelection();
    $('selectAll').checked = rows.length > 0 && rows.every((g) => state.selected.has(g.id));
  }

  function updateSelection() {
    const n = state.selected.size;
    $('selCount').textContent = String(n);
    ['disableBtn', 'enableBtn', 'deleteBtn'].forEach((id) => { $(id).disabled = n === 0; });
    ['exportXlsxBtn', 'exportCsvBtn'].forEach((id) => { $(id).disabled = state.filtered.length === 0; });
  }

  function onTableChange(e) {
    const cb = e.target.closest('input[data-act="sel"]');
    if (!cb) return;
    const id = cb.closest('tr').dataset.id;
    if (cb.checked) state.selected.add(id); else state.selected.delete(id);
    $('selectAll').checked = false;
    updateSelection();
  }
  function onSortClick(e) {
    const th = e.target.closest('th[data-sort]');
    if (!th) return;
    const col = th.dataset.sort;
    if (state.sort.col === col) state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
    else state.sort = { col, dir: col === 'inactiveDays' ? 'desc' : 'asc' };
    renderTable();
  }
  function onSelectAll(e) {
    const rows = sorted(state.filtered).slice(0, MAX_ROWS);
    for (const g of rows) { if (e.target.checked) state.selected.add(g.id); else state.selected.delete(g.id); }
    renderTable();
  }

  // ---------- actions ----------
  const label = (g) => `${g.displayName || '(no name)'}  ${g.userPrincipalName || ''}`.trim();

  async function bulkAction(verb, fn, { danger = false, note } = {}) {
    const rows = state.guests.filter((g) => state.selected.has(g.id));
    if (!rows.length) return;
    const ok = await confirm({
      title: `${verb} ${rows.length} guest${rows.length === 1 ? '' : 's'}?`,
      body: `These guest accounts will be ${verb.toLowerCase()}d in the tenant${note ? `. ${note}` : '.'}`,
      items: rows.map(label),
      okLabel: `${verb} ${rows.length}`,
      danger,
    });
    if (!ok) return;
    // Ask for write consent before the first row, not halfway through.
    try { await Auth.getWriteToken(); }
    catch (e) { toast(`No write permission: ${e.message}`, 'err'); return; }

    const list = $('actionResults');
    list.hidden = false;
    list.replaceChildren();
    let ok2 = 0;
    const results = await Concurrency.pmap(rows, async (g) => {
      const li = Toolkit.h('li', null, [Toolkit.h('span.warn-mark', { text: '...' }), ` ${label(g)}`]);
      list.appendChild(li);
      try {
        await fn(g);
        ok2++;
        li.replaceChildren(Toolkit.h('span.ok-mark', { text: '✓' }), ` ${verb}d: ${label(g)}`);
      } catch (e) {
        li.replaceChildren(Toolkit.h('span.err-mark', { text: '✗' }), ` ${label(g)}: ${e.message}`);
        Log.err(`${verb} failed for ${g.userPrincipalName}:`, e);
        throw e;
      }
    }, { concurrency: 4 });
    const failed = results.errors.length;
    toast(`${verb}d ${ok2} of ${rows.length}${failed ? `, ${failed} failed (see the list and the debug log)` : ''}`, failed ? 'warn' : 'ok');
    state.selected.clear();
    await loadGuests();
  }

  // ---------- recycle bin ----------
  async function loadDeleted() {
    $('loadDelStatus').textContent = 'Loading...';
    await busy($('loadDeletedBtn'), async () => {
      try {
        state.deleted = await Graph.listDeletedGuests((n) => { $('loadDelStatus').textContent = `${n} loaded`; });
        $('loadDelStatus').textContent = `${state.deleted.length} deleted guest${state.deleted.length === 1 ? '' : 's'}`;
        renderDeleted();
      } catch (e) {
        Log.err('Loading deleted guests failed:', e);
        $('loadDelStatus').textContent = '';
        toast(`Loading failed: ${e.message}`, 'err');
      }
    });
  }
  function renderDeleted() {
    const empty = !state.deleted.length;
    $('delTableWrap').hidden = empty;
    $('delEmpty').hidden = !empty;
    const now = new Date();
    document.querySelector('#deletedTable tbody').innerHTML = state.deleted.map((u) => {
      const deleted = u.deletedDateTime ? new Date(u.deletedDateTime) : null;
      const purge = deleted ? Math.max(0, 30 - Guests.daysBetween(deleted, now)) : null;
      return `<tr data-id="${esc(u.id)}">
        <td>${esc(u.displayName || '(no name)')}</td>
        <td><code>${esc(u.userPrincipalName || '')}</code></td>
        <td class="nowrap">${esc(Guests.fmtDate(deleted))}</td>
        <td class="nowrap">${purge == null ? '?' : `<span class="badge ${purge <= 5 ? 'danger' : 'neutral'}">${purge} days</span>`}</td>
        <td class="right"><button class="btn btn-small" data-act="restore">Restore</button></td>
      </tr>`;
    }).join('');
    Toolkit.hydrateIcons(document.getElementById('deletedTable'));
  }
  async function onRestoreClick(e) {
    const btn = e.target.closest('button[data-act="restore"]');
    if (!btn) return;
    const u = state.deleted.find((x) => x.id === btn.closest('tr').dataset.id);
    if (!u) return;
    const ok = await confirm({
      title: 'Restore this guest?',
      body: 'The account comes back with its group memberships. It is restored in the state it was deleted in.',
      items: [label(u)],
      okLabel: 'Restore',
    });
    if (!ok) return;
    await busy(btn, async () => {
      try {
        await Graph.restoreUser(u.id);
        Log.info(`Restored ${u.userPrincipalName}`);
        toast(`Restored ${u.displayName || u.userPrincipalName}`, 'ok');
        await loadDeleted();
      } catch (err) {
        Log.err('Restore failed:', err);
        toast(`Restore failed: ${err.message}`, 'err');
      }
    });
  }

  // ---------- export ----------
  const EXPORT_COLUMNS = [
    { key: 'displayName', label: 'Display name' },
    { key: 'userPrincipalName', label: 'UPN' },
    { key: 'mail', label: 'Mail' },
    { key: 'companyName', label: 'Company' },
    { key: 'accountEnabled', label: 'Enabled' },
    { key: 'redemptionState', label: 'Redemption' },
    { key: 'createdDateTime', label: 'Created', format: (v) => Guests.fmtDate(v) },
    { key: 'accountAgeDays', label: 'Age (days)' },
    { key: 'lastSignIn', label: 'Last sign-in', format: (v) => Guests.fmtDate(v) },
    { key: 'inactiveDays', label: 'Inactive (days)' },
    { key: 'id', label: 'Object id' },
  ];
  function exportRows() { return sorted(state.filtered); }

  // ---------- boot ----------
  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    Toolkit.tabs('#tabs');
    $('signInBtn').addEventListener('click', onSignIn);
    $('loadGuestsBtn').addEventListener('click', loadGuests);
    $('loadDeletedBtn').addEventListener('click', loadDeleted);
    ['search', 'minInactive', 'minAge', 'redemption', 'enabled', 'neverSignedIn'].forEach((id) => $(id).addEventListener('input', applyFilter));
    document.querySelector('#guestsTable thead').addEventListener('click', onSortClick);
    document.querySelector('#guestsTable tbody').addEventListener('change', onTableChange);
    $('selectAll').addEventListener('change', onSelectAll);
    document.querySelector('#deletedTable tbody').addEventListener('click', onRestoreClick);

    $('disableBtn').addEventListener('click', () => bulkAction('Disable', (g) => Graph.disableUser(g.id), { note: 'They can no longer sign in until enabled again' }));
    $('enableBtn').addEventListener('click', () => bulkAction('Enable', (g) => Graph.enableUser(g.id)));
    $('deleteBtn').addEventListener('click', () => bulkAction('Delete', (g) => Graph.deleteUser(g.id), { danger: true, note: 'They lose access to everything right away. Restorable from the recycle bin for 30 days' }));
    $('exportXlsxBtn').addEventListener('click', () => busy($('exportXlsxBtn'), () => Export.xlsx([{ name: 'Guests', rows: exportRows(), columns: EXPORT_COLUMNS }], `inactive-guests-${Export.stamp()}.xlsx`).catch((e) => toast(e.message, 'err'))));
    $('exportCsvBtn').addEventListener('click', () => Export.csv(exportRows(), { columns: EXPORT_COLUMNS, filename: `inactive-guests-${Export.stamp()}.csv` }));

    try {
      await Auth.init();
      refreshAuthUI();
      Log.info('InactiveGuests ready.');
    } catch (e) {
      Log.err('Start-up failed:', e);
      toast(`Start-up failed: ${e.message}`, 'err');
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
