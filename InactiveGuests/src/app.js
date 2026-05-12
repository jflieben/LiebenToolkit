// InactiveGuests - SPA wiring.
(() => {
  const state = {
    guests: [],
    filtered: [],
    selected: new Set(),
    sort: { col: 'inactiveDays', dir: 'desc' },
    deleted: [],
    cancelToken: null,
  };

  // ---------- THEME ----------
  function initTheme() {
    const saved = localStorage.getItem('inactiveguests-theme') || 'light';
    document.documentElement.setAttribute('data-theme', saved);
  }
  function toggleTheme() {
    const cur = document.documentElement.getAttribute('data-theme');
    const nxt = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', nxt);
    localStorage.setItem('inactiveguests-theme', nxt);
  }

  // ---------- TABS ----------
  function activateTab(id) {
    document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === id));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `tab-${id}`));
  }

  // ---------- TOAST + MODAL ----------
  function toast(msg, kind = 'info') {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.className = `toast toast-${kind}`;
    el.classList.remove('hidden');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.add('hidden'), 4000);
  }
  function confirmModal(title, htmlBody, okLabel = 'Confirm') {
    return new Promise(resolve => {
      document.getElementById('confirmTitle').textContent = title;
      document.getElementById('confirmBody').innerHTML = htmlBody;
      const ok = document.getElementById('confirmOkBtn');
      ok.textContent = okLabel;
      const modal = document.getElementById('confirmModal');
      modal.classList.remove('hidden');
      const cleanup = (val) => {
        modal.classList.add('hidden');
        ok.removeEventListener('click', okFn);
        document.getElementById('confirmCancelBtn').removeEventListener('click', cancelFn);
        document.getElementById('confirmCloseBtn').removeEventListener('click', cancelFn);
        resolve(val);
      };
      const okFn = () => cleanup(true);
      const cancelFn = () => cleanup(false);
      ok.addEventListener('click', okFn);
      document.getElementById('confirmCancelBtn').addEventListener('click', cancelFn);
      document.getElementById('confirmCloseBtn').addEventListener('click', cancelFn);
    });
  }

  // ---------- AUTH ----------
  async function refreshAuthUI() {
    const acc = Auth.getAccount();
    const userBox = document.getElementById('userBox');
    const userName = document.getElementById('userName');
    const connectCard = document.getElementById('connectCard');
    const browseCard = document.getElementById('browseCard');
    if (acc) {
      userBox.classList.remove('hidden');
      userName.textContent = acc.username || acc.name || 'Signed in';
      connectCard.classList.add('hidden');
      browseCard.classList.remove('hidden');
    } else {
      userBox.classList.add('hidden');
      connectCard.classList.remove('hidden');
      browseCard.classList.add('hidden');
    }
  }
  async function onSignIn() {
    try { await Auth.signIn(); await refreshAuthUI(); }
    catch (e) { Log.err('Sign-in failed:', e); toast('Sign-in failed: ' + e.message, 'err'); }
  }
  async function onSignOut() {
    await Auth.signOut();
    state.guests = []; state.filtered = []; state.selected.clear();
    renderTable();
    await refreshAuthUI();
  }

  // ---------- LOAD GUESTS ----------
  async function loadGuests() {
    const btn = document.getElementById('loadGuestsBtn');
    btn.disabled = true;
    const status = document.getElementById('loadStatus');
    const box = document.getElementById('progressBox');
    const fill = document.getElementById('progressFill');
    const count = document.getElementById('progressCount');
    box.classList.remove('hidden');
    fill.style.width = '0%';
    count.textContent = '0';
    state.cancelToken = new Concurrency.CancelToken();
    document.getElementById('cancelBtn').onclick = () => state.cancelToken.cancel();
    const t0 = Date.now();
    try {
      const raw = await Graph.listAllGuests((loaded, page) => {
        if (state.cancelToken.cancelled) throw new Error('Cancelled');
        count.textContent = `${loaded} loaded (${page} pages)`;
        fill.style.width = Math.min(95, page * 8) + '%';
      });
      const now = new Date();
      state.guests = raw.map(g => Guests.enrich(g, now));
      fill.style.width = '100%';
      const dt = ((Date.now() - t0) / 1000).toFixed(1);
      status.textContent = `Loaded ${state.guests.length} guests in ${dt}s`;
      Log.info(`Loaded ${state.guests.length} guests in ${dt}s`);
      // Surface the P1/P2 caveat if signInActivity is empty for nearly everyone.
      const hasSignIn = state.guests.filter(g => g.lastSignIn).length;
      if (state.guests.length > 0 && hasSignIn / state.guests.length < 0.1) {
        document.getElementById('ownerNotice').classList.remove('hidden');
      } else {
        document.getElementById('ownerNotice').classList.add('hidden');
      }
      document.getElementById('filterRow').classList.remove('hidden');
      applyFilter();
    } catch (e) {
      Log.err('Load failed:', e);
      toast('Load failed: ' + e.message, 'err');
    } finally {
      btn.disabled = false;
      setTimeout(() => box.classList.add('hidden'), 600);
    }
  }

  function readFilter() {
    return {
      search: document.getElementById('search').value.trim(),
      minInactive: parseInt(document.getElementById('minInactive').value, 10) || 0,
      minAge: parseInt(document.getElementById('minAge').value, 10) || 0,
      redemption: document.getElementById('redemption').value,
      enabled: document.getElementById('enabled').value,
      neverSignedIn: document.getElementById('neverSignedIn').value,
    };
  }
  function applyFilter() {
    const f = readFilter();
    state.filtered = state.guests.filter(g => Guests.passesFilter(g, f));
    // Trim selection to what's still visible.
    for (const id of [...state.selected]) {
      if (!state.filtered.some(g => g.id === id)) state.selected.delete(id);
    }
    renderTable();
  }

  // ---------- TABLE ----------
  function sortGuests(rows) {
    const { col, dir } = state.sort;
    const m = dir === 'asc' ? 1 : -1;
    const k = g => {
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
    };
    return rows.slice().sort((a, b) => {
      const ka = k(a), kb = k(b);
      return ka > kb ? m : ka < kb ? -m : 0;
    });
  }

  function renderTable() {
    const wrap = document.getElementById('tableWrap');
    const hint = document.getElementById('tableHint');
    const bar = document.getElementById('actionBar');
    const tbody = document.querySelector('#guestsTable tbody');
    if (!state.guests.length) {
      wrap.classList.add('hidden'); hint.classList.add('hidden'); bar.classList.add('hidden');
      return;
    }
    wrap.classList.remove('hidden'); hint.classList.remove('hidden'); bar.classList.remove('hidden');
    const rows = sortGuests(state.filtered).slice(0, 1000);
    hint.textContent = `${state.filtered.length} match${state.filtered.length === 1 ? '' : 'es'}` +
      (state.filtered.length > rows.length ? ` (showing first ${rows.length})` : '');
    tbody.innerHTML = rows.map(g => `
      <tr data-id="${escAttr(g.id)}">
        <td><input type="checkbox" data-act="sel" ${state.selected.has(g.id) ? 'checked' : ''}/></td>
        <td><strong>${escHtml(g.displayName || '(no name)')}</strong>${g.companyName ? `<br><span class="muted small">${escHtml(g.companyName)}</span>` : ''}</td>
        <td><code>${escHtml(g.userPrincipalName || '')}</code>${g.mail && g.mail !== g.userPrincipalName ? `<br><span class="muted small">${escHtml(g.mail)}</span>` : ''}</td>
        <td>${escHtml(g.redemptionState)}</td>
        <td>${g.accountEnabled ? 'Yes' : '<span class="muted">No</span>'}</td>
        <td>${escHtml(Guests.fmtDate(g.createdDateTime))}<br><span class="muted small">${g.accountAgeDays != null ? g.accountAgeDays + ' days' : ''}</span></td>
        <td>${g.lastSignIn ? escHtml(Guests.fmtDate(g.lastSignIn)) : '<span class="muted">Never</span>'}</td>
        <td>${g.inactiveDays == null ? '' : g.inactiveDays + ' days'}</td>
      </tr>
    `).join('');
    document.querySelectorAll('#guestsTable thead th').forEach(th => {
      th.classList.remove('sort-asc','sort-desc','sortable-col');
      if (th.dataset.sort) {
        th.classList.add('sortable-col');
        if (th.dataset.sort === state.sort.col) th.classList.add(state.sort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
      }
    });
    updateSelCount();
    document.getElementById('selectAll').checked = rows.length > 0 && rows.every(g => state.selected.has(g.id));
  }

  function updateSelCount() {
    document.getElementById('selCount').textContent = String(state.selected.size);
    const sel = state.selected.size;
    document.getElementById('disableBtn').disabled = sel === 0;
    document.getElementById('enableBtn').disabled = sel === 0;
    document.getElementById('deleteBtn').disabled = sel === 0;
    document.getElementById('exportXlsxBtn').disabled = state.filtered.length === 0;
    document.getElementById('exportCsvBtn').disabled = state.filtered.length === 0;
  }

  function onTableChange(e) {
    const cb = e.target.closest('input[data-act="sel"]');
    if (!cb) return;
    const id = cb.closest('tr').dataset.id;
    if (cb.checked) state.selected.add(id); else state.selected.delete(id);
    updateSelCount();
  }
  function onSortClick(e) {
    const th = e.target.closest('th[data-sort]');
    if (!th) return;
    const col = th.dataset.sort;
    if (state.sort.col === col) state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
    else { state.sort.col = col; state.sort.dir = col === 'inactiveDays' ? 'desc' : 'asc'; }
    renderTable();
  }
  function onSelectAll(e) {
    const checked = e.target.checked;
    const rows = sortGuests(state.filtered).slice(0, 1000);
    for (const g of rows) {
      if (checked) state.selected.add(g.id); else state.selected.delete(g.id);
    }
    renderTable();
  }

  // ---------- ACTIONS ----------
  function selectedGuests() {
    return state.guests.filter(g => state.selected.has(g.id));
  }
  function previewList(rows, max = 20) {
    const items = rows.slice(0, max).map(g => `<li>${escHtml(g.displayName || g.userPrincipalName)} <span class="muted small">${escHtml(g.userPrincipalName || '')}</span></li>`).join('');
    return `<ul>${items}</ul>` + (rows.length > max ? `<p class="muted small">+ ${rows.length - max} more</p>` : '');
  }

  async function bulkAction(label, fn, danger = false) {
    const rows = selectedGuests();
    if (!rows.length) return;
    const ok = await confirmModal(
      `${label} ${rows.length} guest(s)?`,
      `<p>You are about to <strong>${label.toLowerCase()}</strong> the following ${rows.length} guest user(s):</p>${previewList(rows)}` +
      (danger ? '<p class="notice notice-warn">Deleted users go to the recycle bin for 30 days, then are purged. They lose access to all M365 resources immediately.</p>' : ''),
      label
    );
    if (!ok) return;
    const resultsEl = document.getElementById('actionResults');
    resultsEl.classList.remove('hidden');
    resultsEl.innerHTML = '';
    const cancelToken = new Concurrency.CancelToken();
    let done = 0;
    const total = rows.length;
    await Concurrency.pmap(rows, async (g) => {
      const li = document.createElement('li');
      li.innerHTML = `<span class="ok-mark">...</span> ${escHtml(g.displayName || g.userPrincipalName)}`;
      resultsEl.appendChild(li);
      try {
        await fn(g);
        li.innerHTML = `<span class="ok-mark">&#10003;</span> ${label} - ${escHtml(g.displayName || g.userPrincipalName)}`;
      } catch (e) {
        li.innerHTML = `<span class="err-mark">&#10007;</span> ${escHtml(g.displayName || g.userPrincipalName)} - ${escHtml(e.message)}`;
        Log.err(`${label} failed for ${g.userPrincipalName}:`, e);
      } finally {
        done++;
      }
    }, { concurrency: 4, cancelToken });
    toast(`${label} done: ${done}/${total}`, 'ok');
    state.selected.clear();
    await loadGuests();
  }

  function onDisable() {
    return bulkAction('Disable', g => Graph.disableUser(g.id));
  }
  function onEnable() {
    return bulkAction('Enable', g => Graph.enableUser(g.id));
  }
  function onDelete() {
    return bulkAction('Delete', g => Graph.deleteUser(g.id), true);
  }
  function onExportXlsx() { return Export.exportXlsx(state.filtered, `inactive-guests-${new Date().toISOString().substring(0,10)}.xlsx`); }
  function onExportCsv()  { return Export.exportCsv(state.filtered,  `inactive-guests-${new Date().toISOString().substring(0,10)}.csv`); }

  // ---------- RECYCLE BIN ----------
  async function loadDeleted() {
    const btn = document.getElementById('loadDeletedBtn');
    btn.disabled = true;
    const status = document.getElementById('loadDelStatus');
    status.textContent = 'Loading...';
    try {
      const list = await Graph.listDeletedGuests((n) => { status.textContent = `${n} loaded`; });
      state.deleted = list;
      status.textContent = `${list.length} deleted guest(s)`;
      renderDeleted();
    } catch (e) {
      Log.err('Load deleted failed:', e);
      status.textContent = 'Failed: ' + e.message;
    } finally {
      btn.disabled = false;
    }
  }
  function renderDeleted() {
    const wrap = document.getElementById('delTableWrap');
    const hint = document.getElementById('delHint');
    const tbody = document.querySelector('#deletedTable tbody');
    if (!state.deleted.length) {
      wrap.classList.add('hidden');
      hint.classList.remove('hidden');
      hint.textContent = 'Recycle bin is empty.';
      return;
    }
    wrap.classList.remove('hidden');
    hint.classList.add('hidden');
    const now = new Date();
    tbody.innerHTML = state.deleted.map(u => {
      const deleted = u.deletedDateTime ? new Date(u.deletedDateTime) : null;
      const purgeDays = deleted ? Math.max(0, 30 - Guests.daysBetween(deleted, now)) : '?';
      return `<tr data-id="${escAttr(u.id)}">
        <td>${escHtml(u.displayName || '(no name)')}</td>
        <td><code>${escHtml(u.userPrincipalName || '')}</code></td>
        <td>${escHtml(Guests.fmtDate(deleted))}</td>
        <td>${purgeDays} day(s)</td>
        <td><button class="ghost-btn" data-act="restore">Restore</button></td>
      </tr>`;
    }).join('');
  }
  async function onRestoreClick(e) {
    const btn = e.target.closest('button[data-act="restore"]');
    if (!btn) return;
    const tr = btn.closest('tr');
    const u = state.deleted.find(x => x.id === tr.dataset.id);
    if (!u) return;
    const ok = await confirmModal(
      'Restore deleted guest?',
      `<p>Restore <strong>${escHtml(u.displayName || u.userPrincipalName)}</strong>? The account will be re-enabled and group memberships should come back along with it.</p>`,
      'Restore'
    );
    if (!ok) return;
    btn.disabled = true;
    try {
      await Graph.restoreUser(u.id);
      Log.info(`Restored ${u.userPrincipalName}`);
      toast('Restored', 'ok');
      await loadDeleted();
    } catch (e) {
      Log.err('Restore failed:', e);
      toast('Restore failed: ' + e.message, 'err');
      btn.disabled = false;
    }
  }

  // ---------- HELPERS ----------
  function escHtml(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function escAttr(s) { return escHtml(s); }

  // ---------- BOOT ----------
  async function boot() {
    initTheme();
    document.getElementById('darkToggle').addEventListener('click', toggleTheme);
    document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => activateTab(t.dataset.tab)));
    document.getElementById('signInBtn').addEventListener('click', onSignIn);
    document.getElementById('signOutBtn').addEventListener('click', onSignOut);
    document.getElementById('loadGuestsBtn').addEventListener('click', loadGuests);
    document.getElementById('loadDeletedBtn').addEventListener('click', loadDeleted);

    ['search','minInactive','minAge','redemption','enabled','neverSignedIn'].forEach(id => {
      document.getElementById(id).addEventListener('input', applyFilter);
      document.getElementById(id).addEventListener('change', applyFilter);
    });
    document.querySelector('#guestsTable thead').addEventListener('click', onSortClick);
    document.querySelector('#guestsTable tbody').addEventListener('change', onTableChange);
    document.getElementById('selectAll').addEventListener('change', onSelectAll);
    document.querySelector('#deletedTable tbody').addEventListener('click', onRestoreClick);

    document.getElementById('disableBtn').addEventListener('click', onDisable);
    document.getElementById('enableBtn').addEventListener('click', onEnable);
    document.getElementById('deleteBtn').addEventListener('click', onDelete);
    document.getElementById('exportXlsxBtn').addEventListener('click', onExportXlsx);
    document.getElementById('exportCsvBtn').addEventListener('click', onExportCsv);

    document.getElementById('clearLogBtn').addEventListener('click', () => Log.clear());
    document.getElementById('verboseLog').addEventListener('change', e => Log.setVerbose(e.target.checked));

    try {
      await Auth.init();
      Log.info('InactiveGuests ready.');
      await refreshAuthUI();
    } catch (e) {
      Log.err('Boot failed:', e);
      toast('Initialization failed: ' + e.message, 'err');
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
