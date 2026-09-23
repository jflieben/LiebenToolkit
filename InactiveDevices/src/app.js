// InactiveDevices - SPA wiring.
(() => {
  const state = {
    devices: [],
    filtered: [],
    expired: [],
    selected: new Set(),
    sort: { col: 'inactiveDays', dir: 'desc' },
    cancelToken: null,
  };

  function initTheme() {
    const saved = localStorage.getItem('inactivedevices-theme') || 'light';
    document.documentElement.setAttribute('data-theme', saved);
  }

  function toggleTheme() {
    const cur = document.documentElement.getAttribute('data-theme');
    const nxt = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', nxt);
    localStorage.setItem('inactivedevices-theme', nxt);
  }

  function activateTab(id) {
    document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === id));
    document.querySelectorAll('.tab-panel').forEach((p) => p.classList.toggle('active', p.id === `tab-${id}`));
  }

  function toast(msg, kind = 'info') {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.className = `toast toast-${kind}`;
    el.classList.remove('hidden');
    clearTimeout(toast.t);
    toast.t = setTimeout(() => el.classList.add('hidden'), 4000);
  }

  function loadPrefs() {
    try {
      const p = JSON.parse(localStorage.getItem('inactivedevices-prefs') || '{}');
      if (p.extAttr) document.getElementById('extAttrSelect').value = p.extAttr;
      if (p.gracePeriod != null) document.getElementById('gracePeriod').value = p.gracePeriod;
      if (p.concurrency != null) document.getElementById('concurrency').value = p.concurrency;
      if (p.cleanupMode) {
        const r = document.querySelector(`input[name=cleanupMode][value="${p.cleanupMode}"]`);
        if (r) r.checked = true;
      }
    } catch {}
  }

  function savePrefs() {
    const p = {
      extAttr: document.getElementById('extAttrSelect').value,
      gracePeriod: +document.getElementById('gracePeriod').value || 30,
      concurrency: +document.getElementById('concurrency').value || 4,
      cleanupMode: document.querySelector('input[name=cleanupMode]:checked').value,
    };
    localStorage.setItem('inactivedevices-prefs', JSON.stringify(p));
  }

  function refreshActionLabel() {
    const mode = document.querySelector('input[name=cleanupMode]:checked').value;
    document.getElementById('actionLabel').textContent = mode;
  }

  function confirmModal(title, htmlBody, okLabel = 'Confirm') {
    return new Promise((resolve) => {
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

  function parseDisableStamp(value) {
    if (typeof value !== 'string') return null;
    const m = value.match(/^InactiveDevices:disabled:(.+)$/);
    if (!m) return null;
    const d = new Date(m[1]);
    return isNaN(d.getTime()) ? null : d;
  }

  function buildDisableStampValue() {
    return `InactiveDevices:disabled:${new Date().toISOString()}`;
  }

  function setSignedInUI(acc) {
    const connectCard = document.getElementById('connectCard');
    const discoveryCard = document.getElementById('discoveryCard');
    const userBox = document.getElementById('userBox');
    const userName = document.getElementById('userName');
    if (acc) {
      connectCard.classList.add('hidden');
      discoveryCard.classList.remove('hidden');
      userBox.classList.remove('hidden');
      userName.textContent = acc.username || acc.name || 'Signed in';
      document.getElementById('findExpiredBtn').disabled = false;
    } else {
      connectCard.classList.remove('hidden');
      discoveryCard.classList.add('hidden');
      userBox.classList.add('hidden');
      document.getElementById('findExpiredBtn').disabled = true;
    }
  }

  async function onSignIn() {
    try {
      const acc = await Auth.signIn();
      if (!acc) return;
      setSignedInUI(acc);
      loadPrefs();
      refreshActionLabel();
    } catch (e) {
      Log.err('Sign-in failed:', e);
      toast('Sign-in failed: ' + e.message, 'err');
    }
  }

  async function onSignOut() {
    await Auth.signOut();
    state.devices = [];
    state.filtered = [];
    state.expired = [];
    state.selected.clear();
    setSignedInUI(null);
    renderResults();
    renderExpired();
  }

  async function scanDevices() {
    savePrefs();

    const btn = document.getElementById('scanBtn');
    const status = document.getElementById('loadStatus');
    const box = document.getElementById('progressBox');
    const fill = document.getElementById('progressFill');
    const count = document.getElementById('progressCount');
    const label = document.getElementById('progressLabel');

    btn.disabled = true;
    box.classList.remove('hidden');
    fill.style.width = '0%';
    count.textContent = '0';
    label.textContent = 'Loading Entra devices...';

    state.cancelToken = new Concurrency.CancelToken();
    document.getElementById('cancelBtn').onclick = () => state.cancelToken.cancel();

    const t0 = Date.now();
    try {
      const raw = await Graph.listAllDevices((loaded, page) => {
        if (state.cancelToken.cancelled) throw new Error('Cancelled');
        count.textContent = `${loaded} devices (${page} pages)`;
        fill.style.width = String(Math.min(70, page * 6)) + '%';
      });

      label.textContent = 'Loading Intune enrichment...';
      const managedRes = await Graph.tryListManagedDevices((loaded, page) => {
        if (state.cancelToken.cancelled) throw new Error('Cancelled');
        count.textContent = `${raw.length} devices + ${loaded} managed devices (${page} pages)`;
        fill.style.width = String(Math.min(95, 70 + page * 6)) + '%';
      });

      const noticeEl = document.getElementById('managedNotice');
      if (managedRes.unavailable) {
        noticeEl.textContent = managedRes.message;
        noticeEl.classList.remove('hidden');
      } else {
        noticeEl.classList.add('hidden');
      }

      const managedMap = Devices.buildManagedMap(managedRes.list);
      const now = new Date();
      state.devices = raw.map((d) => Devices.enrich(d, managedMap, now));
      state.selected.clear();

      fill.style.width = '100%';
      const dt = ((Date.now() - t0) / 1000).toFixed(1);
      status.textContent = `Loaded ${state.devices.length} devices in ${dt}s`;
      Log.info(`Loaded ${state.devices.length} devices in ${dt}s`);

      document.getElementById('filterRow').classList.remove('hidden');
      applyFilter();
      activateTab('results');
    } catch (e) {
      Log.err('Load failed:', e);
      toast('Load failed: ' + e.message, 'err');
    } finally {
      btn.disabled = false;
      setTimeout(() => box.classList.add('hidden'), 600);
      state.cancelToken = null;
    }
  }

  function readFilter() {
    return {
      search: document.getElementById('search').value.trim(),
      minInactive: parseInt(document.getElementById('minInactive').value, 10) || 0,
      minAge: parseInt(document.getElementById('minAge').value, 10) || 0,
      osFilter: document.getElementById('osFilter').value,
      joinType: document.getElementById('joinType').value,
      autopilot: document.getElementById('autopilot').value,
      managed: document.getElementById('managed').value,
      hasPrimaryUser: document.getElementById('hasPrimaryUser').value,
      enabled: document.getElementById('enabled').value,
    };
  }

  function applyFilter() {
    const f = readFilter();
    state.filtered = state.devices.filter((d) => Devices.passesFilter(d, f));

    for (const id of [...state.selected]) {
      if (!state.filtered.some((d) => d.id === id)) state.selected.delete(id);
    }
    renderResults();
  }

  function sortDevices(rows) {
    const { col, dir } = state.sort;
    const m = dir === 'asc' ? 1 : -1;

    const key = (d) => {
      switch (col) {
        case 'displayName': return (d.displayName || '').toLowerCase();
        case 'primaryUser': return (d.primaryUser || '').toLowerCase();
        case 'operatingSystem': return (d.operatingSystem || '').toLowerCase();
        case 'joinType': return d.joinType || '';
        case 'isAutopilot': return d.isAutopilot ? 1 : 0;
        case 'isManagedEffective': return d.isManagedEffective ? 1 : 0;
        case 'accountEnabled': return d.accountEnabled ? 1 : 0;
        case 'createdDateTime': return d.createdDateTime || '';
        case 'lastActive': return d.lastActive ? d.lastActive.getTime() : -1;
        case 'inactiveDays': return d.inactiveDays == null ? -1 : d.inactiveDays;
        default: return '';
      }
    };

    return rows.slice().sort((a, b) => {
      const ka = key(a);
      const kb = key(b);
      return ka > kb ? m : ka < kb ? -m : 0;
    });
  }

  function renderResults() {
    const wrap = document.getElementById('tableWrap');
    const hint = document.getElementById('tableHint');
    const summary = document.getElementById('summaryBar');
    const bar = document.getElementById('actionBar');
    const tbody = document.querySelector('#devicesTable tbody');

    document.getElementById('exportXlsxBtn').disabled = state.devices.length === 0;
    document.getElementById('exportCsvBtn').disabled = state.devices.length === 0;
    document.getElementById('selectAllFilteredBtn').disabled = state.filtered.length === 0;
    document.getElementById('selectNoneBtn').disabled = state.selected.size === 0;

    if (!state.devices.length) {
      wrap.classList.add('hidden');
      hint.classList.add('hidden');
      summary.classList.add('hidden');
      bar.classList.add('hidden');
      document.getElementById('executeBtn').disabled = true;
      document.getElementById('selCount').textContent = '0';
      return;
    }

    summary.classList.remove('hidden');
    summary.innerHTML = `
      <span><strong>${state.devices.length}</strong> devices scanned</span>
      <span><strong>${state.filtered.length}</strong> matching filter</span>
      <span><strong>${state.filtered.filter((d) => d.accountEnabled !== false).length}</strong> enabled in view</span>
      <span><strong>${state.selected.size}</strong> selected</span>`;

    wrap.classList.remove('hidden');
    hint.classList.remove('hidden');
    bar.classList.toggle('hidden', state.selected.size === 0);

    const rows = sortDevices(state.filtered).slice(0, 1000);
    hint.textContent = `${state.filtered.length} match${state.filtered.length === 1 ? '' : 'es'}` +
      (state.filtered.length > rows.length ? ` (showing first ${rows.length})` : '');

    tbody.innerHTML = rows.map((d) => `
      <tr data-id="${escAttr(d.id)}">
        <td><input type="checkbox" data-act="sel" ${state.selected.has(d.id) ? 'checked' : ''}/></td>
        <td class="col-name" title="${escAttr(d.displayName || '')}"><strong>${escHtml(d.displayName || '(no name)')}</strong><br><span class="muted small">${escHtml(d.deviceId || '')}</span></td>
        <td class="col-user" title="${escAttr(d.primaryUser || '')}">${d.primaryUser ? `<code>${escHtml(d.primaryUser)}</code>` : '<span class="muted">Unknown</span>'}</td>
        <td title="${escAttr((d.operatingSystem || '') + (d.operatingSystemVersion ? ' ' + d.operatingSystemVersion : ''))}">${escHtml(d.operatingSystem || '')}${d.operatingSystemVersion ? `<br><span class="muted small">${escHtml(d.operatingSystemVersion)}</span>` : ''}</td>
        <td>${escHtml(d.joinType)}</td>
        <td>${d.isAutopilot ? 'Yes' : '<span class="muted">No</span>'}</td>
        <td>${d.isManagedEffective ? 'Yes' : '<span class="muted">No</span>'}</td>
        <td>${d.accountEnabled ? 'Yes' : '<span class="muted">No</span>'}</td>
        <td>${escHtml(Devices.fmtDate(d.createdDateTime))}<br><span class="muted small">${d.deviceAgeDays == null ? '' : d.deviceAgeDays + ' days'}</span></td>
        <td>${d.lastActive ? escHtml(Devices.fmtDate(d.lastActive)) : '<span class="muted">Never</span>'}</td>
        <td>${d.inactiveDays == null ? '' : d.inactiveDays + ' days'}</td>
      </tr>
    `).join('');

    document.querySelectorAll('#devicesTable thead th').forEach((th) => {
      th.classList.remove('sort-asc', 'sort-desc', 'sortable-col');
      if (th.dataset.sort) {
        th.classList.add('sortable-col');
        if (th.dataset.sort === state.sort.col) {
          th.classList.add(state.sort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
        }
      }
    });

    document.getElementById('selCount').textContent = String(state.selected.size);
    document.getElementById('executeBtn').disabled = state.selected.size === 0;
    document.getElementById('selectAll').checked = rows.length > 0 && rows.every((d) => state.selected.has(d.id));
  }

  function onTableChange(e) {
    const cb = e.target.closest('input[data-act="sel"]');
    if (!cb) return;
    const id = cb.closest('tr').dataset.id;
    if (cb.checked) state.selected.add(id);
    else state.selected.delete(id);
    renderResults();
  }

  function onSortClick(e) {
    const th = e.target.closest('th[data-sort]');
    if (!th) return;
    const col = th.dataset.sort;
    if (state.sort.col === col) state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
    else {
      state.sort.col = col;
      state.sort.dir = col === 'inactiveDays' ? 'desc' : 'asc';
    }
    renderResults();
  }

  function onSelectAll(e) {
    const checked = e.target.checked;
    const rows = sortDevices(state.filtered).slice(0, 1000);
    for (const d of rows) {
      if (checked) state.selected.add(d.id);
      else state.selected.delete(d.id);
    }
    renderResults();
  }

  function onSelectAllFiltered() {
    const rows = sortDevices(state.filtered).slice(0, 1000);
    for (const d of rows) state.selected.add(d.id);
    renderResults();
  }

  function onSelectNone() {
    state.selected.clear();
    renderResults();
  }

  function selectedDevices() {
    return state.devices.filter((d) => state.selected.has(d.id));
  }

  function previewList(rows, max = 20) {
    const items = rows.slice(0, max).map((d) =>
      `<li>${escHtml(d.displayName || d.deviceId)} <span class="muted small">${escHtml(d.primaryUser || d.deviceId || '')}</span></li>`
    ).join('');
    return `<ul>${items}</ul>` + (rows.length > max ? `<p class="muted small">+ ${rows.length - max} more</p>` : '');
  }

  async function executeSelected() {
    const mode = document.querySelector('input[name=cleanupMode]:checked').value;
    const rows = selectedDevices();
    if (!rows.length) return;

    const blocked = rows.filter((d) => d.isAutopilot);
    const actionable = rows.filter((d) => !d.isAutopilot);

    if (!actionable.length) {
      await confirmModal(
        'Cannot perform action',
        `<p>All selected devices are Autopilot-registered and cannot be ${mode}d via Graph in this tool.</p>`,
        'OK'
      );
      return;
    }

    const blockedNote = blocked.length
      ? `<p class="notice notice-warn"><strong>${blocked.length} Autopilot device(s) will be skipped</strong> — these object types cannot be ${mode}d via this workflow.</p>`
      : '';

    const ok = await confirmModal(
      `${mode === 'delete' ? 'Delete' : 'Disable'} ${actionable.length} device(s)?`,
      `<p>You are about to <strong>${mode}</strong> ${actionable.length} device(s):</p>${previewList(actionable)}` +
      blockedNote +
      (mode === 'delete' ? '<p class="notice notice-warn">Deleting devices is usually permanent.</p>' : ''),
      mode === 'delete' ? 'Delete' : 'Disable'
    );
    if (!ok) return;

    const resultsEl = document.getElementById('actionResults');
    resultsEl.classList.remove('hidden');
    resultsEl.innerHTML = '';

    const extAttr = document.getElementById('extAttrSelect').value;
    const ext = mode === 'disable' && extAttr ? { [extAttr]: buildDisableStampValue() } : null;
    const concurrency = Math.max(1, Math.min(20, +document.getElementById('concurrency').value || 4));

    let done = 0;
    const total = actionable.length;

    await Concurrency.pmap(actionable, async (d) => {
      const li = document.createElement('li');
      li.innerHTML = `<span class="ok-mark">...</span> ${escHtml(d.displayName || d.deviceId)}`;
      resultsEl.appendChild(li);
      try {
        if (mode === 'delete') {
          await Graph.deleteDevice(d.id);
          li.innerHTML = `<span class="ok-mark">&#10003;</span> Deleted - ${escHtml(d.displayName || d.deviceId)}`;
        } else {
          await Graph.disableDevice(d.id, ext);
          li.innerHTML = `<span class="ok-mark">&#10003;</span> Disabled - ${escHtml(d.displayName || d.deviceId)}${ext ? ` <span class="muted">(stamped ${escHtml(extAttr)})</span>` : ''}`;
        }
      } catch (e) {
        li.innerHTML = `<span class="err-mark">&#10007;</span> ${escHtml(d.displayName || d.deviceId)} - ${escHtml(e.message)}`;
        Log.err(`${mode} failed for ${d.id}:`, e);
      } finally {
        done++;
      }
    }, { concurrency });

    toast(`${mode} done: ${done}/${total}`, 'ok');
    state.selected.clear();
    await scanDevices();
  }

  async function findExpired() {
    const attr = document.getElementById('extAttrSelect').value;
    if (!attr) {
      alert('Choose an extension attribute first in Scan settings.');
      return;
    }
    const grace = +document.getElementById('gracePeriod').value || 30;

    let devices = state.devices;
    if (!devices.length) {
      const raw = await Graph.listAllDevices(() => {});
      const managedRes = await Graph.tryListManagedDevices(() => {});
      const managedMap = Devices.buildManagedMap(managedRes.list);
      devices = raw.map((d) => Devices.enrich(d, managedMap, new Date()));
      state.devices = devices;
      applyFilter();
    }

    const cutoff = Date.now() - grace * 86400000;
    state.expired = devices
      .filter((d) => d.accountEnabled === false)
      .map((d) => {
        const stamp = parseDisableStamp((d.extensionAttributes || {})[attr]);
        if (!stamp) return null;
        const days = Math.floor((Date.now() - stamp.getTime()) / 86400000);
        return stamp.getTime() <= cutoff ? { device: d, disabledAt: stamp, daysDisabled: days } : null;
      })
      .filter(Boolean)
      .sort((a, b) => a.disabledAt - b.disabledAt);

    renderExpired();
  }

  function renderExpired() {
    const root = document.getElementById('expiredContainer');
    const delBtn = document.getElementById('deleteExpiredBtn');
    root.innerHTML = '';

    if (!state.expired.length) {
      root.innerHTML = '<p class="table-hint">No disabled devices past the grace period found.</p>';
      delBtn.disabled = true;
      return;
    }

    delBtn.disabled = false;
    const box = document.createElement('div');
    box.className = 'table-wrap';
    box.innerHTML = `<table class="data-table"><thead><tr><th>Display name</th><th>OS</th><th>Autopilot</th><th>Disabled at</th><th>Days disabled</th><th>Object id</th></tr></thead><tbody></tbody></table>`;
    const tbody = box.querySelector('tbody');
    for (const e of state.expired) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${escHtml(e.device.displayName || '')}</td>
        <td>${escHtml([e.device.operatingSystem, e.device.operatingSystemVersion].filter(Boolean).join(' '))}</td>
        <td>${e.device.isAutopilot ? 'Yes' : '<span class="muted">No</span>'}</td>
        <td>${escHtml(Devices.fmtDate(e.disabledAt))}</td>
        <td>${e.daysDisabled}</td>
        <td class="url-cell" title="${escAttr(e.device.id)}">${escHtml(e.device.id || '')}</td>`;
      tbody.appendChild(tr);
    }
    root.appendChild(box);
  }

  async function deleteExpired() {
    if (!state.expired.length) return;

    const blocked = state.expired.filter((e) => e.device.isAutopilot);
    const actionable = state.expired.filter((e) => !e.device.isAutopilot);

    if (!actionable.length) {
      alert('All expired devices are Autopilot-registered and cannot be deleted via this workflow.');
      return;
    }

    const message = `PERMANENTLY DELETE ${actionable.length} expired disabled device(s)?` +
      (blocked.length ? `\n\n${blocked.length} Autopilot device(s) will be skipped.` : '');
    if (!confirm(message)) return;

    const concurrency = Math.max(1, Math.min(20, +document.getElementById('concurrency').value || 4));
    const ul = document.createElement('ul');
    ul.className = 'action-result-list';
    document.getElementById('expiredContainer').appendChild(ul);

    const survivors = [];
    await Concurrency.pmap(actionable, async (e) => {
      const li = document.createElement('li');
      li.innerHTML = `<span class="muted">...</span> ${escHtml(e.device.displayName || e.device.id)}`;
      ul.appendChild(li);
      try {
        await Graph.deleteDevice(e.device.id);
        li.innerHTML = `<span class="ok-mark">DEL</span> Deleted: ${escHtml(e.device.displayName || e.device.id)}`;
      } catch (err) {
        li.innerHTML = `<span class="err-mark">ERR</span> ${escHtml(e.device.displayName || e.device.id)}: ${escHtml(err.message)}`;
        survivors.push(e);
      }
    }, { concurrency });

    state.expired = survivors.concat(blocked);
    renderExpired();
  }

  function onExportXlsx() {
    const name = `inactive-devices-${new Date().toISOString().substring(0, 10)}.xlsx`;
    return Export.exportXlsx(state.filtered, name);
  }

  function onExportCsv() {
    const name = `inactive-devices-${new Date().toISOString().substring(0, 10)}.csv`;
    return Export.exportCsv(state.filtered, name);
  }

  function escHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    }[c]));
  }

  function escAttr(s) { return escHtml(s); }

  async function boot() {
    initTheme();

    document.getElementById('darkToggle').addEventListener('click', toggleTheme);
    document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => activateTab(t.dataset.tab)));

    document.getElementById('signInBtn').addEventListener('click', onSignIn);
    document.getElementById('signOutBtn').addEventListener('click', onSignOut);
    document.getElementById('scanBtn').addEventListener('click', scanDevices);

    ['search', 'minInactive', 'minAge', 'osFilter', 'joinType', 'autopilot', 'managed', 'hasPrimaryUser', 'enabled'].forEach((id) => {
      document.getElementById(id).addEventListener('input', applyFilter);
      document.getElementById(id).addEventListener('change', applyFilter);
    });

    ['extAttrSelect', 'gracePeriod', 'concurrency'].forEach((id) => {
      document.getElementById(id).addEventListener('change', savePrefs);
    });
    document.querySelectorAll('input[name=cleanupMode]').forEach((r) => {
      r.addEventListener('change', () => { savePrefs(); refreshActionLabel(); });
    });

    document.querySelector('#devicesTable thead').addEventListener('click', onSortClick);
    document.querySelector('#devicesTable tbody').addEventListener('change', onTableChange);
    document.getElementById('selectAll').addEventListener('change', onSelectAll);

    document.getElementById('selectAllFilteredBtn').addEventListener('click', onSelectAllFiltered);
    document.getElementById('selectNoneBtn').addEventListener('click', onSelectNone);

    document.getElementById('executeBtn').addEventListener('click', executeSelected);
    document.getElementById('exportXlsxBtn').addEventListener('click', onExportXlsx);
    document.getElementById('exportCsvBtn').addEventListener('click', onExportCsv);

    document.getElementById('findExpiredBtn').addEventListener('click', findExpired);
    document.getElementById('deleteExpiredBtn').addEventListener('click', deleteExpired);

    document.getElementById('clearLogBtn').addEventListener('click', () => Log.clear());
    document.getElementById('verboseLog').addEventListener('change', (e) => Log.setVerbose(e.target.checked));

    try {
      await Auth.init();
      Log.info('InactiveDevices ready.');
      setSignedInUI(Auth.getAccount());
      loadPrefs();
      refreshActionLabel();
      renderResults();
      renderExpired();
    } catch (e) {
      Log.err('Boot failed:', e);
      toast('Initialization failed: ' + e.message, 'err');
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();