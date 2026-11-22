// InactiveDevices UI.
(() => {
  const { esc, toast, confirm, busy, setUser, h } = Toolkit;
  const DAY = 86400000;
  const MAX_ROWS = 1000;

  // Graph documents no narrower delegated permission than Directory.AccessAsUser.All
  // for PATCH or DELETE on /devices, so that is the write set.
  window.Auth = ToolkitAuth.create({
    scopes: {
      read: ['User.Read', 'Device.Read.All', 'DeviceManagementManagedDevices.Read.All'],
      write: ['Directory.AccessAsUser.All'],
    },
  });

  const state = {
    devices: [],
    filtered: [],
    selected: new Set(),
    sort: { col: 'inactiveDays', dir: 'desc' },
    managedMap: null,
    expired: [],
    expiredCtx: null,
    cancelToken: null,
  };
  const $ = (id) => document.getElementById(id);

  // ---------- preferences (nothing destructive is remembered) ----------
  const PREFS_KEY = 'inactivedevices-prefs';
  function loadPrefs() {
    let p = {};
    try { p = JSON.parse(localStorage.getItem(PREFS_KEY) || '{}') || {}; } catch { p = {}; }
    if (/^extensionAttribute([1-9]|1[0-5])$/.test(p.extAttr || '')) $('extAttrSelect').value = p.extAttr;
    if (Devices.parseGraceDays(p.gracePeriod) != null) $('gracePeriod').value = p.gracePeriod;
    if (p.concurrency >= 1 && p.concurrency <= 20) $('concurrency').value = p.concurrency;
  }
  function savePrefs() {
    const p = { extAttr: $('extAttrSelect').value, gracePeriod: Devices.parseGraceDays($('gracePeriod').value) || 30, concurrency: concurrency() };
    try { localStorage.setItem(PREFS_KEY, JSON.stringify(p)); } catch { /* storage blocked */ }
  }
  function concurrency() { return Math.max(1, Math.min(20, parseInt($('concurrency').value, 10) || 4)); }

  // ---------- auth ----------
  function refreshAuthUI() {
    const signedIn = !!Auth.getAccount();
    setUser(Auth.userInfo(onSignOut));
    $('connectCard').hidden = signedIn;
    $('scanCard').hidden = !signedIn;
    $('findExpiredBtn').disabled = !signedIn;
  }
  async function onSignIn() {
    try { await Auth.signIn('read'); }
    catch (e) { Log.err('Sign-in failed:', e); toast(`Sign-in failed: ${e.message}`, 'err'); }
  }
  async function onSignOut() {
    state.devices = []; state.filtered = []; state.selected.clear(); state.expired = [];
    await Auth.signOut();
  }

  // ---------- load ----------
  async function loadDevices({ switchTab = true } = {}) {
    savePrefs();
    const box = $('progressBox');
    box.hidden = false;
    $('progressLabel').textContent = 'Loading devices';
    $('progressCount').textContent = '0 loaded';
    state.cancelToken = new Concurrency.CancelToken();
    const signal = state.cancelToken.signal;
    const t0 = Date.now();
    await busy($('scanBtn'), async () => {
      try {
        const raw = await Graph.listAllDevices((n, page) => { $('progressCount').textContent = `${n} devices, page ${page}`; }, signal);
        $('progressLabel').textContent = 'Loading Intune data';
        const managedRes = await Graph.tryListManagedDevices((n, page) => { $('progressCount').textContent = `${raw.length} devices, ${n} Intune devices, page ${page}`; }, signal);
        $('managedNotice').hidden = !managedRes.unavailable;
        $('managedNoticeText').textContent = managedRes.message;
        state.managedMap = Devices.buildManagedMap(managedRes.list);
        const now = new Date();
        state.devices = raw.map((d) => Devices.enrich(d, state.managedMap, now));
        state.selected.clear();
        const dt = ((Date.now() - t0) / 1000).toFixed(1);
        const msg = `${state.devices.length} devices loaded in ${dt}s. Nothing changes until you pick an action and confirm it.`;
        $('loadStatus').textContent = msg;
        $('resultsStatus').textContent = msg;
        Log.info(`Loaded ${state.devices.length} devices in ${dt}s`);
        $('filterRow').hidden = false;
        applyFilter();
        if (switchTab) activateTab('results');
      } catch (e) {
        if (e.message === 'Cancelled') { toast('Loading cancelled', 'warn'); return; }
        Log.err('Load failed:', e);
        toast(`Loading failed: ${e.message}`, 'err');
      } finally {
        box.hidden = true;
        state.cancelToken = null;
      }
    });
  }

  function readFilter() {
    return {
      search: $('search').value.trim(),
      minInactive: parseInt($('minInactive').value, 10) || 0,
      minAge: parseInt($('minAge').value, 10) || 0,
      osFilter: $('osFilter').value,
      joinType: $('joinType').value,
      autopilot: $('autopilot').value,
      managed: $('managed').value,
      hasPrimaryUser: $('hasPrimaryUser').value,
      enabled: $('enabled').value,
    };
  }
  function applyFilter() {
    const f = readFilter();
    state.filtered = state.devices.filter((d) => Devices.passesFilter(d, f));
    // Only what is visible stays selected, so an action never hits hidden rows.
    Devices.trimSelection(state.selected, visibleIds());
    renderResults();
  }
  function visibleRows() { return sortDevices(state.filtered).slice(0, MAX_ROWS); }
  function visibleIds() { return new Set(visibleRows().map((d) => d.id)); }

  // ---------- table ----------
  function sortKey(d, col) {
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
  }
  function sortDevices(rows) {
    const { col, dir } = state.sort;
    const m = dir === 'asc' ? 1 : -1;
    return rows.slice().sort((a, b) => {
      const ka = sortKey(a, col), kb = sortKey(b, col);
      return ka > kb ? m : ka < kb ? -m : 0;
    });
  }

  function renderResults() {
    const has = state.devices.length > 0;
    $('statTiles').hidden = !has;
    $('tableWrap').hidden = !has;
    $('tableHint').hidden = !has;
    $('actionBar').hidden = !has;
    const tbody = document.querySelector('#devicesTable tbody');
    if (!has) { tbody.replaceChildren(); updateSelection(); return; }

    $('statTotal').textContent = String(state.devices.length);
    $('statMatching').textContent = String(state.filtered.length);
    $('statEnabled').textContent = String(state.filtered.filter((d) => d.accountEnabled !== false).length);

    const rows = visibleRows();
    $('tableHint').textContent = `${state.filtered.length} match${state.filtered.length === 1 ? '' : 'es'}` +
      (state.filtered.length > rows.length ? `, showing the first ${rows.length}` : '');
    tbody.innerHTML = rows.map((d) => `
      <tr data-id="${esc(d.id)}">
        <td><input type="checkbox" data-act="sel" aria-label="Select ${esc(d.displayName || d.deviceId)}" ${state.selected.has(d.id) ? 'checked' : ''}/></td>
        <td class="col-name" title="${esc(d.displayName || '')}"><strong>${esc(d.displayName || '(no name)')}</strong><span class="id-line muted small">${esc(d.deviceId || '')}</span></td>
        <td class="col-user" title="${esc(d.primaryUser || '')}">${d.primaryUser ? esc(d.primaryUser) : '<span class="muted">Unknown</span>'}</td>
        <td class="col-os" title="${esc(Devices.osLabel(d))}">${esc(d.operatingSystem || '')}${d.operatingSystemVersion ? `<br><span class="muted small">${esc(d.operatingSystemVersion)}</span>` : ''}</td>
        <td>${esc(d.joinType)}</td>
        <td>${d.isAutopilot ? '<span class="badge info">Yes</span>' : '<span class="muted">No</span>'}</td>
        <td>${d.isManagedEffective ? 'Yes' : '<span class="muted">No</span>'}</td>
        <td>${d.accountEnabled ? '<span class="badge ok">Yes</span>' : '<span class="badge neutral">No</span>'}</td>
        <td>${esc(Devices.fmtDate(d.createdDateTime))}<br><span class="muted small">${d.deviceAgeDays == null ? '' : `${d.deviceAgeDays} days`}</span></td>
        <td>${d.lastActive ? esc(Devices.fmtDate(d.lastActive)) : '<span class="muted">Never</span>'}</td>
        <td class="num">${d.inactiveDays == null ? '' : `${d.inactiveDays} d`}</td>
      </tr>`).join('');

    document.querySelectorAll('#devicesTable thead th').forEach((th) => {
      th.classList.remove('sort-asc', 'sort-desc');
      if (th.dataset.sort === state.sort.col) th.classList.add(state.sort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
    });
    $('selectAll').checked = rows.length > 0 && rows.every((d) => state.selected.has(d.id));
    updateSelection();
  }

  function updateSelection() {
    const n = state.selected.size;
    $('selCount').textContent = String(n);
    $('statSelected').textContent = String(n);
    ['disableBtn', 'enableBtn', 'deleteBtn', 'selectNoneBtn'].forEach((id) => { $(id).disabled = n === 0; });
    $('selectAllFilteredBtn').disabled = state.filtered.length === 0;
    $('exportXlsxBtn').disabled = state.filtered.length === 0;
    $('exportCsvBtn').disabled = state.filtered.length === 0;
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
    renderResults();
  }
  function onSelectAll(e) {
    for (const d of visibleRows()) { if (e.target.checked) state.selected.add(d.id); else state.selected.delete(d.id); }
    renderResults();
  }

  // ---------- bulk actions ----------
  function startProgress(label, total) {
    $('actionProgress').hidden = false;
    $('actionLabel').textContent = label;
    $('actionCount').textContent = `0/${total}`;
    $('actionEta').textContent = '';
    $('actionFill').style.width = '0%';
    const t0 = Date.now();
    return (done) => {
      $('actionCount').textContent = `${done}/${total}`;
      $('actionEta').textContent = Concurrency.eta(t0, done, total);
      $('actionFill').style.width = `${Math.round((100 * done) / total)}%`;
    };
  }

  function resultLine(list, d) {
    const name = d.displayName || d.deviceId || d.id;
    const li = h('li', null, [h('span.warn-mark', { text: '...' }), ` ${name}`]);
    list.appendChild(li);
    return {
      ok: (text) => li.replaceChildren(h('span.ok-mark', { text: 'OK' }), ` ${text}`),
      skip: (text) => li.replaceChildren(h('span.warn-mark', { text: 'SKIP' }), ` ${text}`),
      fail: (text) => li.replaceChildren(h('span.err-mark', { text: 'ERR' }), ` ${text}`),
    };
  }

  const VERBS = {
    disable: { title: 'Disable', done: 'Disabled', busy: 'Disabling', btn: 'disableBtn' },
    enable: { title: 'Enable', done: 'Enabled', busy: 'Enabling', btn: 'enableBtn' },
    delete: { title: 'Delete', done: 'Deleted', busy: 'Deleting', btn: 'deleteBtn' },
  };

  async function bulkAction(kind) {
    const v = VERBS[kind];
    const visible = visibleIds();
    Devices.trimSelection(state.selected, visible);
    const rows = state.filtered.filter((d) => state.selected.has(d.id));
    if (!rows.length) { renderResults(); return; }
    const { actionable, skipped } = Devices.splitActionable(rows, kind);
    if (!actionable.length) {
      toast(`All ${rows.length} selected devices are Autopilot devices. They are never ${kind}d by this tool.`, 'warn');
      return;
    }
    const extAttr = kind === 'disable' ? $('extAttrSelect').value : '';
    const n = actionable.length;
    const bodies = {
      disable: `These devices can no longer be used to sign in to Entra ID until they are enabled again.${extAttr ? ` The disable timestamp goes into ${extAttr}.` : ''}`,
      enable: 'These devices can sign in to Entra ID again.',
      delete: 'These devices are deleted from Entra ID. This cannot be undone; a device that is still in use has to register again.',
    };
    const ok = await confirm({
      title: `${v.title} ${n} device${n === 1 ? '' : 's'}?`,
      body: bodies[kind] + (skipped.length ? ` ${skipped.length} selected Autopilot device${skipped.length === 1 ? ' is' : 's are'} skipped.` : ''),
      items: actionable.map(Devices.describe),
      okLabel: `${v.title} ${n}`,
      danger: kind === 'delete',
      wide: true,
    });
    if (!ok) return;
    // Consent for the write scope before the first device, by popup, so this page stays put.
    try { await Auth.getWriteToken(); }
    catch (e) { toast(`No write permission: ${e.message}`, 'err'); return; }

    const list = $('actionResults');
    list.hidden = false;
    list.replaceChildren();
    const tick = startProgress(v.busy, n);
    const ext = extAttr ? { [extAttr]: Devices.buildDisableStampValue() } : null;
    let done = 0, okCount = 0;
    await busy($(v.btn), () => Concurrency.pmap(actionable, async (d) => {
      const line = resultLine(list, d);
      const name = d.displayName || d.deviceId || d.id;
      try {
        if (kind === 'delete') await Graph.deleteDevice(d.id);
        else if (kind === 'enable') await Graph.enableDevice(d.id);
        else await Graph.disableDevice(d.id, ext);
        okCount++;
        line.ok(`${v.done} ${name}${ext ? `, stamped ${extAttr}` : ''}`);
      } catch (e) {
        line.fail(`${name}: ${e.message}`);
        Log.err(`${kind} failed for ${d.id}:`, e);
        throw e;
      } finally {
        tick(++done);
      }
    }, { concurrency: concurrency() }));
    $('actionProgress').hidden = true;
    const failed = n - okCount;
    toast(`${v.done} ${okCount} of ${n}${failed ? `, ${failed} failed (see the list and the debug log)` : ''}`, failed ? 'warn' : 'ok');
    Log.info(`${kind}: ${okCount}/${n} succeeded`);
    state.selected.clear();
    await loadDevices({ switchTab: false });
  }

  // ---------- past grace period ----------
  function clearExpired(message) {
    state.expired = [];
    state.expiredCtx = null;
    $('expiredStatus').textContent = message || '';
    renderExpired(false);
  }

  async function findExpired() {
    const attr = $('extAttrSelect').value;
    if (!attr) { toast('Pick an extension attribute on the Scan tab first. Without it there is no timestamp to look at.', 'warn'); return; }
    const grace = Devices.parseGraceDays($('gracePeriod').value);
    if (grace == null) { toast('Grace period must be a whole number of days, 1 to 3650.', 'err'); return; }
    savePrefs();
    $('expiredResults').hidden = true;
    await busy($('findExpiredBtn'), async () => {
      try {
        $('expiredStatus').textContent = 'Loading disabled devices';
        const raw = await Graph.listDisabledDevices((n) => { $('expiredStatus').textContent = `${n} disabled devices loaded`; });
        if (!state.managedMap) state.managedMap = Devices.buildManagedMap((await Graph.tryListManagedDevices()).list);
        const now = new Date();
        const devices = raw.map((d) => Devices.enrich(d, state.managedMap, now));
        state.expired = Devices.findExpiredDisabled(devices, attr, grace, now);
        state.expiredCtx = { attr, grace };
        $('expiredStatus').textContent = `${raw.length} disabled devices checked, ${state.expired.length} with a ${attr} timestamp older than ${grace} days.`;
        Log.info(`Grace period search: ${state.expired.length} of ${raw.length} disabled devices expired (${attr}, ${grace} days)`);
        renderExpired(true);
      } catch (e) {
        Log.err('Grace period search failed:', e);
        $('expiredStatus').textContent = '';
        toast(`Search failed: ${e.message}`, 'err');
      }
    });
  }

  function renderExpired(searched) {
    const has = state.expired.length > 0;
    $('expiredWrap').hidden = !has;
    $('expiredEmpty').hidden = has || !searched;
    $('deleteExpiredBtn').disabled = !state.expired.some((e) => !e.device.isAutopilot);
    document.querySelector('#expiredTable tbody').innerHTML = state.expired.map((e) => `<tr>
      <td>${esc(e.device.displayName || '(no name)')}</td>
      <td>${esc(Devices.osLabel(e.device))}</td>
      <td>${e.device.isAutopilot ? '<span class="badge info">Yes, skipped</span>' : '<span class="muted">No</span>'}</td>
      <td class="nowrap">${esc(Devices.fmtDate(e.disabledAt))}</td>
      <td class="num">${e.daysDisabled}</td>
      <td class="url-cell mono" title="${esc(e.device.id)}">${esc(e.device.id || '')}</td>
    </tr>`).join('');
  }

  async function deleteExpired() {
    const ctx = state.expiredCtx;
    if (!ctx) return;
    const { actionable, skipped } = Devices.splitActionable(state.expired.map((e) => e.device), 'delete');
    const rows = state.expired.filter((e) => actionable.includes(e.device));
    if (!rows.length) return;
    const n = rows.length;
    const ok = await confirm({
      title: `Delete ${n} expired device${n === 1 ? '' : 's'}?`,
      body: `These disabled devices are deleted from Entra ID. Each one is read again first; a device that was enabled again or whose ${ctx.attr} changed is skipped. Deleting cannot be undone.${skipped.length ? ` ${skipped.length} Autopilot device${skipped.length === 1 ? ' is' : 's are'} skipped.` : ''}`,
      items: rows.map((e) => `${Devices.describe(e.device)}  |  disabled ${Devices.fmtDate(e.disabledAt)}`),
      okLabel: `Delete ${n}`,
      danger: true,
      wide: true,
    });
    if (!ok) return;
    try { await Auth.getWriteToken(); }
    catch (e) { toast(`No write permission: ${e.message}`, 'err'); return; }

    const list = $('expiredResults');
    list.hidden = false;
    list.replaceChildren();
    const deleted = new Set();
    let skippedNow = 0, failed = 0;
    await busy($('deleteExpiredBtn'), () => Concurrency.pmap(rows, async (e) => {
      const d = e.device;
      const name = d.displayName || d.id;
      const line = resultLine(list, d);
      try {
        const fresh = await Graph.getDevice(d.id);
        const check = Devices.recheckExpired(fresh, ctx.attr, e.stampValue, Date.now() - ctx.grace * DAY);
        if (!check.ok) {
          skippedNow++;
          line.skip(`${name}: ${check.reason}`);
          Log.warn(`Grace delete skipped ${d.id}: ${check.reason}`);
          return;
        }
        await Graph.deleteDevice(d.id);
        deleted.add(d.id);
        line.ok(`Deleted ${name}`);
      } catch (err) {
        failed++;
        line.fail(`${name}: ${err.message}`);
        Log.err(`Grace delete failed for ${d.id}:`, err);
      }
    }, { concurrency: concurrency() }));
    toast(`Deleted ${deleted.size} of ${n}${skippedNow ? `, ${skippedNow} skipped` : ''}${failed ? `, ${failed} failed` : ''}`, skippedNow || failed ? 'warn' : 'ok');
    state.expired = state.expired.filter((e) => !deleted.has(e.device.id));
    renderExpired(true);
    if (deleted.size && state.devices.length) {
      state.devices = state.devices.filter((d) => !deleted.has(d.id));
      applyFilter();
    }
  }

  // ---------- export ----------
  const EXPORT_COLUMNS = [
    { key: 'displayName', label: 'Display name' },
    { key: 'deviceId', label: 'Device id' },
    { key: 'primaryUser', label: 'Primary user' },
    { key: 'operatingSystem', label: 'OS' },
    { key: 'operatingSystemVersion', label: 'OS version' },
    { key: 'joinType', label: 'Join type' },
    { key: 'isAutopilot', label: 'Autopilot' },
    { key: 'isManagedEffective', label: 'Managed' },
    { key: 'accountEnabled', label: 'Enabled' },
    { key: 'createdDateTime', label: 'Created', format: (v) => Devices.fmtDate(v) },
    { key: 'lastSignIn', label: 'Last sign-in (Entra)', format: (v) => Devices.fmtDate(v) },
    { key: 'lastSync', label: 'Last sync (Intune)', format: (v) => Devices.fmtDate(v) },
    { key: 'lastActive', label: 'Last active', format: (v) => Devices.fmtDate(v) },
    { key: 'inactiveDays', label: 'Inactive (days)' },
    { key: 'deviceAgeDays', label: 'Age (days)' },
    { key: 'id', label: 'Object id' },
  ];
  function exportRows() { return sortDevices(state.filtered); }

  // ---------- boot ----------
  let activateTab = () => {};

  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    activateTab = Toolkit.tabs('#tabs');
    loadPrefs();
    $('signInBtn').addEventListener('click', onSignIn);
    $('scanBtn').addEventListener('click', () => loadDevices());
    $('cancelBtn').addEventListener('click', () => { if (state.cancelToken) state.cancelToken.cancel(); });
    ['extAttrSelect', 'gracePeriod', 'concurrency'].forEach((id) => $(id).addEventListener('change', () => {
      savePrefs();
      if (id !== 'concurrency' && state.expiredCtx) clearExpired('Settings changed, search again.');
    }));
    ['search', 'minInactive', 'minAge', 'osFilter', 'joinType', 'autopilot', 'managed', 'hasPrimaryUser', 'enabled']
      .forEach((id) => $(id).addEventListener('input', applyFilter));
    document.querySelector('#devicesTable thead').addEventListener('click', onSortClick);
    document.querySelector('#devicesTable tbody').addEventListener('change', onTableChange);
    $('selectAll').addEventListener('change', onSelectAll);
    $('selectAllFilteredBtn').addEventListener('click', () => { for (const d of visibleRows()) state.selected.add(d.id); renderResults(); });
    $('selectNoneBtn').addEventListener('click', () => { state.selected.clear(); renderResults(); });
    $('disableBtn').addEventListener('click', () => bulkAction('disable'));
    $('enableBtn').addEventListener('click', () => bulkAction('enable'));
    $('deleteBtn').addEventListener('click', () => bulkAction('delete'));
    $('findExpiredBtn').addEventListener('click', findExpired);
    $('deleteExpiredBtn').addEventListener('click', deleteExpired);
    $('exportXlsxBtn').addEventListener('click', () => busy($('exportXlsxBtn'), () => Export.xlsx([{ name: 'Devices', rows: exportRows(), columns: EXPORT_COLUMNS }], `inactive-devices-${Export.stamp()}.xlsx`).catch((e) => toast(e.message, 'err'))));
    $('exportCsvBtn').addEventListener('click', () => Export.csv(exportRows(), { columns: EXPORT_COLUMNS, filename: `inactive-devices-${Export.stamp()}.csv` }));

    try {
      await Auth.init();
      refreshAuthUI();
      Log.info('InactiveDevices ready.');
    } catch (e) {
      Log.err('Start-up failed:', e);
      toast(`Start-up failed: ${e.message}`, 'err');
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
