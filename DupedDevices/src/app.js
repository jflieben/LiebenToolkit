// DupedDevices UI.
(() => {
  const { esc, toast, confirm, busy, setUser, h } = Toolkit;
  const DAY = 86400000;

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
    groups: [],
    visible: [],
    selected: new Set(),
    expired: [],
    expiredCtx: null,
    cancelToken: null,
  };
  const $ = (id) => document.getElementById(id);

  // ---------- preferences (nothing destructive is remembered) ----------
  const PREFS_KEY = 'duped-prefs';
  const PREF_FIELDS = ['extAttrSelect', 'gracePeriod', 'concurrency'];
  function loadPrefs() {
    let p = {};
    try { p = JSON.parse(localStorage.getItem(PREFS_KEY) || '{}') || {}; } catch { p = {}; }
    if (/^extensionAttribute([1-9]|1[0-5])$/.test(p.extAttr || '')) $('extAttrSelect').value = p.extAttr;
    if (Devices.parseGraceDays(p.gracePeriod) != null) $('gracePeriod').value = p.gracePeriod;
    if (p.concurrency >= 1 && p.concurrency <= 20) $('concurrency').value = p.concurrency;
    $('includeDisabled').checked = p.includeDisabled === true;
    $('hideMixedAutopilot').checked = p.hideMixedAutopilot !== false;
  }
  function savePrefs() {
    const p = {
      extAttr: $('extAttrSelect').value,
      gracePeriod: Devices.parseGraceDays($('gracePeriod').value) || 30,
      concurrency: concurrency(),
      includeDisabled: $('includeDisabled').checked,
      hideMixedAutopilot: $('hideMixedAutopilot').checked,
    };
    try { localStorage.setItem(PREFS_KEY, JSON.stringify(p)); } catch { /* storage blocked */ }
  }
  function concurrency() { return Math.max(1, Math.min(20, parseInt($('concurrency').value, 10) || 6)); }

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
    state.devices = []; state.groups = []; state.visible = []; state.selected.clear(); state.expired = [];
    await Auth.signOut();
  }

  // ---------- scan ----------
  async function scan() {
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
        const devices = await Graph.listAllDevices((n, page) => { $('progressCount').textContent = `${n} devices, page ${page}`; }, signal);
        $('progressLabel').textContent = 'Loading Intune data';
        const managedRes = await Graph.tryListManagedDevices((n, page) => { $('progressCount').textContent = `${devices.length} devices, ${n} Intune devices, page ${page}`; }, signal);
        $('managedNotice').hidden = !managedRes.unavailable;
        $('managedNoticeText').textContent = managedRes.message;

        const managedMap = Devices.buildManagedMap(managedRes.list);
        for (const d of devices) {
          const managed = managedMap.get(Devices.keyForDevice(d)) || null;
          d.isAutopilot = Devices.isAutopilot(d, managed);
          d.primaryUser = (managed && managed.userPrincipalName) || '';
        }
        state.devices = devices;
        regroup();
        const dt = ((Date.now() - t0) / 1000).toFixed(1);
        Log.info(`Scan finished: ${devices.length} devices, ${state.groups.length} duplicate groups, ${dt}s`);
        $('resultsStatus').textContent = `${devices.length} devices scanned in ${dt}s. Nothing changes until you pick an action and confirm it.`;
        activateTab('results');
      } catch (e) {
        if (e.message === 'Cancelled') { toast('Scan cancelled', 'warn'); return; }
        Log.err('Scan failed:', e);
        toast(`Scan failed: ${e.message}`, 'err');
      } finally {
        box.hidden = true;
        state.cancelToken = null;
      }
    });
  }

  function regroup() {
    state.groups = Devices.findDuplicates(state.devices, { includeDisabledInComparison: $('includeDisabled').checked });
    state.selected.clear();
    renderGroups();
  }

  // Recompute what is visible and drop selections the filter hides.
  function applyView() {
    state.visible = Devices.filterGroups(state.groups, {
      search: $('resultFilter').value,
      hideMixedAutopilot: $('hideMixedAutopilot').checked,
    });
    Devices.trimSelection(state.selected, Devices.staleIds(state.visible));
  }

  // ---------- render ----------
  function roleBadge(d, g) {
    if (d === g.current) return '<span class="badge ok">Keep</span>';
    return d.accountEnabled === false ? '<span class="badge neutral">Disabled</span>' : '<span class="badge warn">Stale</span>';
  }
  function dateCell(v) {
    const s = Devices.fmtDate(v);
    return s ? esc(s) : '<span class="muted">Never</span>';
  }

  function groupHtml(g) {
    const title = g.current.displayName || g.devices[0].displayName || '(no name)';
    const rows = g.devices.map((d) => {
      const isCurrent = d === g.current;
      const check = isCurrent
        ? ''
        : `<input type="checkbox" data-act="sel" aria-label="Select ${esc(d.displayName || d.id)}" ${state.selected.has(d.id) ? 'checked' : ''}/>`;
      return `<tr data-id="${esc(d.id)}" class="${isCurrent ? 'is-current' : ''} ${d.accountEnabled === false ? 'is-disabled' : ''}">
        <td class="col-check"${isCurrent ? ' title="The most recent device in a group is never selected"' : ''}>${check}</td>
        <td>${roleBadge(d, g)}</td>
        <td><strong>${esc(d.displayName || '(no name)')}</strong>${d.primaryUser ? `<br><span class="muted small">${esc(d.primaryUser)}</span>` : ''}</td>
        <td>${esc(Devices.osLabel(d))}</td>
        <td>${esc(d.trustType || '')}</td>
        <td>${d.isAutopilot ? 'Yes' : '<span class="muted">No</span>'}</td>
        <td>${d.accountEnabled === false ? '<span class="muted">No</span>' : 'Yes'}</td>
        <td class="nowrap">${dateCell(d.approximateLastSignInDateTime)}</td>
        <td class="nowrap">${dateCell(d.createdDateTime)}</td>
        <td class="url-cell mono" title="${esc(d.id)}">${esc(d.id || '')}</td>
      </tr>`;
    }).join('');
    return `<div class="group-card">
      <div class="group-header">
        <div>
          <div class="group-title">${esc(title)}</div>
          <div class="hwid">HWID ${esc(g.hwid)}</div>
        </div>
        <span class="badge neutral">${g.deviceCount} devices, ${g.stale.length} stale</span>
      </div>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th class="col-check"></th><th>Role</th><th>Display name</th><th>OS</th><th>Trust</th><th>Autopilot</th><th>Enabled</th><th>Last sign-in</th><th>Created</th><th>Object id</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    </div>`;
  }

  function renderGroups() {
    applyView();
    const scanned = state.devices.length > 0;
    const hasGroups = state.groups.length > 0;
    const visibleStale = state.visible.reduce((n, g) => n + g.stale.length, 0);
    const totalStale = state.groups.reduce((n, g) => n + g.stale.length, 0);

    $('statTiles').hidden = !scanned;
    $('filterRow').hidden = !hasGroups;
    $('statDevices').textContent = String(state.devices.length);
    $('statGroups').textContent = String(state.groups.length);
    $('statGroupsFoot').textContent = `${state.visible.length} shown`;
    $('statStale').textContent = String(totalStale);
    $('statStaleFoot').textContent = `${visibleStale} shown`;
    $('exportXlsxBtn').disabled = !state.visible.length;
    $('exportCsvBtn').disabled = !state.visible.length;
    $('selectAllStaleBtn').disabled = !visibleStale;

    const root = $('groupsContainer');
    root.innerHTML = state.visible.map(groupHtml).join('');
    if (hasGroups && !state.visible.length) root.innerHTML = '<p class="table-hint">No groups match the filter.</p>';

    const empty = $('noResults');
    empty.hidden = !scanned || hasGroups;
    if (scanned && !hasGroups) {
      $('noResultsTitle').textContent = 'No duplicates';
      $('noResultsText').textContent = `None of the ${state.devices.length} devices share a hardware id. If you expected duplicates, tick "Include already disabled devices" on the Scan tab.`;
    }
    updateSelection();
  }

  function updateSelection() {
    const n = state.selected.size;
    $('selCount').textContent = String(n);
    $('statSelected').textContent = String(n);
    $('actionBar').hidden = !state.groups.length;
    $('disableBtn').disabled = n === 0;
    $('deleteBtn').disabled = n === 0;
    $('selectNoneBtn').disabled = n === 0;
  }

  function onGroupsChange(e) {
    const cb = e.target.closest('input[data-act="sel"]');
    if (!cb) return;
    const id = cb.closest('tr').dataset.id;
    if (cb.checked) state.selected.add(id); else state.selected.delete(id);
    updateSelection();
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
    const li = h('li', null, [h('span.warn-mark', { text: '...' }), ` ${d.displayName || d.id}`]);
    list.appendChild(li);
    return {
      ok: (text) => li.replaceChildren(h('span.ok-mark', { text: 'OK' }), ` ${text}`),
      skip: (text) => li.replaceChildren(h('span.warn-mark', { text: 'SKIP' }), ` ${text}`),
      fail: (text) => li.replaceChildren(h('span.err-mark', { text: 'ERR' }), ` ${text}`),
    };
  }

  async function bulkAction(kind) {
    applyView();
    const allowed = Devices.staleIds(state.visible);
    const rows = state.devices.filter((d) => state.selected.has(d.id) && allowed.has(d.id));
    if (!rows.length) { updateSelection(); return; }
    const del = kind === 'delete';
    const extAttr = del ? '' : $('extAttrSelect').value;
    const n = rows.length;
    const ok = await confirm({
      title: `${del ? 'Delete' : 'Disable'} ${n} device${n === 1 ? '' : 's'}?`,
      body: del
        ? 'These devices are deleted from Entra ID. This cannot be undone; a device that is still in use has to register again.'
        : `These devices can no longer be used to sign in to Entra ID until they are enabled again.${extAttr ? ` The disable timestamp goes into ${extAttr}.` : ''}`,
      items: rows.map(Devices.describe),
      okLabel: `${del ? 'Delete' : 'Disable'} ${n}`,
      danger: del,
      wide: true,
    });
    if (!ok) return;
    // Consent for the write scope before the first device, by popup, so this page stays put.
    try { await Auth.getWriteToken(); }
    catch (e) { toast(`No write permission: ${e.message}`, 'err'); return; }

    const list = $('actionResults');
    list.hidden = false;
    list.replaceChildren();
    const tick = startProgress(del ? 'Deleting' : 'Disabling', n);
    const ext = extAttr ? { [extAttr]: Devices.buildDisableStampValue() } : null;
    let done = 0, okCount = 0;
    const deleted = new Set();
    await busy($(del ? 'deleteBtn' : 'disableBtn'), () => Concurrency.pmap(rows, async (d) => {
      const line = resultLine(list, d);
      try {
        if (del) {
          await Graph.deleteDevice(d.id);
          deleted.add(d.id);
          line.ok(`Deleted ${d.displayName || d.id}`);
        } else {
          await Graph.disableDevice(d.id, ext);
          d.accountEnabled = false;
          if (ext) d.extensionAttributes = { ...(d.extensionAttributes || {}), ...ext };
          line.ok(`Disabled ${d.displayName || d.id}${ext ? `, stamped ${extAttr}` : ''}`);
        }
        okCount++;
      } catch (e) {
        line.fail(`${d.displayName || d.id}: ${e.message}`);
        Log.err(`${kind} failed for ${d.id}:`, e);
        throw e;
      } finally {
        tick(++done);
      }
    }, { concurrency: concurrency() }));
    $('actionProgress').hidden = true;
    const failed = n - okCount;
    toast(`${del ? 'Deleted' : 'Disabled'} ${okCount} of ${n}${failed ? `, ${failed} failed (see the list and the debug log)` : ''}`, failed ? 'warn' : 'ok');
    Log.info(`${kind}: ${okCount}/${n} succeeded`);
    if (deleted.size) state.devices = state.devices.filter((d) => !deleted.has(d.id));
    regroup();
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
        const list = await Graph.listDisabledDevices((n) => { $('expiredStatus').textContent = `${n} disabled devices loaded`; });
        state.expired = Devices.findExpiredDisabled(list, attr, grace, new Date());
        state.expiredCtx = { attr, grace };
        $('expiredStatus').textContent = `${list.length} disabled devices checked, ${state.expired.length} with a ${attr} timestamp older than ${grace} days.`;
        Log.info(`Grace period search: ${state.expired.length} of ${list.length} disabled devices expired (${attr}, ${grace} days)`);
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
    $('deleteExpiredBtn').disabled = !has;
    document.querySelector('#expiredTable tbody').innerHTML = state.expired.map((e) => `<tr>
      <td>${esc(e.device.displayName || '(no name)')}</td>
      <td>${esc(Devices.osLabel(e.device))}</td>
      <td class="nowrap">${esc(Devices.fmtDate(e.disabledAt))}</td>
      <td class="num">${e.daysDisabled}</td>
      <td class="url-cell mono" title="${esc(e.device.id)}">${esc(e.device.id || '')}</td>
    </tr>`).join('');
  }

  async function deleteExpired() {
    const ctx = state.expiredCtx;
    if (!state.expired.length || !ctx) return;
    const rows = state.expired.slice();
    const n = rows.length;
    const ok = await confirm({
      title: `Delete ${n} expired device${n === 1 ? '' : 's'}?`,
      body: `These disabled devices are deleted from Entra ID. Each one is read again first; a device that was enabled again or whose ${ctx.attr} changed is skipped. Deleting cannot be undone.`,
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
    let skipped = 0, failed = 0;
    await busy($('deleteExpiredBtn'), () => Concurrency.pmap(rows, async (e) => {
      const d = e.device;
      const line = resultLine(list, d);
      try {
        const fresh = await Graph.getDevice(d.id);
        const check = Devices.recheckExpired(fresh, ctx.attr, e.stampValue, Date.now() - ctx.grace * DAY);
        if (!check.ok) {
          skipped++;
          line.skip(`${d.displayName || d.id}: ${check.reason}`);
          Log.warn(`Grace delete skipped ${d.id}: ${check.reason}`);
          return;
        }
        await Graph.deleteDevice(d.id);
        deleted.add(d.id);
        line.ok(`Deleted ${d.displayName || d.id}`);
      } catch (err) {
        failed++;
        line.fail(`${d.displayName || d.id}: ${err.message}`);
        Log.err(`Grace delete failed for ${d.id}:`, err);
      }
    }, { concurrency: concurrency() }));
    toast(`Deleted ${deleted.size} of ${n}${skipped ? `, ${skipped} skipped` : ''}${failed ? `, ${failed} failed` : ''}`, skipped || failed ? 'warn' : 'ok');
    state.expired = state.expired.filter((e) => !deleted.has(e.device.id));
    state.devices = state.devices.filter((d) => !deleted.has(d.id));
    renderExpired(true);
    if (deleted.size && state.groups.length) regroup();
  }

  // ---------- export ----------
  const EXPORT_COLUMNS = [
    { key: 'hwid', label: 'Hardware id' },
    { key: 'count', label: 'Devices in group' },
    { key: 'role', label: 'Role' },
    { key: 'displayName', label: 'Display name' },
    { key: 'deviceId', label: 'Device id' },
    { key: 'id', label: 'Object id' },
    { key: 'os', label: 'OS' },
    { key: 'trustType', label: 'Trust type' },
    { key: 'autopilot', label: 'Autopilot' },
    { key: 'primaryUser', label: 'Primary user' },
    { key: 'enabled', label: 'Enabled' },
    { key: 'lastSignIn', label: 'Last sign-in' },
    { key: 'created', label: 'Created' },
    { key: 'managed', label: 'Managed' },
    { key: 'compliant', label: 'Compliant' },
    { key: 'stamp', label: 'Disable timestamp' },
  ];
  function exportRows() { applyView(); return Devices.rowsFromGroups(state.visible, $('extAttrSelect').value); }

  // ---------- boot ----------
  let activateTab = () => {};

  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    activateTab = Toolkit.tabs('#tabs');
    loadPrefs();
    $('signInBtn').addEventListener('click', onSignIn);
    $('scanBtn').addEventListener('click', scan);
    $('cancelBtn').addEventListener('click', () => { if (state.cancelToken) state.cancelToken.cancel(); });
    PREF_FIELDS.forEach((id) => $(id).addEventListener('change', () => {
      savePrefs();
      if (id !== 'concurrency' && state.expiredCtx) clearExpired('Settings changed, search again.');
    }));
    $('includeDisabled').addEventListener('change', () => { savePrefs(); if (state.devices.length) regroup(); });
    $('hideMixedAutopilot').addEventListener('change', () => { savePrefs(); renderGroups(); });
    $('resultFilter').addEventListener('input', renderGroups);
    $('groupsContainer').addEventListener('change', onGroupsChange);
    $('selectAllStaleBtn').addEventListener('click', () => {
      applyView();
      state.selected = Devices.staleIds(state.visible);
      renderGroups();
    });
    $('selectNoneBtn').addEventListener('click', () => { state.selected.clear(); renderGroups(); });
    $('disableBtn').addEventListener('click', () => bulkAction('disable'));
    $('deleteBtn').addEventListener('click', () => bulkAction('delete'));
    $('findExpiredBtn').addEventListener('click', findExpired);
    $('deleteExpiredBtn').addEventListener('click', deleteExpired);
    $('exportXlsxBtn').addEventListener('click', () => busy($('exportXlsxBtn'), () => Export.xlsx([{ name: 'Duplicates', rows: exportRows(), columns: EXPORT_COLUMNS }], `duped-devices-${Export.stamp()}.xlsx`).catch((e) => toast(e.message, 'err'))));
    $('exportCsvBtn').addEventListener('click', () => Export.csv(exportRows(), { columns: EXPORT_COLUMNS, filename: `duped-devices-${Export.stamp()}.csv` }));

    try {
      await Auth.init();
      refreshAuthUI();
      Log.info('DupedDevices ready.');
    } catch (e) {
      Log.err('Start-up failed:', e);
      toast(`Start-up failed: ${e.message}`, 'err');
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
