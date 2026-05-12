// DupedDevices main wiring.
(() => {
  const state = {
    devices: [],
    groups: [],
    expired: [],
    selected: new Set(),     // selected stale device ObjectIds
    cancelToken: null,
    haveWriteConsent: false,
  };

  // ---------- Tabs ----------
  document.querySelectorAll('.tab').forEach(t => {
    t.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(x => x.classList.remove('active'));
      t.classList.add('active');
      document.getElementById(`tab-${t.dataset.tab}`).classList.add('active');
    });
  });

  // ---------- Dark mode ----------
  if (localStorage.getItem('duped-theme') === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
  document.getElementById('darkToggle').addEventListener('click', () => {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    if (isDark) { document.documentElement.removeAttribute('data-theme'); localStorage.setItem('duped-theme', ''); }
    else { document.documentElement.setAttribute('data-theme', 'dark'); localStorage.setItem('duped-theme', 'dark'); }
  });

  // ---------- Persisted prefs ----------
  function loadPrefs() {
    try {
      const p = JSON.parse(localStorage.getItem('duped-prefs') || '{}');
      if (p.extAttr) document.getElementById('extAttrSelect').value = p.extAttr;
      if (p.gracePeriod != null) document.getElementById('gracePeriod').value = p.gracePeriod;
      if (p.concurrency) document.getElementById('concurrency').value = p.concurrency;
      if (p.cleanupMode) {
        const r = document.querySelector(`input[name=cleanupMode][value="${p.cleanupMode}"]`);
        if (r) r.checked = true;
      }
      if (p.includeDisabled) document.getElementById('includeDisabled').checked = true;
      if (p.hideMixedAutopilot === false) document.getElementById('hideMixedAutopilot').checked = false;
    } catch {}
  }
  function savePrefs() {
    const p = {
      extAttr: document.getElementById('extAttrSelect').value,
      gracePeriod: +document.getElementById('gracePeriod').value || 30,
      concurrency: +document.getElementById('concurrency').value || 6,
      cleanupMode: document.querySelector('input[name=cleanupMode]:checked').value,
      includeDisabled: document.getElementById('includeDisabled').checked,
      hideMixedAutopilot: document.getElementById('hideMixedAutopilot').checked,
    };
    localStorage.setItem('duped-prefs', JSON.stringify(p));
  }
  ['extAttrSelect','gracePeriod','concurrency','includeDisabled','hideMixedAutopilot'].forEach(id =>
    document.getElementById(id).addEventListener('change', savePrefs));
  document.querySelectorAll('input[name=cleanupMode]').forEach(r => r.addEventListener('change', () => { savePrefs(); refreshActionLabel(); }));

  // ---------- Verbose log ----------
  document.getElementById('verboseLog').addEventListener('change', (e) => Log.setVerbose(e.target.checked));
  document.getElementById('clearLogBtn').addEventListener('click', () => Log.clear());

  // ---------- Auth ----------
  function setSignedInUI(account) {
    if (account) {
      document.getElementById('connectCard').classList.add('hidden');
      document.getElementById('discoveryCard').classList.remove('hidden');
      document.getElementById('userBox').classList.remove('hidden');
      document.getElementById('userName').textContent = account.username || account.name || 'Signed in';
      document.getElementById('findExpiredBtn').disabled = false;
    } else {
      document.getElementById('connectCard').classList.remove('hidden');
      document.getElementById('discoveryCard').classList.add('hidden');
      document.getElementById('userBox').classList.add('hidden');
      document.getElementById('findExpiredBtn').disabled = true;
    }
  }
  document.getElementById('signInBtn').addEventListener('click', async () => {
    try {
      const acct = await Auth.signIn(false);
      if (acct) {
        Log.info(`Signed in as ${acct.username}`);
        setSignedInUI(acct);
        loadPrefs();
      }
    } catch (e) { Log.err('Sign-in failed:', e.message); alert('Sign-in failed: ' + e.message); }
  });
  document.getElementById('signOutBtn').addEventListener('click', async () => {
    await Auth.signOut();
    state.devices = []; state.groups = []; state.selected.clear();
    setSignedInUI(null);
    renderGroups();
  });

  // ---------- Scan ----------
  document.getElementById('scanBtn').addEventListener('click', async () => {
    savePrefs();
    const btn = document.getElementById('scanBtn');
    const box = document.getElementById('progressBox');
    btn.disabled = true; box.classList.remove('hidden');
    setProgress(0, 0, 'Loading devices from Graph...');
    state.cancelToken = new Concurrency.CancelToken();
    try {
      const t0 = Date.now();
      const devices = await Graph.listAllDevices((count, page) => {
        setProgress(count, count, `Loaded ${count} devices (page ${page})...`);
        if (state.cancelToken.cancelled) throw new Error('Cancelled');
      });

      const managedRes = await Graph.tryListManagedDevices((count, page) => {
        setProgress(count, count, `Loaded ${count} Intune managed devices (page ${page})...`);
        if (state.cancelToken.cancelled) throw new Error('Cancelled');
      });

      const notice = document.getElementById('managedNotice');
      if (managedRes.unavailable) {
        notice.textContent = managedRes.message;
        notice.classList.remove('hidden');
      } else {
        notice.classList.add('hidden');
      }

      const managedMap = Devices.buildManagedMap(managedRes.list);
      for (const d of devices) {
        const managed = managedMap.get(Devices.keyForDevice(d)) || null;
        d.managedDevice = managed;
        d.isAutopilot = Devices.isAutopilot(d, managed);
        d.primaryUser = (managed && managed.userPrincipalName) || '';
      }

      state.devices = devices;
      const includeDisabled = document.getElementById('includeDisabled').checked;
      const groups = Devices.findDuplicates(devices, { includeDisabledInComparison: includeDisabled });
      state.groups = groups;
      state.selected.clear();
      const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
      Log.info(`Scan finished: ${devices.length} devices total, ${groups.length} duplicate group(s), in ${elapsed}s`);
      renderGroups();
      // jump to results tab
      document.querySelector('.tab[data-tab="results"]').click();
    } catch (e) {
      Log.err('Scan failed:', e.message);
      alert('Scan failed: ' + e.message);
    } finally {
      btn.disabled = false; box.classList.add('hidden');
      state.cancelToken = null;
    }
  });
  document.getElementById('cancelBtn').addEventListener('click', () => {
    if (state.cancelToken) { state.cancelToken.cancel(); Log.warn('Cancellation requested.'); }
  });
  function setProgress(done, total, label) {
    document.getElementById('progressLabel').textContent = label || '';
    document.getElementById('progressCount').textContent = total ? `${done}/${total}` : `${done}`;
    document.getElementById('progressFill').style.width = total ? `${Math.min(100, Math.round(100 * done / total))}%` : '50%';
  }

  // ---------- Render groups ----------
  function refreshActionLabel() {
    const mode = document.querySelector('input[name=cleanupMode]:checked').value;
    document.getElementById('actionLabel').textContent = mode;
  }

  function isMixedAutopilotGroup(group) {
    let hasAp = false;
    let hasNonAp = false;
    for (const d of group.devices || []) {
      if (d.isAutopilot) hasAp = true;
      else hasNonAp = true;
      if (hasAp && hasNonAp) return true;
    }
    return false;
  }

  function getVisibleGroups() {
    const filter = (document.getElementById('resultFilter').value || '').toLowerCase().trim();
    const hideMixedAutopilot = document.getElementById('hideMixedAutopilot').checked;
    return state.groups.filter((g) => {
      if (hideMixedAutopilot && isMixedAutopilotGroup(g)) return false;
      if (filter) {
        const hay = (g.hwid + ' ' + g.devices.map(d => `${d.displayName||''} ${d.operatingSystem||''}`).join(' ')).toLowerCase();
        if (!hay.includes(filter)) return false;
      }
      return true;
    });
  }

  function renderGroups() {
    const root = document.getElementById('groupsContainer');
    root.innerHTML = '';
    const visibleGroups = getVisibleGroups();
    const summary = document.getElementById('summaryBar');
    const noRes = document.getElementById('noResults');
    let totalStale = 0;
    let visibleStale = 0;
    for (const g of state.groups) totalStale += g.stale.length;
    for (const g of visibleGroups) visibleStale += g.stale.length;
    summary.classList.toggle('hidden', state.groups.length === 0);
    summary.innerHTML = `
      <span><strong>${state.devices.length}</strong> devices scanned</span>
      <span><strong>${state.groups.length}</strong> duplicate group(s)</span>
      <span><strong>${visibleGroups.length}</strong> visible group(s)</span>
      <span><strong>${totalStale}</strong> stale candidate(s)</span>
      <span><strong>${visibleStale}</strong> visible stale candidate(s)</span>
      <span><strong>${state.selected.size}</strong> selected</span>`;
    document.getElementById('exportXlsxBtn').disabled = state.groups.length === 0;
    document.getElementById('exportCsvBtn').disabled = state.groups.length === 0;
    document.getElementById('selectAllStaleBtn').disabled = state.groups.length === 0;
    document.getElementById('selectNoneBtn').disabled = state.selected.size === 0;
    document.getElementById('actionBar').classList.toggle('hidden', state.selected.size === 0);
    document.getElementById('selCount').textContent = state.selected.size;
    document.getElementById('executeBtn').disabled = state.selected.size === 0;
    refreshActionLabel();

    for (const g of visibleGroups) {
      const card = document.createElement('div');
      card.className = 'group-card';
      const titleName = escapeHtml(g.current.displayName || g.devices[0].displayName || '(no name)');
      card.innerHTML = `
        <div class="group-header">
          <div>
            <div class="group-title">${titleName}</div>
            <div class="hwid">HWID: ${escapeHtml(g.hwid)}</div>
          </div>
          <div class="group-meta">${g.deviceCount} devices &middot; ${g.stale.length} stale</div>
        </div>
        <div class="table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th style="width:32px"></th>
                <th>Role</th>
                <th>Display name</th>
                <th>OS</th>
                <th>Trust</th>
                <th>Autopilot</th>
                <th>Enabled</th>
                <th>Last sign-in</th>
                <th>Created</th>
                <th>Object id</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>`;
      const tbody = card.querySelector('tbody');
      for (const d of g.devices) {
        const isCurrent = d === g.current;
        const isDisabled = d.accountEnabled === false;
        const tr = document.createElement('tr');
        tr.className = `device-row ${isCurrent ? 'is-current' : 'is-stale'} ${isDisabled ? 'is-disabled' : ''}`;
        const role = isCurrent
          ? `<span class="role-pill current">Keep</span>`
          : (isDisabled ? `<span class="role-pill disabled">Disabled</span>` : `<span class="role-pill stale">Stale</span>`);
        const checkbox = isCurrent
          ? '<span class="muted" title="The most recently active duplicate is never selected">-</span>'
          : `<input type="checkbox" data-id="${escapeAttr(d.id)}" ${state.selected.has(d.id) ? 'checked' : ''}>`;
        tr.innerHTML = `
          <td>${checkbox}</td>
          <td>${role}</td>
          <td>${escapeHtml(d.displayName || '')}</td>
          <td>${escapeHtml([d.operatingSystem, d.operatingSystemVersion].filter(Boolean).join(' '))}</td>
          <td>${escapeHtml(d.trustType || '')}</td>
          <td>${d.isAutopilot ? 'Yes' : '<span class="muted">No</span>'}</td>
          <td>${isDisabled ? 'No' : 'Yes'}</td>
          <td>${formatDate(d.approximateLastSignInDateTime)}</td>
          <td>${formatDate(d.createdDateTime)}</td>
          <td class="url-cell" title="${escapeAttr(d.id)}">${escapeHtml(d.id || '')}</td>`;
        tbody.appendChild(tr);
      }
      root.appendChild(card);
    }
    noRes.classList.toggle('hidden', state.groups.length > 0);
    if (state.groups.length === 0) {
      const n = state.devices.length;
      if (n === 0) {
        noRes.innerHTML = `No devices were returned by Microsoft Graph for your account. Either the tenant is empty, the scan failed, or your account lacks <code>Device.Read.All</code>. Check the Debug tab.`;
      } else {
        noRes.innerHTML = `No duplicate groups detected among the <strong>${n}</strong> device(s) returned by Microsoft Graph. Either you genuinely have no duplicates (nice) or no devices in this tenant share a hardware id. If you expected duplicates, try ticking "Include already-disabled devices" on the previous tab.`;
      }
    }
    if (state.groups.length > 0 && visibleGroups.length === 0) {
      const p = document.createElement('p');
      p.className = 'table-hint';
      p.textContent = 'No groups match your filter.';
      root.appendChild(p);
    }
    // wire checkboxes
    root.querySelectorAll('input[type=checkbox][data-id]').forEach(cb => {
      cb.addEventListener('change', () => {
        const id = cb.dataset.id;
        if (cb.checked) state.selected.add(id);
        else state.selected.delete(id);
        // update summary numbers without full re-render
        document.getElementById('selCount').textContent = state.selected.size;
        document.getElementById('actionBar').classList.toggle('hidden', state.selected.size === 0);
        document.getElementById('executeBtn').disabled = state.selected.size === 0;
        document.getElementById('selectNoneBtn').disabled = state.selected.size === 0;
      });
    });
  }
  document.getElementById('resultFilter').addEventListener('input', renderGroups);
  document.getElementById('hideMixedAutopilot').addEventListener('change', () => { savePrefs(); renderGroups(); });
  document.getElementById('selectAllStaleBtn').addEventListener('click', () => {
    state.selected.clear();
    for (const g of getVisibleGroups()) for (const d of g.stale) state.selected.add(d.id);
    renderGroups();
  });
  document.getElementById('selectNoneBtn').addEventListener('click', () => { state.selected.clear(); renderGroups(); });

  // ---------- Execute action ----------
  document.getElementById('executeBtn').addEventListener('click', async () => {
    const mode = document.querySelector('input[name=cleanupMode]:checked').value;
    const ids = Array.from(state.selected);
    if (!ids.length) return;
    const extAttr = document.getElementById('extAttrSelect').value;
    const concurrency = Math.max(1, Math.min(20, +document.getElementById('concurrency').value || 6));
    const verb = mode === 'delete' ? 'PERMANENTLY DELETE' : 'disable';
    if (!confirm(`About to ${verb} ${ids.length} device(s). This will be done as the signed-in user. Continue?`)) return;

    if (!state.haveWriteConsent) {
      try {
        Log.info('Requesting write consent (Directory.AccessAsUser.All)...');
        await Auth.signIn(true);
        state.haveWriteConsent = true;
      } catch (e) {
        Log.err('Failed to acquire write consent:', e.message);
        alert('Failed to acquire write permissions: ' + e.message);
        return;
      }
    }

    // Build deviceId -> device map
    const idMap = new Map();
    for (const g of state.groups) for (const d of g.devices) idMap.set(d.id, d);

    const results = [];
    const ul = document.createElement('ul');
    ul.className = 'action-result-list';
    document.getElementById('actionResults').innerHTML = '';
    document.getElementById('actionResults').appendChild(ul);

    const ext = extAttr ? { [extAttr]: Devices.buildDisableStampValue() } : null;
    document.getElementById('executeBtn').disabled = true;

    await Concurrency.pmap(ids, async (id) => {
      const d = idMap.get(id);
      const li = document.createElement('li');
      li.innerHTML = `<span class="muted">...</span> ${escapeHtml(d.displayName || id)}`;
      ul.appendChild(li);
      try {
        if (mode === 'delete') {
          await Graph.deleteDevice(id);
          li.innerHTML = `<span class="ok-mark">DEL</span> Deleted: ${escapeHtml(d.displayName || id)}`;
          d._dupedAction = 'Deleted';
        } else {
          await Graph.disableDevice(id, ext);
          li.innerHTML = `<span class="ok-mark">OK</span> Disabled: ${escapeHtml(d.displayName || id)}${ext ? ` <span class="muted">(stamped ${escapeHtml(extAttr)})</span>` : ''}`;
          d.accountEnabled = false;
          if (ext) { d.extensionAttributes = { ...(d.extensionAttributes || {}), ...ext }; }
          d._dupedAction = 'Disabled';
        }
        results.push({ id, ok: true });
      } catch (e) {
        li.innerHTML = `<span class="err-mark">ERR</span> ${escapeHtml(d.displayName || id)}: ${escapeHtml(e.message)}`;
        results.push({ id, ok: false, err: e.message });
        Log.err(`Action failed for ${id}:`, e.message);
      }
    }, { concurrency });

    const ok = results.filter(r => r.ok).length;
    Log.info(`Action complete: ${ok}/${results.length} succeeded.`);
    state.selected.clear();
    // Re-derive groups now that some devices are disabled
    const includeDisabled = document.getElementById('includeDisabled').checked;
    state.groups = Devices.findDuplicates(state.devices, { includeDisabledInComparison: includeDisabled });
    renderGroups();
  });

  // ---------- Export ----------
  document.getElementById('exportXlsxBtn').addEventListener('click', () => {
    Exporter.exportXlsx(getVisibleGroups(), document.getElementById('extAttrSelect').value);
  });
  document.getElementById('exportCsvBtn').addEventListener('click', () => {
    Exporter.exportCsv(getVisibleGroups(), document.getElementById('extAttrSelect').value);
  });

  // ---------- Past grace period ----------
  document.getElementById('findExpiredBtn').addEventListener('click', async () => {
    const attr = document.getElementById('extAttrSelect').value;
    if (!attr) { alert('Pick an extension attribute on the first tab; without it the grace-period workflow has nothing to look at.'); return; }
    const grace = +document.getElementById('gracePeriod').value || 30;
    const btn = document.getElementById('findExpiredBtn');
    btn.disabled = true;
    try {
      Log.info(`Looking for disabled devices with ${attr} stamp older than ${grace} days...`);
      let devices = state.devices;
      if (!devices.length) devices = await Graph.listAllDevices(() => {});
      // Only consider devices we (DupedDevices) disabled
      const stamped = devices.filter(d => {
        const v = (d.extensionAttributes || {})[attr];
        return Devices.parseDupedDevicesStamp(v) != null;
      });
      // Use parseDupedDevicesStamp via a custom expired calc
      const cutoff = Date.now() - grace * 86400000;
      const expired = [];
      for (const d of stamped) {
        if (d.accountEnabled !== false) continue;
        const ts = Devices.parseDupedDevicesStamp(d.extensionAttributes[attr]);
        if (ts && ts.getTime() <= cutoff) {
          expired.push({ device: d, disabledAt: ts, daysDisabled: Math.floor((Date.now() - ts.getTime()) / 86400000) });
        }
      }
      expired.sort((a, b) => a.disabledAt - b.disabledAt);
      state.expired = expired;
      renderExpired();
    } catch (e) {
      Log.err('Find expired failed:', e.message);
      alert('Find expired failed: ' + e.message);
    } finally {
      btn.disabled = false;
    }
  });

  function renderExpired() {
    const root = document.getElementById('expiredContainer');
    root.innerHTML = '';
    const del = document.getElementById('deleteExpiredBtn');
    if (!state.expired.length) {
      root.innerHTML = '<p class="table-hint">No disabled devices past the grace period found.</p>';
      del.disabled = true;
      return;
    }
    del.disabled = false;
    const table = document.createElement('div');
    table.className = 'table-wrap';
    table.innerHTML = `<table class="data-table">
      <thead><tr><th>Display name</th><th>OS</th><th>Disabled at</th><th>Days disabled</th><th>Object id</th></tr></thead>
      <tbody></tbody></table>`;
    const tbody = table.querySelector('tbody');
    for (const e of state.expired) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${escapeHtml(e.device.displayName || '')}</td>
        <td>${escapeHtml([e.device.operatingSystem, e.device.operatingSystemVersion].filter(Boolean).join(' '))}</td>
        <td>${formatDate(e.disabledAt)}</td>
        <td>${e.daysDisabled}</td>
        <td class="url-cell" title="${escapeAttr(e.device.id)}">${escapeHtml(e.device.id || '')}</td>`;
      tbody.appendChild(tr);
    }
    root.appendChild(table);
  }

  document.getElementById('deleteExpiredBtn').addEventListener('click', async () => {
    if (!state.expired.length) return;
    if (!confirm(`PERMANENTLY DELETE ${state.expired.length} expired disabled device(s)?`)) return;
    if (!state.haveWriteConsent) {
      try { await Auth.signIn(true); state.haveWriteConsent = true; }
      catch (e) { alert('Failed to acquire write permissions: ' + e.message); return; }
    }
    const concurrency = Math.max(1, Math.min(20, +document.getElementById('concurrency').value || 6));
    const ul = document.createElement('ul');
    ul.className = 'action-result-list';
    document.getElementById('expiredContainer').appendChild(ul);
    const survivors = [];
    await Concurrency.pmap(state.expired, async (e) => {
      const li = document.createElement('li');
      li.innerHTML = `<span class="muted">...</span> ${escapeHtml(e.device.displayName || e.device.id)}`;
      ul.appendChild(li);
      try {
        await Graph.deleteDevice(e.device.id);
        li.innerHTML = `<span class="ok-mark">DEL</span> Deleted: ${escapeHtml(e.device.displayName || e.device.id)}`;
      } catch (err) {
        li.innerHTML = `<span class="err-mark">ERR</span> ${escapeHtml(e.device.displayName || e.device.id)}: ${escapeHtml(err.message)}`;
        survivors.push(e);
      }
    }, { concurrency });
    state.expired = survivors;
    renderExpired();
  });

  // ---------- Helpers ----------
  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);
  }
  function escapeAttr(s) { return escapeHtml(s); }
  function formatDate(d) {
    if (!d) return '<span class="muted">-</span>';
    const dt = d instanceof Date ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    return dt.toISOString().substring(0, 10);
  }

  // ---------- Boot ----------
  (async () => {
    try {
      const acct = await Auth.init();
      if (acct) {
        Log.info(`Already signed in as ${acct.username}`);
        setSignedInUI(acct);
        loadPrefs();
      } else {
        loadPrefs();
      }
    } catch (e) {
      Log.err('Auth init failed:', e.message);
      alert('Auth init failed: ' + e.message);
    }
  })();
})();
