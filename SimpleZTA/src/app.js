// SimpleZTA UI: sign-in, run tab, catalog and boot.
(() => {
  const { esc, toast, confirm, busy, setUser, h } = Toolkit;

  const GRAPH_READ = [
    'User.Read',
    'Directory.Read.All',
    'Policy.Read.All',
    'Application.Read.All',
    'RoleManagement.Read.Directory',
    'UserAuthenticationMethod.Read.All',
    'AuditLog.Read.All',
    'Reports.Read.All',
    'AccessReview.Read.All',
    'DirectoryRecommendations.Read.All',
    'DeviceManagementConfiguration.Read.All',
    'DeviceManagementServiceConfig.Read.All',
    'DeviceManagementManagedDevices.Read.All',
    'DeviceManagementApps.Read.All',
    'EntitlementManagement.Read.All',
    'DeviceManagementRBAC.Read.All',
    'IdentityRiskEvent.Read.All',
    'IdentityRiskyUser.Read.All',
    'NetworkAccess.Read.All',
    'InformationProtectionPolicy.Read',
  ];

  // 'arm' is only requested when a run contains Azure controls.
  window.Auth = ToolkitAuth.create({
    scopes: {
      read: GRAPH_READ,
      arm: ['https://management.azure.com/user_impersonation'],
    },
  });

  const PENDING_RUN_KEY = 'SimpleZTA:pending-run:v1';
  const PENDING_MAX_AGE_MS = 10 * 60 * 1000;
  const $ = (id) => document.getElementById(id);
  let showTab = () => {};

  function fmtDuration(ms) {
    if (!Number.isFinite(Number(ms))) return '-';
    const val = Math.max(0, Math.round(Number(ms)));
    if (val < 1000) return `${val} ms`;
    const sec = Math.round(val / 1000);
    if (sec < 60) return `${sec}s`;
    return `${Math.floor(sec / 60)}m ${sec % 60}s`;
  }

  function implBadge(level) {
    if (level === 'full') return '<span class="badge ok">full</span>';
    if (level === 'partial') return '<span class="badge warn">partial</span>';
    return '<span class="badge neutral">planned</span>';
  }

  // ---------- pending run (survives a consent redirect when the popup is blocked) ----------
  function savePendingRun(ids) {
    try { sessionStorage.setItem(PENDING_RUN_KEY, JSON.stringify({ ids, createdAt: Date.now() })); } catch { /* storage blocked */ }
  }
  function loadPendingRun() {
    try {
      const parsed = JSON.parse(sessionStorage.getItem(PENDING_RUN_KEY) || 'null');
      const ids = Array.isArray(parsed?.ids) ? parsed.ids.map((x) => `${x}`) : [];
      if (!ids.length || Date.now() - Number(parsed.createdAt || 0) > PENDING_MAX_AGE_MS) return null;
      return { ids };
    } catch { return null; }
  }
  function clearPendingRun() {
    try { sessionStorage.removeItem(PENDING_RUN_KEY); } catch { /* storage blocked */ }
  }

  // ---------- auth ----------
  function refreshAuthUI() {
    const signedIn = !!Auth.getAccount();
    setUser(Auth.userInfo(onSignOut));
    $('connectCard').hidden = signedIn;
    $('runCard').hidden = !signedIn;
  }

  async function onSignIn() {
    try { await Auth.signIn('read'); }
    catch (e) { Log.err('Sign-in failed:', e); toast(`Sign-in failed: ${e.message}`, 'err'); }
  }

  async function onSignOut() {
    const scans = Store.loadScans();
    if (scans.length) {
      const wipe = await confirm({
        title: 'Also clear the scan history?',
        body: 'Signing out ends the session in this tab. The scan history stays in this browser unless you clear it now.',
        items: [Store.describeHistory(scans)],
        okLabel: 'Clear history and sign out',
        cancelLabel: 'Keep history',
        danger: true,
      });
      if (wipe) Store.clearScans();
    }
    await Auth.signOut();
  }

  // ---------- run list ----------
  function renderRunList(tests) {
    const q = ($('searchBox').value || '').toLowerCase();
    const pillar = $('pillarSelect').value;
    const tenantType = $('tenantTypeSelect').value;
    const risk = $('riskSelect').value;
    const impl = $('implementationSelect').value;

    const filtered = tests.filter((t) => {
      if (pillar && pillar !== 'All' && !(t.pillars && t.pillars.length ? t.pillars : [t.pillar]).includes(pillar)) return false;
      if (tenantType && !(t.tenantType || []).includes(tenantType)) return false;
      if (risk && t.risk !== risk) return false;
      if (impl && t.implementationLevel !== impl) return false;
      if (!q) return true;
      return `${t.id} ${t.title} ${t.category} ${t.pillar}`.toLowerCase().includes(q);
    });

    const selected = new Set((window.__selectedTests || []).map((x) => `${x}`));
    const html = filtered.map((t) => {
      const partialNotes = t.implementationLevel === 'partial' && t.implementationNotes
        ? `<div class="partial-notes"><div><strong>Implemented:</strong> ${esc(t.implementationNotes.implemented || '-')}</div><div><strong>Not implemented:</strong> ${esc(t.implementationNotes.notImplemented || '-')}</div></div>`
        : '';
      const docs = t.documentationUrl
        ? `<div class="test-docs"><a href="${esc(Toolkit.safeUrl(t.documentationUrl))}" target="_blank" rel="noopener noreferrer">Docs root</a> <a href="${esc(Toolkit.safeUrl(t.documentationSearchUrl))}" target="_blank" rel="noopener noreferrer">Search docs</a></div>`
        : '';
      return `<label class="test-item"><input type="checkbox" data-id="${esc(t.id)}" ${selected.has(t.id) ? 'checked' : ''}/><div><div><strong>${esc(t.id)}</strong> ${esc(t.title)}</div><div class="test-meta">${esc(t.pillar)} | ${esc(t.category)} | risk ${esc(t.risk)}${t.needsArm ? ' | Azure' : ''}</div>${partialNotes}${docs}</div><div>${implBadge(t.implementationLevel)}</div></label>`;
    }).join('');

    $('testList').innerHTML = html || '<div class="empty-state small">No controls match this filter.</div>';
    lastFiltered = filtered;
    window.__visibleTests = filtered.map((x) => x.id);
    renderRunStats(tests.length);
  }

  let lastFiltered = [];
  function renderRunStats(total) {
    const count = (lvl) => lastFiltered.filter((x) => x.implementationLevel === lvl).length;
    $('runStats').textContent = `${lastFiltered.length} of ${total} controls shown, ${(window.__selectedTests || []).length} selected | full ${count('full')} | partial ${count('partial')} | planned ${count('none')}`;
  }

  // ---------- profiles ----------
  const BUILTIN_PROFILES = {
    'Quick wins (high risk, low effort)': (t) => t.risk === 'High' && t.implementationCost === 'Low' && t.implemented,
    'High risk only': (t) => t.risk === 'High',
    'Identity baseline': (t) => (t.pillars || [t.pillar]).includes('Identity') && t.implemented,
    'Implemented controls only': (t) => t.implemented,
  };

  function refreshProfilePicker() {
    const picker = $('profileSelect');
    const saved = Object.keys(Store.loadProfiles()).sort();
    picker.innerHTML = '<option value="">Profiles...</option>'
      + Object.keys(BUILTIN_PROFILES).map((n) => `<option value="builtin:${esc(n)}">${esc(n)}</option>`).join('')
      + saved.map((n) => `<option value="saved:${esc(n)}">Saved: ${esc(n)}</option>`).join('');
    $('deleteProfileBtn').disabled = !saved.length;
  }

  async function askProfileName(count) {
    const input = h('input', { type: 'text', maxlength: '80', 'aria-label': 'Profile name', class: 'profile-name' });
    const body = h('div', null, [h('p', { text: `Save the ${count} selected controls as a profile in this browser.` }), input]);
    setTimeout(() => input.focus(), 0);
    const ok = await Toolkit.modal({ title: 'Save profile', body, actions: [{ label: 'Cancel', value: false }, { label: 'Save', value: true, variant: 'primary' }] });
    return ok ? Store.cleanProfileName(input.value) : '';
  }

  async function deleteProfileDialog() {
    const names = Object.keys(Store.loadProfiles()).sort();
    if (!names.length) return;
    const select = h('select', { 'aria-label': 'Profile' }, names.map((n) => h('option', { value: n, text: n })));
    const ok = await Toolkit.modal({
      title: 'Delete a saved profile',
      body: h('div', null, [h('p', { text: 'Only the profile is removed, scans stay.' }), select]),
      actions: [{ label: 'Cancel', value: false }, { label: 'Delete', value: true, variant: 'danger' }],
    });
    if (!ok) return;
    Store.deleteProfile(select.value);
    refreshProfilePicker();
    toast(`Profile "${select.value}" deleted`, 'ok');
  }

  // ---------- run ----------
  function setProgress({ done, total, current, elapsedTotalMs, elapsedCurrentMs, etaMs, currentEstimateMs, subProgress }) {
    const pct = total ? Math.round((100 * done) / total) : 0;
    $('progressFill').style.width = `${pct}%`;
    $('progressText').textContent = `${done}/${total} ${current ? `${current.id} ${current.title}` : ''}`;
    $('progressMeta').textContent = `Elapsed ${fmtDuration(elapsedTotalMs)}, about ${fmtDuration(etaMs)} left`;
    const testInfo = current ? `Control running ${fmtDuration(elapsedCurrentMs)}${currentEstimateMs ? ` of about ${fmtDuration(currentEstimateMs)}` : ''}` : '';
    const subInfo = subProgress?.message
      ? `${subProgress.message}${Number.isFinite(subProgress.percent) ? ` (${Math.round(subProgress.percent)}%)` : ''}`
      : '';
    $('subTestProgress').textContent = [testInfo, subInfo].filter(Boolean).join(' | ');
  }

  async function runSelectedTests(selected, { resume = false } = {}) {
    if (!selected.length) { toast('Select at least one control.', 'warn'); return; }
    if (!Auth.getAccount()) { toast('Sign in first.', 'warn'); return; }
    const ids = selected.map((t) => t.id);
    Store.rememberSelection(ids);

    // A blocked consent popup falls back to a redirect; the run resumes once after it.
    // On resume the marker is cleared first so a second redirect cannot loop.
    if (resume) clearPendingRun(); else savePendingRun(ids);
    try {
      await Auth.getToken('read');
    } catch (e) {
      clearPendingRun();
      Log.err('Graph token failed:', e);
      toast(`Could not get a Microsoft Graph token: ${e.message}`, 'err');
      return;
    }
    if (selected.some((t) => t.needsArm)) {
      try {
        await Auth.getToken('arm');
        Log.info('Azure Resource Manager token acquired for the Azure controls.');
      } catch (e) {
        Log.warn('No Azure Resource Manager token, Azure controls will be skipped:', e.message || e);
        toast('No Azure access granted. Azure controls are skipped in this run.', 'warn');
      }
    }
    clearPendingRun();

    $('progressBox').hidden = false;
    $('progressLog').textContent = '';
    try {
      const scan = await Runner.run(selected, {
        collectTenantInfo: $('tenantInfoToggle').checked,
        tenantInfoDays: Number($('signInDaysSelect').value) || 7,
        progress: setProgress,
        log: (line) => {
          const el = $('progressLog');
          el.textContent += `${line}\n`;
          el.scrollTop = el.scrollHeight;
          Log.info(line);
        },
      });

      UIResults.refreshPicker(scan.id);
      UIHistory.refresh();
      if (scan?.persistence?.saved === false) toast('Scan done, but it could not be stored: browser storage is full. Export it and clear the history.', 'warn');
      else if (scan?.persistence?.warning) toast(`Scan done. ${scan.persistence.warning}`, 'warn');
      else toast(`Scan done: ${scan.summary.passed} passed, ${scan.summary.failed} failed.`, 'ok');
      showTab('dashboard');
    } catch (e) {
      Log.err('Run failed:', e);
      toast(`Run failed: ${e.message || e}`, 'err');
    } finally {
      $('progressBox').hidden = true;
    }
  }

  function renderCatalog(tests) {
    const partial = tests.filter((t) => t.implementationLevel === 'partial');
    const count = (lvl) => tests.filter((t) => t.implementationLevel === lvl).length;
    const partialRows = partial.length
      ? `<div class="table-wrap"><table class="data-table"><thead><tr><th>Control</th><th>What is implemented</th><th>What is not</th></tr></thead><tbody>${partial.map((t) => `<tr><td><strong>${esc(t.id)}</strong> ${esc(t.title)}</td><td>${esc(t.implementationNotes?.implemented || '-')}</td><td>${esc(t.implementationNotes?.notImplemented || '-')}</td></tr>`).join('')}</tbody></table></div>`
      : '<p class="muted">No partial implementations registered.</p>';
    $('catalogStats').innerHTML = `
      <div class="grid grid-4 stats">
        <div class="stat accent-primary"><span class="label">Controls</span><span class="value">${tests.length}</span><span class="foot">from Microsoft metadata</span></div>
        <div class="stat accent-ok"><span class="label">Full</span><span class="value">${count('full')}</span><span class="foot">same logic as the module</span></div>
        <div class="stat accent-warn"><span class="label">Partial</span><span class="value">${partial.length}</span><span class="foot">approximation, see below</span></div>
        <div class="stat"><span class="label">Azure controls</span><span class="value">${tests.filter((t) => t.needsArm).length}</span><span class="foot">need Azure Resource Manager</span></div>
      </div>
      <h3>Partial implementations</h3>
      ${partialRows}`;
  }

  async function initRunUi() {
    const tests = await Registry.getAll();
    window.__allTests = tests;
    const remembered = Store.loadSelection();
    window.__selectedTests = remembered.length ? remembered.filter((id) => tests.some((t) => t.id === id)) : tests.map((t) => t.id);

    const refresh = () => renderRunList(tests);
    ['searchBox', 'pillarSelect', 'tenantTypeSelect', 'riskSelect', 'implementationSelect'].forEach((id) => $(id).addEventListener('input', refresh));

    $('testList').addEventListener('change', (e) => {
      const cb = e.target.closest('input[type="checkbox"][data-id]');
      if (!cb) return;
      const s = new Set(window.__selectedTests || []);
      if (cb.checked) s.add(cb.dataset.id); else s.delete(cb.dataset.id);
      window.__selectedTests = Array.from(s);
      renderRunStats(tests.length);
    });
    $('selectVisibleBtn').addEventListener('click', () => {
      window.__selectedTests = Array.from(new Set([...(window.__selectedTests || []), ...(window.__visibleTests || [])]));
      refresh();
    });
    $('clearSelectionBtn').addEventListener('click', () => { window.__selectedTests = []; refresh(); });

    $('profileSelect').addEventListener('change', (e) => {
      const value = e.target.value;
      if (!value) return;
      if (value.startsWith('builtin:')) {
        const fn = BUILTIN_PROFILES[value.slice(8)];
        if (fn) window.__selectedTests = tests.filter(fn).map((t) => t.id);
      } else if (value.startsWith('saved:')) {
        const list = Store.loadProfiles()[value.slice(6)] || [];
        window.__selectedTests = list.filter((id) => tests.some((t) => t.id === id));
      }
      refresh();
      toast(`Profile applied: ${(window.__selectedTests || []).length} controls selected.`, 'info');
      e.target.value = '';
    });
    $('saveProfileBtn').addEventListener('click', async () => {
      const sel = window.__selectedTests || [];
      if (!sel.length) { toast('Select at least one control before saving a profile.', 'warn'); return; }
      const name = await askProfileName(sel.length);
      if (!name) return;
      Store.saveProfile(name, sel);
      refreshProfilePicker();
      toast(`Profile "${name}" saved.`, 'ok');
    });
    $('deleteProfileBtn').addEventListener('click', deleteProfileDialog);
    refreshProfilePicker();

    $('runBtn').addEventListener('click', () => busy($('runBtn'), () => {
      const ids = new Set(window.__selectedTests || []);
      return runSelectedTests(tests.filter((t) => ids.has(t.id)));
    }));
    $('cancelBtn').addEventListener('click', () => { Runner.cancel(); toast('Cancelling after the current control.', 'info'); });

    refresh();
    renderCatalog(tests);

    const pending = loadPendingRun();
    if (pending && Auth.getAccount()) {
      const pendingSelected = tests.filter((t) => pending.ids.includes(t.id));
      if (pendingSelected.length) {
        window.__selectedTests = pendingSelected.map((t) => t.id);
        refresh();
        toast('Resuming the scan after sign-in or consent.', 'info');
        await busy($('runBtn'), () => runSelectedTests(pendingSelected, { resume: true }));
        return;
      }
    }
    clearPendingRun();
  }

  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    showTab = Toolkit.tabs('#tabs', {
      onChange: (id) => {
        if (id === 'history') UIHistory.refresh();
        if (id === 'dashboard') UIDashboard.refresh();
      },
    });
    window.ZtaApp = { showTab: (id) => showTab(id) };
    UIResults.init();
    UIHistory.init();
    $('signInBtn').addEventListener('click', onSignIn);

    try {
      await Auth.init();
      refreshAuthUI();
    } catch (e) {
      Log.err('Start-up failed:', e);
      toast(`Start-up failed: ${e.message}`, 'err');
    }
    try {
      UIResults.refreshPicker();
      UIHistory.refresh();
      await initRunUi();
      Log.info('SimpleZTA ready.');
    } catch (e) {
      Log.err('Loading the control catalog failed:', e);
      toast(`Loading the control catalog failed: ${e.message}`, 'err');
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
