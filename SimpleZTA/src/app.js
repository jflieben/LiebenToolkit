(() => {
  // Suppress known non-blocking MSAL 3.27.0 iframe sandbox warnings.
  // These warnings indicate the iframe sandbox can escape or attempt top-level navigation,
  // but do not prevent the auth redirect flow from completing successfully.
  // See: https://github.com/AzureAD/microsoft-authentication-library-for-js/issues/
  const origWarn = console.warn;
  console.warn = function(...args) {
    const text = args.map(a => `${a}`).join(' ');
    if (text.includes('iframe') && text.includes('sandbox') && text.includes('allow-')) return; // MSAL iframe warnings
    if (text.includes('allow-top-navigation') && text.includes('microsoftonline')) return;      // MSAL redirect warnings
    return origWarn.apply(console, args);
  };

  const THEME_KEY = 'SimpleZTA:theme';
  const PENDING_RUN_KEY = 'SimpleZTA:pending-run:v1';

  function fmtDuration(ms) {
    if (!Number.isFinite(Number(ms))) return '-';
    const val = Math.max(0, Math.round(Number(ms)));
    if (val < 1000) return `${val} ms`;
    const sec = Math.round(val / 1000);
    if (sec < 60) return `${sec}s`;
    const min = Math.floor(sec / 60);
    const rem = sec % 60;
    return `${min}m ${rem}s`;
  }

  function toast(text) {
    const el = document.getElementById('toast');
    el.textContent = text;
    el.classList.remove('hidden');
    setTimeout(() => el.classList.add('hidden'), 3500);
  }

  function setTheme(theme) {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem(THEME_KEY, theme);
  }

  function initTheme() {
    const t = localStorage.getItem(THEME_KEY) || 'light';
    setTheme(t);
    document.getElementById('themeBtn').addEventListener('click', () => {
      setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');
    });
  }

  function initTabs() {
    document.querySelectorAll('.tab').forEach(btn => {
      btn.addEventListener('click', () => {
        const tab = btn.dataset.tab;
        document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x === btn));
        document.querySelectorAll('.panel').forEach(p => p.classList.toggle('active', p.id === `panel-${tab}`));
        if (tab === 'history') UIHistory.refresh();
        if (tab === 'dashboard') UIDashboard.refresh();
      });
    });
  }

  function statusBadge(level) {
    if (level === 'full') return '<span class="badge implemented">implemented</span>';
    if (level === 'partial') return '<span class="badge partial">partial</span>';
    return '<span class="badge">planned</span>';
  }

  function savePendingRun(testIds) {
    try {
      sessionStorage.setItem(PENDING_RUN_KEY, JSON.stringify({
        ids: testIds,
        createdAt: Date.now(),
      }));
    } catch {
      // Ignore storage failures.
    }
  }

  function loadPendingRun() {
    try {
      const raw = sessionStorage.getItem(PENDING_RUN_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      const ids = Array.isArray(parsed?.ids) ? parsed.ids.map(x => `${x}`) : [];
      if (!ids.length) return null;
      return { ids, createdAt: Number(parsed?.createdAt || 0) };
    } catch {
      return null;
    }
  }

  function clearPendingRun() {
    try {
      sessionStorage.removeItem(PENDING_RUN_KEY);
    } catch {
      // Ignore storage failures.
    }
  }

  function renderRunList(tests) {
    const q = (document.getElementById('searchBox').value || '').toLowerCase();
    const pillar = document.getElementById('pillarSelect').value;
    const tenantType = document.getElementById('tenantTypeSelect').value;
    const risk = document.getElementById('riskSelect').value;
    const impl = document.getElementById('implementationSelect').value;

    const filtered = tests.filter(t => {
      if (pillar && pillar !== 'All' && !(t.pillars && t.pillars.length ? t.pillars : [t.pillar]).includes(pillar)) return false;
      if (tenantType && !(t.tenantType || []).includes(tenantType)) return false;
      if (risk && t.risk !== risk) return false;
      if (impl === 'full' && t.implementationLevel !== 'full') return false;
      if (impl === 'partial' && t.implementationLevel !== 'partial') return false;
      if (impl === 'none' && t.implementationLevel !== 'none') return false;
      if (!q) return true;
      return `${t.id} ${t.title} ${t.category} ${t.pillar}`.toLowerCase().includes(q);
    });

    const selected = new Set((window.__selectedTests || []).map(x => `${x}`));
    const html = filtered.map(t => {
      const checked = selected.has(t.id) ? 'checked' : '';
      const partialNotes = t.implementationLevel === 'partial' && t.implementationNotes
        ? `<div class="partial-notes"><div><strong>Implemented:</strong> ${t.implementationNotes.implemented || '-'}</div><div><strong>Not implemented:</strong> ${t.implementationNotes.notImplemented || '-'}</div></div>`
        : '';
      const docs = t.documentationUrl
        ? `<div class="test-docs"><a href="${t.documentationUrl}" target="_blank" rel="noopener noreferrer">Docs root</a> · <a href="${t.documentationSearchUrl}" target="_blank" rel="noopener noreferrer">Search docs</a></div>`
        : '';
      return `<label class="test-item"><input type="checkbox" data-id="${t.id}" ${checked}/><div><div><strong>${t.id}</strong> ${t.title}</div><div class="test-meta">${t.pillar} | ${t.category} | risk ${t.risk}</div>${partialNotes}${docs}</div><div>${statusBadge(t.implementationLevel)}</div></label>`;
    }).join('');

    document.getElementById('testList').innerHTML = html || '<div class="test-item">No tests match this filter.</div>';
    const fullCount = filtered.filter(x => x.implementationLevel === 'full').length;
    const partialCount = filtered.filter(x => x.implementationLevel === 'partial').length;
    const plannedCount = filtered.filter(x => x.implementationLevel === 'none').length;
    document.getElementById('runStats').textContent = `${filtered.length} visible of ${tests.length} total controls | full ${fullCount} | partial ${partialCount} | planned ${plannedCount}`;

    document.querySelectorAll('#testList input[type="checkbox"]').forEach(cb => {
      cb.addEventListener('change', () => {
        const id = cb.dataset.id;
        const s = new Set(window.__selectedTests || []);
        if (cb.checked) s.add(id); else s.delete(id);
        window.__selectedTests = Array.from(s);
      });
    });

    window.__visibleTests = filtered.map(x => x.id);
  }

  async function initRunUi() {
    const tests = await Registry.getAll();
    window.__allTests = tests;

    const remembered = Store.loadSelection();
    window.__selectedTests = remembered.length ? remembered.filter(id => tests.some(t => t.id === id)) : tests.map(t => t.id);

    const refresh = () => renderRunList(tests);
    ['searchBox', 'pillarSelect', 'tenantTypeSelect', 'riskSelect', 'implementationSelect'].forEach(id => {
      document.getElementById(id).addEventListener('input', refresh);
      document.getElementById(id).addEventListener('change', refresh);
    });

    document.getElementById('selectVisibleBtn').addEventListener('click', () => {
      window.__selectedTests = Array.from(new Set([...(window.__selectedTests || []), ...(window.__visibleTests || [])]));
      renderRunList(tests);
    });

    document.getElementById('clearSelectionBtn').addEventListener('click', () => {
      window.__selectedTests = [];
      renderRunList(tests);
    });

    // Scan profiles: built-in selections plus user-saved ones (stored locally).
    const BUILTIN_PROFILES = {
      'Quick wins (high risk, low effort)': t => t.risk === 'High' && t.implementationCost === 'Low' && t.implemented,
      'High risk only': t => t.risk === 'High',
      'Identity baseline': t => (t.pillars || [t.pillar]).includes('Identity') && t.implemented,
      'Implemented controls only': t => t.implemented,
    };

    function refreshProfilePicker() {
      const picker = document.getElementById('profileSelect');
      const saved = Object.keys(Store.loadProfiles()).sort();
      picker.innerHTML = '<option value="">Profiles...</option>'
        + Object.keys(BUILTIN_PROFILES).map(n => `<option value="builtin:${n}">${n}</option>`).join('')
        + saved.map(n => `<option value="saved:${n}">★ ${n}</option>`).join('');
    }

    document.getElementById('profileSelect').addEventListener('change', e => {
      const value = e.target.value;
      if (!value) return;
      if (value.startsWith('builtin:')) {
        const fn = BUILTIN_PROFILES[value.slice(8)];
        if (fn) window.__selectedTests = tests.filter(fn).map(t => t.id);
      } else if (value.startsWith('saved:')) {
        const ids = Store.loadProfiles()[value.slice(6)] || [];
        window.__selectedTests = ids.filter(id => tests.some(t => t.id === id));
      }
      renderRunList(tests);
      toast(`Profile applied: ${(window.__selectedTests || []).length} controls selected.`);
      e.target.value = '';
    });

    document.getElementById('saveProfileBtn').addEventListener('click', () => {
      const selected = window.__selectedTests || [];
      if (!selected.length) {
        toast('Select at least one control before saving a profile.');
        return;
      }
      const name = prompt(`Save ${selected.length} selected controls as profile:`);
      if (!name || !name.trim()) return;
      Store.saveProfile(name.trim(), selected);
      refreshProfilePicker();
      toast(`Profile "${name.trim()}" saved.`);
    });

    refreshProfilePicker();

    async function runSelectedTests(selected, options) {
      const opts = options || {};
      if (!selected.length) {
        toast('Select at least one test.');
        return;
      }

      if (!Auth.getAccount()) {
        if (!opts.resume) savePendingRun(selected.map(t => t.id));
        toast('Signing in first. Scan will continue automatically after redirect.');
        await Auth.signIn();
        return;
      }

      Store.rememberSelection(selected.map(t => t.id));
      const scopes = Array.from(new Set(selected.flatMap(t => t.requiredScopes || Auth.coreScopes)));

      if (!opts.resume) {
        savePendingRun(selected.map(t => t.id));
      } else {
        // One-shot resume: clear pending before scope check so a second redirect
        // cannot trap the app in a refresh/consent loop.
        clearPendingRun();
      }
      let ok = false;
      try {
        ok = await Auth.ensureScopes(scopes, { forceConsent: false });
      } catch (scopeErr) {
        Log.err('Scope preflight failed', scopeErr?.message || `${scopeErr}`);
        clearPendingRun();
        toast(`Could not acquire required scopes: ${scopeErr?.message || scopeErr}`);
        return;
      }
      if (!ok) {
        if (opts.resume) {
          toast('Consent still required. Click Run selected again after consent completes.');
        } else {
          toast('Completing consent/sign-in. Scan will start automatically after redirect.');
        }
        return;
      }

      clearPendingRun();

      document.getElementById('progressCard').classList.remove('hidden');
      document.getElementById('cancelBtn').classList.remove('hidden');
      document.getElementById('runBtn').disabled = true;
      document.getElementById('progressLog').textContent = '';

      try {
        const scan = await Runner.run(selected, {
          collectTenantInfo: document.getElementById('tenantInfoToggle')?.checked !== false,
          tenantInfoDays: Number(document.getElementById('signInDaysSelect')?.value) || 7,
          progress: ({ done, total, current, elapsedTotalMs, elapsedCurrentMs, etaMs, currentEstimateMs, subProgress }) => {
            const pct = total ? Math.round((100 * done) / total) : 0;
            document.getElementById('progressFill').style.width = `${pct}%`;
            document.getElementById('progressText').textContent = `${done}/${total} ${current ? `${current.id} ${current.title}` : ''}`;
            document.getElementById('progressMeta').textContent = `Elapsed ${fmtDuration(elapsedTotalMs)} | Est. remaining ${fmtDuration(etaMs)}`;
            const testInfo = current
              ? `Test elapsed ${fmtDuration(elapsedCurrentMs)}${currentEstimateMs ? ` / est ${fmtDuration(currentEstimateMs)}` : ''}`
              : '';
            const subInfo = subProgress?.message
              ? `${subProgress.message}${Number.isFinite(subProgress.percent) ? ` (${Math.round(subProgress.percent)}%)` : ''}`
              : '';
            document.getElementById('subTestProgress').textContent = [testInfo, subInfo].filter(Boolean).join(' | ');
          },
          log: (line) => {
            const el = document.getElementById('progressLog');
            el.textContent += `${line}\n`;
            el.scrollTop = el.scrollHeight;
            Log.info(line);
          },
        });

        UIResults.refreshPicker(scan.id);
        UIHistory.refresh();
        if (scan?.persistence?.saved === false) {
          toast('Scan completed, but results could not be stored locally due to browser storage limits. Consider exporting and clearing history.');
        } else if (scan?.persistence?.warning) {
          toast(`Scan complete with storage cleanup: ${scan.persistence.warning}`);
        } else {
          toast(`Scan complete: ${scan.summary.passed} passed, ${scan.summary.failed} failed.`);
        }

        UIDashboard.refresh();
        document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x.dataset.tab === 'dashboard'));
        document.querySelectorAll('.panel').forEach(p => p.classList.toggle('active', p.id === 'panel-dashboard'));
      } catch (e) {
        Log.err('Run failed', e.message || `${e}`);
        toast(`Run failed: ${e.message || e}`);
      } finally {
        document.getElementById('runBtn').disabled = false;
        document.getElementById('cancelBtn').classList.add('hidden');
      }
    }

    document.getElementById('runBtn').addEventListener('click', async () => {
      const selectedIds = new Set(window.__selectedTests || []);
      const selected = tests.filter(t => selectedIds.has(t.id));
      await runSelectedTests(selected);
    });

    document.getElementById('cancelBtn').addEventListener('click', () => Runner.cancel());

    renderRunList(tests);

    const implCount = tests.filter(t => t.implemented).length;
    const partial = tests.filter(t => t.implementationLevel === 'partial');
    const docsCount = tests.filter(t => !!t.documentationUrl).length;
    const partialRows = partial.length
      ? `<table class="results-table"><thead><tr><th>Test</th><th>What is implemented</th><th>What is not implemented</th></tr></thead><tbody>${partial.map(t => `<tr><td>${t.id} ${t.title}</td><td>${t.implementationNotes?.implemented || '-'}</td><td>${t.implementationNotes?.notImplemented || '-'}</td></tr>`).join('')}</tbody></table>`
      : '<div class="muted">No partial implementations registered.</div>';

    document.getElementById('catalogStats').innerHTML = `
      <p><strong>${tests.length}</strong> controls loaded from Microsoft metadata.</p>
      <p><strong>${implCount}</strong> implemented (${tests.filter(t => t.implementationLevel === 'full').length} full, ${partial.length} partial), ${tests.filter(t => t.implementationLevel === 'none').length} planned.</p>
      <p><strong>${docsCount}</strong> controls with supporting documentation links.</p>
      <h3>Partial implementation details</h3>
      ${partialRows}
    `;

    // Resume a pending run after redirect-based consent/sign-in.
    const pending = loadPendingRun();
    if (pending && pending.ids.length) {
      const pendingSelected = tests.filter(t => pending.ids.includes(t.id));
      if (pendingSelected.length) {
        window.__selectedTests = pendingSelected.map(t => t.id);
        renderRunList(tests);
        toast('Resuming scan after sign-in/consent...');
        await runSelectedTests(pendingSelected, { resume: true });
      } else {
        clearPendingRun();
      }
    }
  }

  async function boot() {
    initTheme();
    initTabs();
    UIResults.init();
    UIHistory.init();

    document.getElementById('clearLogBtn').addEventListener('click', () => Log.clear());
    document.getElementById('toolkitBtn').addEventListener('click', () => {
      window.location.href = '../index.html';
    });

    document.getElementById('signInBtn').addEventListener('click', async () => {
      try {
        await Auth.signIn();
      } catch (e) {
        Log.err('Sign-in failed', e.message || `${e}`);
        toast(`Sign-in failed: ${e.message || e}`);
      }
    });

    document.getElementById('signOutBtn').addEventListener('click', async () => {
      try { await Auth.signOut(); }
      catch (e) { Log.warn('Sign-out warning', e.message || `${e}`); }
    });

    try {
      const account = await Auth.init();
      const redirectErr = Auth.getLastRedirectError?.();
      if (redirectErr) {
        Log.err('Auth redirect error', redirectErr?.message || `${redirectErr}`);
        toast(`Sign-in/consent error: ${redirectErr?.message || redirectErr}`);
      }
      if (account) {
        document.getElementById('connectCard').classList.add('hidden');
        document.getElementById('runCard').classList.remove('hidden');
        document.getElementById('userBox').classList.remove('hidden');
        document.getElementById('userText').textContent = account.username || account.name || 'Signed in';
      }
      await initRunUi();
      UIResults.refreshPicker();
      UIHistory.refresh();
    } catch (e) {
      Log.err('Startup error', e.message || `${e}`);
      toast(`Startup error: ${e.message || e}`);
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
