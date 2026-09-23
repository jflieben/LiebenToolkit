// SimpleMaester app entry point. Wires the UI together, drives the runner and
// keeps the tabs/theme/sign-in plumbing in one place.
(() => {
  const THEME_KEY = 'm365tk-theme';
  const PENDING_RUN_KEY = 'SimpleMaester:pendingRun';

  function savePendingRun(tests, label) {
    try {
      const testIds = (tests || []).map(t => t?.id).filter(Boolean);
      if (!testIds.length) return;
      sessionStorage.setItem(PENDING_RUN_KEY, JSON.stringify({
        label: label || 'Selection',
        testIds,
        at: Date.now(),
      }));
    } catch (e) {
      Log.warn('Could not store pending run state:', e.message);
    }
  }

  function loadPendingRun() {
    try {
      const raw = sessionStorage.getItem(PENDING_RUN_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || !Array.isArray(parsed.testIds) || !parsed.testIds.length) return null;
      // Ignore stale run intents from old sessions.
      if (typeof parsed.at === 'number' && (Date.now() - parsed.at) > (60 * 60 * 1000)) return null;
      return parsed;
    } catch {
      return null;
    }
  }

  function clearPendingRun() {
    try { sessionStorage.removeItem(PENDING_RUN_KEY); }
    catch {}
  }

  async function tryResumePendingRun() {
    const pending = loadPendingRun();
    if (!pending) return false;
    const all = await Registry.getAll();
    const byId = new Map(all.map(t => [t.id, t]));
    const tests = pending.testIds.map(id => byId.get(id)).filter(t => t && t.implemented);
    if (!tests.length) {
      clearPendingRun();
      return false;
    }
    toast(`Resuming pending run (${tests.length} tests) after sign-in/consent...`);
    await runTests(tests, pending.label || 'Selection');
    return true;
  }

  function applyTheme(t) {
    document.documentElement.dataset.theme = t;
    document.getElementById('darkToggle').textContent = t === 'dark' ? '\u2600\uFE0F' : '\uD83C\uDF19';
  }
  function initTheme() {
    const t = localStorage.getItem(THEME_KEY) || 'light';
    applyTheme(t);
    document.getElementById('darkToggle').addEventListener('click', () => {
      const next = (document.documentElement.dataset.theme === 'dark') ? 'light' : 'dark';
      localStorage.setItem(THEME_KEY, next);
      applyTheme(next);
    });
  }

  function initTabs() {
    document.querySelectorAll('.top-tab').forEach(b => b.addEventListener('click', () => switchTab(b.dataset.top)));
    document.body.addEventListener('click', e => {
      const goto = e.target.closest('[data-goto]');
      if (goto) { e.preventDefault(); switchTab(goto.dataset.goto); }
    });
  }
  function switchTab(id) {
    document.querySelectorAll('.top-tab').forEach(b => b.classList.toggle('active', b.dataset.top === id));
    document.querySelectorAll('.top-panel').forEach(p => p.classList.toggle('active', p.id === ('top-' + id)));
    if (id === 'history') UIHistory.refresh();
  }

  function showUserBox(account) {
    const box = document.getElementById('userBox');
    if (!account) { box.classList.add('hidden'); return; }
    box.classList.remove('hidden');
    document.getElementById('userName').textContent = account.username || account.name || '(signed in)';
  }

  async function refreshScanPicker(selectId) {
    const sel = document.getElementById('scanPicker');
    const all = await History.getAllScans();
    all.sort((a, b) => b.startedAt.localeCompare(a.startedAt));
    sel.innerHTML = all.length
      ? all.map(s => `<option value="${s.id}">${(s.tenantDomain||s.tenantId)} - ${new Date(s.startedAt).toLocaleString()} (${s.summary.failed}F/${s.summary.passed}P)</option>`).join('')
      : '<option value="">(no scans)</option>';
    if (selectId) sel.value = selectId;
    if (sel.value) {
      const scan = all.find(s => s.id === sel.value);
      if (scan) UIResults.show(scan);
    }
    sel.onchange = () => {
      const scan = all.find(s => s.id === sel.value);
      if (scan) UIResults.show(scan);
    };
  }

  async function showLatestIfAny() {
    const all = await History.getAllScans();
    if (!all.length) return;
    all.sort((a, b) => b.startedAt.localeCompare(a.startedAt));
    refreshScanPicker(all[0].id);
  }

  async function onSignedIn(account) {
    showUserBox(account);
    document.getElementById('connectCard').classList.add('hidden');
    document.getElementById('runWorkspace').classList.remove('hidden');
    Log.info('Signed in as ' + (account.username || account.name));
    await UICatalog.init();
    showLatestIfAny();
  }

  async function startRun() {
    const tests = UICatalog.getSelectedTests();
    if (!tests.length) return;
    UICatalog.recordCustomRun();
    return runTests(tests, 'Custom selection');
  }

  async function runTests(tests, label) {
    if (!tests || !tests.length) { toast('No tests to run.'); return; }

    // Persist intent so redirect-based login/consent can resume this run.
    savePendingRun(tests, label);

    // Pre-consent every scope needed by the chosen test set to avoid any
    // interactive prompts in the middle of a scan.
    const requiredScopes = Array.from(new Set(
      tests.flatMap(t => (Array.isArray(t.requiredScopes) && t.requiredScopes.length)
        ? t.requiredScopes
        : Auth.SCOPES.graphFull)
    ));
    try {
      const ok = await Auth.ensureScopes(requiredScopes);
      if (!ok) return; // redirect flow started for consent
    } catch (e) {
      clearPendingRun();
      Log.err('Permission preflight failed:', e.message);
      toast('Could not start scan: ' + e.message);
      return;
    }

    // Consent succeeded in-page; no redirect resume is needed anymore.
    clearPendingRun();

    document.getElementById('categoryPickerCard').classList.add('hidden');
    document.getElementById('customPickerCard').classList.add('hidden');
    document.getElementById('progressCard').classList.remove('hidden');
    const fill = document.getElementById('progressFill');
    const txt  = document.getElementById('progressText');
    const subWrap = document.getElementById('subProgressWrap');
    const subFill = document.getElementById('subProgressFill');
    const subTxt  = document.getElementById('subProgressText');
    const log  = document.getElementById('progressLog');
    subWrap.classList.add('hidden');
    subFill.style.width = '0%';
    subTxt.textContent = '';
    fill.style.width = '0%'; txt.textContent = `0 / ${tests.length} - ${label||''}`; log.textContent = '';
    document.getElementById('runBtn').disabled = true;
    document.getElementById('cancelRunBtn').classList.remove('hidden');

    const target = new EventTarget();
    target.addEventListener('progress', e => {
      const { done, total, current, sub } = e.detail;
      const pct = Math.round(100 * done / total);
      fill.style.width = pct + '%';
      txt.textContent = `${done} / ${total} - ${current?.id || ''}`;
      if (sub && typeof sub.total === 'number' && sub.total > 0) {
        const subPct = Math.max(0, Math.min(100, Math.round(100 * (sub.done || 0) / sub.total)));
        subWrap.classList.remove('hidden');
        subFill.style.width = subPct + '%';
        subTxt.textContent = `${sub.label || 'Current test'} (${sub.done || 0}/${sub.total})`;
      } else {
        subWrap.classList.add('hidden');
        subFill.style.width = '0%';
        subTxt.textContent = '';
      }
    });
    target.addEventListener('log', e => {
      const { level, text } = e.detail;
      const verbose = document.getElementById('verboseLog').checked;
      if (level === 'debug' && !verbose) return;
      log.textContent += text + '\n';
      log.scrollTop = log.scrollHeight;
    });
    target.addEventListener('done', e => {
      Log.info(`Scan done. Saved as ${e.detail.scan.id}`);
      document.getElementById('runBtn').disabled = false;
      document.getElementById('cancelRunBtn').classList.add('hidden');
      document.getElementById('progressCard').classList.add('hidden');
      document.getElementById('categoryPickerCard').classList.remove('hidden');
      clearPendingRun();
      refreshScanPicker(e.detail.scan.id);
      switchTab('results');
      toast(`Scan complete: ${e.detail.scan.summary.passed} passed, ${e.detail.scan.summary.failed} failed.`);
    });

    try {
      await Runner.run(tests, { concurrency: 4, target });
    } catch (e) {
      Log.err('Scan failed:', e.message);
      document.getElementById('runBtn').disabled = false;
      document.getElementById('cancelRunBtn').classList.add('hidden');
      toast('Scan failed: ' + e.message);
    }
  }

  function toast(msg) {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.classList.remove('hidden');
    setTimeout(() => el.classList.add('hidden'), 4000);
  }

  function initButtons() {
    document.getElementById('signInBtn').addEventListener('click', async () => {
      try { const a = await Auth.signIn(); if (a) await onSignedIn(a); }
      catch (e) { Log.err('Sign-in failed:', e.message); toast('Sign-in failed: ' + e.message); }
    });
    document.getElementById('signOutBtn').addEventListener('click', async () => {
      try { await Auth.signOut(); } catch (e) { Log.warn(e.message); }
      location.reload();
    });
    document.getElementById('runBtn').addEventListener('click', startRun);
    document.getElementById('cancelRunBtn').addEventListener('click', () => { Runner.cancel(); Log.warn('Cancel requested'); });
    const cancel2 = document.getElementById('cancelRunBtn2');
    if (cancel2) cancel2.addEventListener('click', () => { Runner.cancel(); Log.warn('Cancel requested'); });
    document.getElementById('clearLogBtn').addEventListener('click', () => Log.clear());
  }

  async function boot() {
    initTheme();
    initTabs();
    initButtons();
    UIResults.init();
    UIHistory.init();

    try {
      const acc = await Auth.init();
      if (acc) {
        await onSignedIn(acc);
        await tryResumePendingRun();
      }
    } catch (e) {
      Log.err('Auth init failed:', e.message);
      const card = document.getElementById('connectCard');
      const warn = document.createElement('p');
      warn.className = 'error-msg';
      warn.textContent = 'Auth init failed: ' + e.message;
      card.appendChild(warn);
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
  window.App = { runTests, toast };
})();
