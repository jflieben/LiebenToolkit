// SimpleMaester entry point: sign-in, tabs, the run flow and the scan picker.
(() => {
  const { toast, busy, setUser } = Toolkit;
  const PENDING_RUN_KEY = 'SimpleMaester:pendingRun';
  const $ = (id) => document.getElementById(id);
  let activateTab = () => {};
  let running = false;

  // A consent redirect (popup blocked) reloads the page; this lets the run resume.
  function savePendingRun(tests, label) {
    try {
      const testIds = (tests || []).map((t) => t && t.id).filter(Boolean);
      if (!testIds.length) return;
      sessionStorage.setItem(PENDING_RUN_KEY, JSON.stringify({ label: label || 'Selection', testIds, at: Date.now() }));
    } catch (e) { Log.warn('Could not store pending run state:', e.message); }
  }
  function loadPendingRun() {
    try {
      const parsed = JSON.parse(sessionStorage.getItem(PENDING_RUN_KEY) || 'null');
      if (!parsed || !Array.isArray(parsed.testIds) || !parsed.testIds.length) return null;
      if (typeof parsed.at === 'number' && Date.now() - parsed.at > 60 * 60 * 1000) return null;
      return parsed;
    } catch { return null; }
  }
  function clearPendingRun() {
    try { sessionStorage.removeItem(PENDING_RUN_KEY); } catch { /* storage blocked */ }
  }

  async function tryResumePendingRun() {
    const pending = loadPendingRun();
    if (!pending) return;
    clearPendingRun();
    const byId = new Map((await Registry.getAll()).map((t) => [t.id, t]));
    const tests = pending.testIds.map((id) => byId.get(id)).filter((t) => t && t.implemented);
    if (!tests.length) return;
    toast(`Resuming the run of ${tests.length} test(s) after consent`, 'info');
    await runTests(tests, pending.label);
  }

  async function refreshScanPicker(selectId) {
    const sel = $('scanPicker');
    const all = await History.getAllScans();
    all.sort((a, b) => b.startedAt.localeCompare(a.startedAt));
    sel.replaceChildren(...(all.length
      ? all.map((s) => new Option(`${s.tenantDomain || s.tenantId} - ${new Date(s.startedAt).toLocaleString()} (${s.summary.failed} failed, ${s.summary.passed} passed)`, s.id))
      : [new Option('No scans', '')]));
    if (!all.length) { UIResults.hide(); return; }
    sel.value = selectId && all.some((s) => s.id === selectId) ? selectId : all[0].id;
    const scan = all.find((s) => s.id === sel.value);
    if (scan) UIResults.show(scan);
  }

  function refreshAuthUI() {
    const signedIn = !!Auth.getAccount();
    setUser(Auth.userInfo(onSignOut));
    $('connectCard').hidden = signedIn;
    $('runWorkspace').hidden = !signedIn;
  }

  async function onSignIn() {
    try { await Auth.signIn('read'); }
    catch (e) { Log.err('Sign-in failed:', e); toast(`Sign-in failed: ${e.message}`, 'err'); }
  }
  async function onSignOut() {
    if (running) Runner.cancel();
    await Auth.signOut();
  }

  async function startRun() {
    const tests = UICatalog.getSelectedTests();
    if (!tests.length) return;
    UICatalog.recordCustomRun();
    await runTests(tests, 'Custom selection');
  }

  function setRunning(on) {
    running = on;
    $('progressCard').hidden = !on;
    $('categoryPickerCard').hidden = on;
    if (on) $('customPickerCard').hidden = true;
    $('runBtn').disabled = on;
  }

  async function runTests(tests, label) {
    if (!tests || !tests.length) { toast('No tests to run.', 'warn'); return; }
    if (running) { toast('A scan is already running.', 'warn'); return; }

    // Ask every consent the selection needs up front so prompts do not interrupt the scan.
    savePendingRun(tests, label);
    try {
      await Auth.ensureScopes(tests.map((t) => t.requiredScopes || 'read'));
    } catch (e) {
      clearPendingRun();
      Log.err('Permission check failed:', e.message);
      toast(`Could not start the scan: ${e.message}`, 'err');
      return;
    }
    clearPendingRun();

    setRunning(true);
    const fill = $('progressFill');
    const subFill = $('subProgressFill');
    const log = $('progressLog');
    $('subProgressWrap').hidden = true;
    subFill.style.width = '0%';
    fill.style.width = '0%';
    $('progressLabel').textContent = `Running ${label || 'selection'}`;
    $('progressCount').textContent = `0 / ${tests.length}`;
    $('progressEta').textContent = '';
    log.textContent = '';
    const t0 = Date.now();

    const target = new EventTarget();
    target.addEventListener('progress', (e) => {
      const { done, total, current, sub } = e.detail;
      fill.style.width = `${Math.round(100 * done / total)}%`;
      $('progressCount').textContent = `${done} / ${total}${current ? ` - ${current.id}` : ''}`;
      $('progressEta').textContent = Concurrency.eta(t0, done, total);
      if (sub && typeof sub.total === 'number' && sub.total > 0) {
        $('subProgressWrap').hidden = false;
        subFill.style.width = `${Math.max(0, Math.min(100, Math.round(100 * (sub.done || 0) / sub.total)))}%`;
        $('subProgressText').textContent = `${sub.label || 'Current test'} (${sub.done || 0}/${sub.total})`;
      } else {
        $('subProgressWrap').hidden = true;
      }
    });
    target.addEventListener('log', (e) => {
      const { level, text } = e.detail;
      if (level === 'debug' && !Log.isVerbose()) return;
      log.textContent += `${text}\n`;
      log.scrollTop = log.scrollHeight;
    });
    target.addEventListener('done', (e) => {
      const { scan } = e.detail;
      Log.info(`Scan done, saved as ${scan.id}`);
      refreshScanPicker(scan.id);
      activateTab('results');
      toast(`Scan complete: ${scan.summary.passed} passed, ${scan.summary.failed} failed.`, 'ok');
    });

    try {
      await Runner.run(tests, { target });
    } catch (e) {
      Log.err('Scan failed:', e.message);
      toast(`Scan failed: ${e.message}`, 'err');
    } finally {
      setRunning(false);
    }
  }

  function initButtons() {
    $('signInBtn').addEventListener('click', (e) => busy(e.currentTarget, onSignIn));
    $('runBtn').addEventListener('click', startRun);
    $('cancelRunBtn').addEventListener('click', () => { Runner.cancel(); Log.warn('Cancel requested'); toast('Cancelling after the current test', 'info'); });
    $('scanPicker').addEventListener('change', async () => {
      const scan = await History.getScan($('scanPicker').value);
      if (scan) UIResults.show(scan);
    });
    document.body.addEventListener('click', (e) => {
      const goto = e.target.closest('[data-goto]');
      if (goto) { e.preventDefault(); activateTab(goto.dataset.goto); }
    });
    document.addEventListener('simplemaester:history-cleared', () => refreshScanPicker());
  }

  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    activateTab = Toolkit.tabs('#tabs', { onChange: (id) => { if (id === 'history') UIHistory.refresh(); } });
    initButtons();
    UIResults.init();
    UIHistory.init();
    refreshScanPicker().catch((e) => Log.warn('Could not read scan history:', e.message));
    const catalog = UICatalog.init().catch((e) => {
      Log.err('Could not load the test catalog:', e.message);
      toast(`Could not load the test catalog: ${e.message}`, 'err');
      return false;
    });

    try {
      await Auth.init();
    } catch (e) {
      Log.err('Auth init failed:', e.message);
      $('authError').textContent = `Sign-in is not available: ${e.message}`;
      $('authError').hidden = false;
    }
    refreshAuthUI();
    if (!Auth.getAccount() || (await catalog) === false) return;
    await tryResumePendingRun();
  }

  document.addEventListener('DOMContentLoaded', boot);
  window.App = { runTests };
})();
