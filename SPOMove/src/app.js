// Wizard orchestration, location pickers, job dashboard and live progress.
(() => {
  // Popup callback: MSAL reads the response from the opener, so don't boot the app.
  if (window.Auth && Auth.isAuthPopup()) { return; }

  const $ = (id) => document.getElementById(id);
  const fmtBytes = (n) => window.Transfer.fmtBytes(n);

  const STEP_LABELS = {
    mode: 'Type', sourceConnect: 'Source login', sourceLoc: 'Source',
    targetConnect: 'Target login', targetLoc: 'Target', options: 'Options', review: 'Review',
  };

  const wiz = {
    step: 0,
    mode: null,
    sourcePicker: null,
    targetPicker: null,
    options: { operation: 'copy', includeVersions: false, includeMetadata: true, overwrite: false, onlyNewer: false, concurrency: 4 },
  };

  function steps() {
    return wiz.mode === 'cross'
      ? ['mode', 'sourceConnect', 'sourceLoc', 'targetConnect', 'targetLoc', 'options', 'review']
      : ['mode', 'sourceConnect', 'sourceLoc', 'targetLoc', 'options', 'review'];
  }

  // ---------- toast ----------
  let toastTimer = null;
  function toast(msg, ms = 4500) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.remove('hidden');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.add('hidden'), ms);
  }

  // ---------- tabs / theme / log ----------
  function initChrome() {
    document.querySelectorAll('.tab').forEach(t => {
      t.addEventListener('click', () => activateTab(t.dataset.tab));
    });
    if (localStorage.getItem('SPOMove-theme') === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
    $('darkToggle').addEventListener('click', () => {
      const dark = document.documentElement.getAttribute('data-theme') === 'dark';
      if (dark) { document.documentElement.removeAttribute('data-theme'); localStorage.setItem('SPOMove-theme', ''); }
      else { document.documentElement.setAttribute('data-theme', 'dark'); localStorage.setItem('SPOMove-theme', 'dark'); }
    });
    $('verboseLog').addEventListener('change', e => Log.setVerbose(e.target.checked));
    $('clearLogBtn').addEventListener('click', () => Log.clear());
    $('signOutBtn').addEventListener('click', async () => {
      await Auth.signOutAll();
      updateUserBox();
      toast('Signed out of all tenants.');
    });
    $('newJobBtn').addEventListener('click', () => { resetWizard(); activateTab('wizard'); });
    $('clearFinishedBtn').addEventListener('click', () => { Jobs.clearFinished(); renderJobs(); });
    $('reportModal').addEventListener('click', (e) => { if (e.target.dataset && e.target.dataset.close) closeReport(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('reportModal').classList.contains('hidden')) closeReport(); });
  }

  function activateTab(name) {
    document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x.dataset.tab === name));
    document.querySelectorAll('.tab-panel').forEach(x => x.classList.toggle('active', x.id === `tab-${name}`));
    if (name === 'jobs') renderJobs();
  }

  function updateUserBox() {
    const s = Auth.getConnection('source');
    const t = Auth.getConnection('target');
    const parts = [];
    if (s?.upn) parts.push(`src: ${s.upn}`);
    if (t?.upn) parts.push(`tgt: ${t.upn}`);
    const box = $('userBox');
    if (parts.length) { box.classList.remove('hidden'); $('userName').textContent = parts.join('  ·  '); }
    else box.classList.add('hidden');
  }

  // ---------- wizard ----------
  function initWizard() {
    document.querySelectorAll('.choice-card').forEach(c => {
      c.addEventListener('click', () => {
        wiz.mode = c.dataset.mode;
        document.querySelectorAll('.choice-card').forEach(x => x.classList.toggle('selected', x === c));
        updateWizard();
      });
    });
    $('sourceSignInBtn').addEventListener('click', () => doSignIn('source'));
    $('targetSignInBtn').addEventListener('click', () => doSignIn('target'));
    $('wizBack').addEventListener('click', () => { if (wiz.step > 0) { wiz.step--; updateWizard(); } });
    $('wizNext').addEventListener('click', () => { if (canAdvance()) { wiz.step++; onEnterStep(); updateWizard(); } });
    $('wizStart').addEventListener('click', startTransfer);

    // options wiring
    document.querySelectorAll('input[name="operation"]').forEach(r =>
      r.addEventListener('change', e => { wiz.options.operation = e.target.value; refreshOptionWarning(); }));
    $('optMetadata').addEventListener('change', e => wiz.options.includeMetadata = e.target.checked);
    $('optVersions').addEventListener('change', e => { wiz.options.includeVersions = e.target.checked; refreshOptionWarning(); });
    $('optOverwrite').addEventListener('change', e => {
      wiz.options.overwrite = e.target.checked;
      const on = $('optOnlyNewer');
      on.disabled = !e.target.checked;
      if (!e.target.checked) { on.checked = false; wiz.options.onlyNewer = false; }
      refreshOptionWarning();
    });
    $('optOnlyNewer').addEventListener('change', e => wiz.options.onlyNewer = e.target.checked);
    $('optConcurrency').addEventListener('change', e => wiz.options.concurrency = Math.max(1, Math.min(12, Number(e.target.value) || 4)));

    resetWizard();
  }

  function resetWizard() {
    wiz.step = 0; wiz.mode = null;
    wiz.sourcePicker = null; wiz.targetPicker = null;
    document.querySelectorAll('.choice-card').forEach(x => x.classList.remove('selected'));
    $('sourcePicker').innerHTML = ''; $('targetPicker').innerHTML = '';
    wiz.options = { operation: 'copy', includeVersions: false, includeMetadata: true, overwrite: false, onlyNewer: false, concurrency: 4 };
    const op = document.querySelector('input[name="operation"][value="copy"]'); if (op) op.checked = true;
    $('optMetadata').checked = true; $('optVersions').checked = false; $('optOverwrite').checked = false; $('optConcurrency').value = 4;
    $('optOnlyNewer').checked = false; $('optOnlyNewer').disabled = true;
    updateWizard();
  }

  function onEnterStep() {
    const id = steps()[wiz.step];
    if (id === 'sourceLoc' && !wiz.sourcePicker) {
      wiz.sourcePicker = createPicker($('sourcePicker'), 'source', updateNavState);
    }
    if (id === 'targetLoc' && !wiz.targetPicker) {
      const role = wiz.mode === 'cross' ? 'target' : 'source';
      wiz.targetPicker = createPicker($('targetPicker'), role, updateNavState);
    }
    if (id === 'review') renderReview();
    if (id === 'options') refreshOptionWarning();
  }

  function renderStepper() {
    const list = steps();
    const stepper = $('stepper');
    stepper.innerHTML = '';
    list.forEach((s, i) => {
      const li = document.createElement('li');
      li.className = 'step' + (i === wiz.step ? ' active' : '') + (i < wiz.step ? ' done' : '');
      li.innerHTML = `<span class="step-dot">${i < wiz.step ? '✓' : i + 1}</span><span class="step-label">${STEP_LABELS[s]}</span>`;
      stepper.appendChild(li);
    });
  }

  function updateWizard() {
    renderStepper();
    const cur = steps()[wiz.step];
    document.querySelectorAll('.wiz-step').forEach(el => el.classList.toggle('hidden', el.dataset.step !== cur));
    // connection states
    renderConnState('source', $('sourceConnState'), $('sourceSignInBtn'));
    renderConnState('target', $('targetConnState'), $('targetSignInBtn'));
    updateNavState();
  }

  function renderConnState(role, stateEl, btn) {
    const c = Auth.getConnection(role);
    if (c?.account) {
      let html = `<div class="notice notice-info">Connected to <strong>${esc(c.tenantName || c.tenantHost || 'tenant')}</strong> as <strong>${esc(c.upn)}</strong>.</div>`;
      if (!c.spoConsent) {
        html += `<div class="notice notice-warn">This tenant hasn't approved the app's SharePoint access yet, which is required to browse and move files. Approve all SharePoint scopes in one prompt.
          <div class="row"><button class="conn-grantspo primary-btn" type="button">Grant SharePoint access</button></div></div>`;
      }
      stateEl.innerHTML = html;
      btn.textContent = 'Switch account';
      const gb = stateEl.querySelector('.conn-grantspo');
      if (gb) gb.addEventListener('click', () => grantSpo(role, gb));
    } else {
      stateEl.innerHTML = '';
      btn.textContent = role === 'source' ? 'Sign in to source tenant' : 'Sign in to target tenant';
    }
  }

  async function grantSpo(role, btn) {
    btn.disabled = true; btn.textContent = 'Requesting consent...';
    try {
      await Auth.grantSharePointConsent(role);
      toast('SharePoint access approved.');
      updateWizard();
      const cur = steps()[wiz.step];
      if ((cur === 'sourceConnect' || cur === 'targetConnect') && canAdvance()) { wiz.step++; onEnterStep(); updateWizard(); }
    } catch (e) {
      btn.disabled = false; btn.textContent = 'Grant SharePoint access';
      toast('Consent failed or was cancelled: ' + e.message);
    }
  }

  function updateNavState() {
    $('wizBack').disabled = wiz.step === 0;
    const cur = steps()[wiz.step];
    const isReview = cur === 'review';
    $('wizNext').classList.toggle('hidden', isReview);
    $('wizStart').classList.toggle('hidden', !isReview);
    $('wizNext').disabled = !canAdvance();
    if (isReview) renderReview();
    updateUserBox();
  }

  function canAdvance() {
    const cur = steps()[wiz.step];
    switch (cur) {
      case 'mode': return !!wiz.mode;
      case 'sourceConnect': return Auth.isConnected('source') && !!Auth.getConnection('source')?.spoConsent;
      case 'sourceLoc': return !!wiz.sourcePicker?.getSelection();
      case 'targetConnect': return Auth.isConnected('target') && !!Auth.getConnection('target')?.spoConsent;
      case 'targetLoc': return !!wiz.targetPicker?.getSelection();
      case 'options': return true;
      default: return false;
    }
  }

  async function doSignIn(role) {
    try {
      await Auth.signIn(role);
      updateUserBox();
      updateWizard();
      // auto-advance from a connect step once connected
      const cur = steps()[wiz.step];
      if ((cur === 'sourceConnect' || cur === 'targetConnect') && canAdvance()) { wiz.step++; onEnterStep(); updateWizard(); }
    } catch (e) {
      toast(e.message);
      Log.err('Sign-in failed:', e);
    }
  }

  function refreshOptionWarning() {
    const w = $('optWarning');
    const msgs = [];
    if (wiz.mode === 'cross') msgs.push('Cross-tenant transfers stream every file through your browser. Large libraries are limited by your bandwidth and memory - keep the browser tab open until the job completes.');
    if (wiz.options.includeVersions) msgs.push('Including previous versions multiplies the amount of data transferred and is markedly slower.');
    if (wiz.options.operation === 'move') msgs.push('Move deletes each source file after it transfers successfully. Verify a test batch first.');
    if (msgs.length) { w.classList.remove('hidden'); w.innerHTML = msgs.map(m => `<div>${esc(m)}</div>`).join(''); }
    else w.classList.add('hidden');
  }

  function renderReview() {
    const src = wiz.sourcePicker?.getSelection();
    const tgt = wiz.targetPicker?.getSelection();
    const sc = Auth.getConnection('source');
    const tc = Auth.getConnection(wiz.mode === 'cross' ? 'target' : 'source');
    const rows = [
      ['Type', wiz.mode === 'cross' ? 'Between two tenants' : 'Same tenant'],
      ['Operation', wiz.options.operation === 'move' ? 'Move (delete originals)' : 'Copy'],
      ['Source tenant', `${esc(sc?.tenantName || sc?.tenantHost || '')}`],
      ['Source', `${esc(src?.siteTitle || '')} › ${esc(src?.library || '')}${src?.subLabel ? ' › ' + esc(src.subLabel) : ''}`],
      ['Target tenant', `${esc(tc?.tenantName || tc?.tenantHost || '')}`],
      ['Target', `${esc(tgt?.siteTitle || '')} › ${esc(tgt?.library || '')}${tgt?.subLabel ? ' › ' + esc(tgt.subLabel) : ''}`],
      ['Metadata', wiz.options.includeMetadata ? 'Preserve original' : 'Reset to new'],
      ['Previous versions', wiz.options.includeVersions ? 'Included' : 'Current only'],
      ['Overwrite existing', wiz.options.overwrite ? (wiz.options.onlyNewer ? 'Yes, only if source newer' : 'Yes') : 'No'],
      ['Parallel files', String(wiz.options.concurrency)],
    ];
    $('reviewList').innerHTML = rows.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${v}</dd>`).join('');
    $('reviewWarning').textContent = 'SPOMove will first scan the source folder to build the file list, then start transferring. You can watch progress and pause, resume or cancel from the Jobs tab.';
  }

  function startTransfer() {
    const src = wiz.sourcePicker.getSelection();
    const tgt = wiz.targetPicker.getSelection();
    const sc = Auth.getConnection('source');
    const tc = Auth.getConnection(wiz.mode === 'cross' ? 'target' : 'source');
    const config = {
      mode: wiz.mode,
      operation: wiz.options.operation,
      includeVersions: wiz.options.includeVersions,
      includeMetadata: wiz.options.includeMetadata,
      overwrite: wiz.options.overwrite,
      onlyNewer: wiz.options.onlyNewer,
      concurrency: wiz.options.concurrency,
      source: { tenantId: sc.tenantId, tenantHost: sc.tenantHost, tenantName: sc.tenantName, upn: sc.upn, siteUrl: src.siteUrl, siteTitle: src.siteTitle, library: src.library, folder: src.folder },
      target: { tenantId: tc.tenantId, tenantHost: tc.tenantHost, tenantName: tc.tenantName, upn: tc.upn, siteUrl: tgt.siteUrl, siteTitle: tgt.siteTitle, library: tgt.library, folder: tgt.folder },
    };
    const job = Jobs.create(config);
    Log.info(`Created job ${job.id}`);
    resetWizard();
    activateTab('jobs');
    startJobRun(job.id);
  }

  // ---------- location picker ----------
  function createPicker(container, role, onChange) {
    const state = { siteUrl: null, siteTitle: null, library: null, libraries: [], folder: null, root: null };
    container.innerHTML = `
      <div class="picker-site">
        <div class="pk-scope">
          <button type="button" class="pk-scope-btn active" data-scope="sites">Sites</button>
          <button type="button" class="pk-scope-btn" data-scope="onedrive">OneDrive (by user)</button>
          <button type="button" class="pk-scope-btn" data-scope="all">All sites (admin)</button>
        </div>
        <div class="row">
          <input class="pk-search" type="search" placeholder="Search sites by name..." />
          <span class="muted">or</span>
          <input class="pk-url" type="text" placeholder="https://tenant.sharepoint.com/sites/Name" />
          <button class="pk-load ghost-btn" type="button">Load URL</button>
        </div>
        <div class="pk-scope-hint muted"></div>
        <div class="pk-results"></div>
        <div class="pk-siteinfo hidden"></div>
        <div class="pk-access hidden"></div>
      </div>
      <div class="picker-lib hidden">
        <label>Document library: <select class="pk-libsel"></select></label>
      </div>
      <div class="picker-folder hidden">
        <div class="pk-breadcrumb"></div>
        <div class="pk-folders"></div>
        <p class="table-hint">The highlighted folder above is the one that will be used. Click a subfolder to drill in.</p>
      </div>`;

    const q = (sel) => container.querySelector(sel);
    const searchEl = q('.pk-search'), urlEl = q('.pk-url'), resultsEl = q('.pk-results');
    const siteInfoEl = q('.pk-siteinfo'), accessEl = q('.pk-access');
    const libWrap = q('.picker-lib'), libSel = q('.pk-libsel');
    const folderWrap = q('.picker-folder'), crumbEl = q('.pk-breadcrumb'), foldersEl = q('.pk-folders');
    const hintEl = q('.pk-scope-hint');

    let scope = 'sites';
    let allSites = null;
    let allSitesNote = '';          // cached admin enumeration of every site collection
    let searchTimer = null;

    const SCOPE_UI = {
      sites:    { placeholder: 'Search sites by name...', minLen: 2, hint: '' },
      onedrive: { placeholder: 'Search a user by name or UPN...', minLen: 2, hint: 'Finds the selected user\u2019s OneDrive. You may be prompted once to consent to reading the directory.' },
      all:      { placeholder: 'Filter every site collection...', minLen: 0, hint: 'Lists every site in the tenant (requires the SharePoint Administrator role). Choose a site you cannot access and use Take ownership to elevate.' },
    };

    container.querySelectorAll('.pk-scope-btn').forEach(btn => {
      btn.addEventListener('click', () => setScope(btn.dataset.scope));
    });

    function setScope(next) {
      scope = next;
      container.querySelectorAll('.pk-scope-btn').forEach(b => b.classList.toggle('active', b.dataset.scope === next));
      const ui = SCOPE_UI[next];
      searchEl.value = '';
      searchEl.placeholder = ui.placeholder;
      hintEl.textContent = ui.hint;
      resultsEl.innerHTML = '';
      if (next === 'all') loadAllSites();
    }

    searchEl.addEventListener('input', () => {
      clearTimeout(searchTimer);
      const term = searchEl.value.trim();
      const minLen = SCOPE_UI[scope].minLen;
      if (term.length < minLen && scope !== 'all') { resultsEl.innerHTML = ''; return; }
      searchTimer = setTimeout(() => runSearch(term), scope === 'all' ? 120 : 350);
    });
    q('.pk-load').addEventListener('click', () => {
      const url = urlEl.value.trim();
      if (url) loadSite(url);
    });
    urlEl.addEventListener('keydown', e => { if (e.key === 'Enter') q('.pk-load').click(); });

    function runSearch(term) {
      if (scope === 'onedrive') return runUserSearch(term);
      if (scope === 'all') return renderAllSites(term);
      return runSiteSearch(term);
    }

    function siteResultRow(displayName, webUrl) {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'pk-result';
      row.innerHTML = `<strong>${esc(displayName)}</strong><span class="muted">${esc(webUrl)}</span>`;
      row.addEventListener('click', () => { resultsEl.innerHTML = ''; searchEl.value = ''; loadSite(webUrl); });
      return row;
    }

    async function runSiteSearch(term) {
      if (term.length < 2) return;
      resultsEl.innerHTML = '<div class="muted">Searching...</div>';
      try {
        const sites = await Graph.searchSites(role, term);
        if (!sites.length) { resultsEl.innerHTML = '<div class="muted">No sites found.</div>'; return; }
        resultsEl.innerHTML = '';
        sites.forEach(s => resultsEl.appendChild(siteResultRow(s.displayName, s.webUrl)));
      } catch (e) { resultsEl.innerHTML = `<div class="notice notice-warn">Search failed: ${esc(e.message)}</div>`; }
    }

    async function runUserSearch(term) {
      if (term.length < 2) return;
      resultsEl.innerHTML = '<div class="muted">Searching users...</div>';
      try {
        const users = await Graph.searchUsers(role, term);
        if (!users.length) { resultsEl.innerHTML = '<div class="muted">No users found.</div>'; return; }
        resultsEl.innerHTML = '';
        users.forEach(u => {
          const row = document.createElement('button');
          row.type = 'button';
          row.className = 'pk-result';
          row.innerHTML = `<strong>${esc(u.displayName)}</strong><span class="muted">${esc(u.userPrincipalName || u.mail || '')}</span>`;
          row.addEventListener('click', () => openUserOneDrive(u, row));
          resultsEl.appendChild(row);
        });
      } catch (e) {
        const msg = (e.status === 403) ? 'You do not have permission to look up users in this tenant.' : e.message;
        resultsEl.innerHTML = `<div class="notice notice-warn">User search failed: ${esc(msg)}</div>`;
      }
    }

    async function openUserOneDrive(user, row) {
      row.disabled = true;
      const orig = row.innerHTML;
      row.innerHTML = `<strong>${esc(user.displayName)}</strong><span class="muted">Resolving OneDrive...</span>`;
      try {
        const siteUrl = await Graph.getUserOneDriveSiteUrl(role, user);
        resultsEl.innerHTML = ''; searchEl.value = '';
        loadSite(siteUrl);
      } catch (e) {
        row.disabled = false; row.innerHTML = orig;
        resultsEl.insertAdjacentHTML('afterbegin', `<div class="notice notice-warn">${esc(e.message)}</div>`);
      }
    }

    async function loadAllSites() {
      if (allSites) { renderAllSites(searchEl.value.trim()); return; }
      resultsEl.innerHTML = '<div class="muted">Loading all site collections (admin)...</div>';
      try {
        allSites = await SharePoint.enumerateAllSites(role, {
          onProgress: (n) => { resultsEl.innerHTML = `<div class="muted">Loading all site collections (admin)... ${n} found</div>`; },
        });
        renderAllSites(searchEl.value.trim());
      } catch (e) {
        allSites = null;
        const denied = (e.status === 401 || e.status === 403);
        const conn = Auth.getConnection(role);
        // The tenant admin API needs the SharePoint-admin role + app consent. When it's denied,
        // fall back to a Graph enumeration of the sites visible to this account so the picker
        // still works (it just won't include sites the account can't see).
        if (denied && conn?.spoConsent) {
          try {
            resultsEl.innerHTML = '<div class="muted">Admin site list unavailable - listing sites visible to your account...</div>';
            const gsites = await Graph.listAllSites(role, (n) => { resultsEl.innerHTML = `<div class="muted">Listing sites visible to your account... ${n} found</div>`; });
            if (gsites.length) {
              allSites = gsites;
              allSitesNote = 'The tenant admin API is unavailable (it needs the SharePoint Administrator role and app consent), so this lists only the sites your account can see - sites you are not a member of may be missing.';
              renderAllSites(searchEl.value.trim());
              return;
            }
          } catch (ge) { Log.dbg('Graph site fallback failed: ' + ge.message); }
        }
        allSitesNote = '';
        // A 403 AFTER consent is granted is a missing SharePoint-admin role, not a consent issue —
        // don't re-offer the grant button (that caused an endless consent loop).
        const needConsent = denied && !conn?.spoConsent;
        const msg = !denied ? esc(e.message)
          : needConsent
            ? 'Listing every site needs the app to be granted SharePoint access in this tenant. Grant it and retry.'
            : `Listing every site requires the SharePoint (or Global) Administrator role. The account <strong>${esc(conn?.upn || '')}</strong> does not appear to have it (the tenant admin API returned ${e.status}). Switch to the <strong>Sites</strong> tab and paste a site URL, or sign in with an administrator account.`;
        resultsEl.innerHTML = `<div class="notice notice-warn">${msg}${needConsent ? '<div class="row"><button class="pk-grantspo primary-btn" type="button">Grant SharePoint access</button></div>' : ''}</div>`;
        const gb = resultsEl.querySelector('.pk-grantspo');
        if (gb) gb.addEventListener('click', async () => {
          gb.disabled = true; gb.textContent = 'Requesting consent...';
          try {
            await Auth.grantSharePointConsent(role);
            allSites = null; loadAllSites();
          } catch (err) {
            gb.disabled = false; gb.textContent = 'Grant SharePoint access';
            toast('Consent failed or was cancelled: ' + err.message);
          }
        });
      }
    }

    function renderAllSites(term) {
      if (!allSites) return;
      const t = (term || '').toLowerCase();
      const matched = t
        ? allSites.filter(s => s.title.toLowerCase().includes(t) || s.url.toLowerCase().includes(t))
        : allSites;
      if (!matched.length) { resultsEl.innerHTML = '<div class="muted">No sites match.</div>'; return; }
      const shown = matched.slice(0, 200);
      resultsEl.innerHTML = '';
      if (allSitesNote) resultsEl.insertAdjacentHTML('beforeend', `<div class="notice notice-info">${esc(allSitesNote)}</div>`);
      if (matched.length > shown.length) {
        resultsEl.insertAdjacentHTML('beforeend', `<div class="muted">${matched.length} sites - showing first ${shown.length}. Refine your filter.</div>`);
      }
      shown.forEach(s => resultsEl.appendChild(siteResultRow(s.title, s.url)));
    }

    async function loadSite(siteUrl) {
      accessEl.classList.add('hidden'); accessEl.innerHTML = '';
      libWrap.classList.add('hidden'); folderWrap.classList.add('hidden');
      siteInfoEl.classList.remove('hidden');
      siteInfoEl.innerHTML = '<div class="muted">Loading site...</div>';
      state.siteUrl = SharePoint.trimUrl(siteUrl);
      try {
        const web = await SharePoint.getWeb(role, state.siteUrl);
        state.siteTitle = web.Title || state.siteUrl;
        siteInfoEl.innerHTML = `<div class="notice notice-info">Site: <strong>${esc(state.siteTitle)}</strong> <span class="muted">${esc(state.siteUrl)}</span></div>`;
        await loadLibraries();
      } catch (e) {
        if (e.status === 401 || e.status === 403) { showAccessDenied(); }
        else { siteInfoEl.innerHTML = `<div class="notice notice-warn">Could not open site: ${esc(e.message)}</div>`; }
      }
      fire();
    }

    async function loadLibraries() {
      const libs = await SharePoint.getDocumentLibraries(role, state.siteUrl);
      state.libraries = libs;
      if (!libs.length) {
        libWrap.classList.add('hidden');
        siteInfoEl.insertAdjacentHTML('beforeend', `<div class="notice notice-warn">No document libraries found on this site. If this tenant hasn't approved SharePoint access for the app, grant it and retry.
          <div class="row"><button class="pk-grantspo primary-btn" type="button">Grant SharePoint access</button></div></div>`);
        const gb = siteInfoEl.querySelector('.pk-grantspo');
        if (gb) gb.addEventListener('click', async () => {
          gb.disabled = true; gb.textContent = 'Requesting consent...';
          try {
            await Auth.grantSharePointConsent(role);
            toast('SharePoint access granted. Reloading site...');
            await loadSite(state.siteUrl);
          } catch (e) {
            gb.disabled = false; gb.textContent = 'Grant SharePoint access';
            toast('Consent failed or was cancelled: ' + e.message);
          }
        });
        return;
      }
      libSel.innerHTML = libs.map((l, i) => `<option value="${i}">${esc(l.title)} (${l.itemCount} items)</option>`).join('');
      libWrap.classList.remove('hidden');
      selectLibrary(0);
      libSel.onchange = () => selectLibrary(Number(libSel.value));
    }

    function selectLibrary(i) {
      state.library = state.libraries[i];
      state.root = state.library.serverRelativeUrl;
      state.folder = state.library.serverRelativeUrl;
      openFolder(state.folder);
    }

    async function openFolder(serverRel) {
      folderWrap.classList.remove('hidden');
      foldersEl.innerHTML = '<div class="muted">Loading folders...</div>';
      state.folder = serverRel;
      renderCrumb();
      fire();
      try {
        const children = await SharePoint.getFolderChildren(role, state.siteUrl, serverRel);
        if (!children.folders.length) { foldersEl.innerHTML = '<div class="muted">No subfolders. This folder will be used.</div>'; return; }
        foldersEl.innerHTML = '';
        children.folders.forEach(f => {
          const b = document.createElement('button');
          b.type = 'button'; b.className = 'pk-folder';
          b.innerHTML = `📁 ${esc(f.name)} <span class="muted">(${f.itemCount})</span>`;
          b.addEventListener('click', () => openFolder(f.serverRelativeUrl));
          foldersEl.appendChild(b);
        });
      } catch (e) { foldersEl.innerHTML = `<div class="notice notice-warn">Could not list folders: ${esc(e.message)}</div>`; }
    }

    function renderCrumb() {
      const crumbs = [{ name: state.library.title, path: state.root }];
      const rel = state.folder.startsWith(state.root) ? state.folder.substring(state.root.length) : '';
      let acc = state.root;
      rel.split('/').filter(Boolean).forEach(seg => { acc += `/${seg}`; crumbs.push({ name: decodeURIComponent(seg), path: acc }); });
      crumbEl.innerHTML = '';
      crumbs.forEach((c, i) => {
        if (i) crumbEl.appendChild(document.createTextNode(' › '));
        const a = document.createElement('button');
        a.type = 'button';
        a.className = 'crumb' + (i === crumbs.length - 1 ? ' current' : '');
        a.textContent = c.name;
        a.addEventListener('click', () => openFolder(c.path));
        crumbEl.appendChild(a);
      });
    }

    function showAccessDenied() {
      const conn = Auth.getConnection(role);
      siteInfoEl.classList.add('hidden');
      accessEl.classList.remove('hidden');
      accessEl.className = 'pk-access notice notice-warn';
      accessEl.innerHTML = `Your signed-in account (<strong>${esc(conn.upn)}</strong>) is not a member of this site, so SharePoint denied access to its contents. Being a global or SharePoint administrator does not by itself grant content access - that is why sites you are a member of (or opened by direct URL) work, but others found in search do not. If you hold the <strong>SharePoint Administrator</strong> role you can temporarily make yourself a site collection admin.
        <div class="row"><button class="pk-takeown primary-btn" type="button">Take ownership &amp; retry</button></div>
        <div class="pk-takeown-err muted" style="margin-top:6px"></div>`;
      const errEl = accessEl.querySelector('.pk-takeown-err');
      accessEl.querySelector('.pk-takeown').addEventListener('click', async (ev) => {
        const btn = ev.target; btn.disabled = true; btn.textContent = 'Granting...'; errEl.textContent = '';
        try {
          const claim = SharePoint.buildClaimLoginName(conn.upn);
          await SharePoint.setSiteAdmin(role, state.siteUrl, claim, true);
          toast('Access granted. Reloading site...');
          await loadSite(state.siteUrl);
        } catch (e) {
          btn.disabled = false; btn.textContent = 'Take ownership & retry';
          errEl.textContent = (e.status === 401 || e.status === 403)
            ? `Could not elevate: the tenant admin API denied it. This account needs the SharePoint Administrator role, and the app must be granted SharePoint access in this tenant. (${e.message})`
            : `Take ownership failed: ${e.message}`;
        }
      });
    }

    function subLabel() {
      if (!state.folder || !state.root || state.folder === state.root) return '';
      return decodeURIComponent(state.folder.substring(state.root.length).replace(/^\/+/, ''));
    }

    function getSelection() {
      if (!state.siteUrl || !state.library || !state.folder) return null;
      return {
        siteUrl: state.siteUrl, siteTitle: state.siteTitle,
        library: state.library.title, folder: state.folder, subLabel: subLabel(),
      };
    }

    function fire() { onChange && onChange(); }

    return { getSelection };
  }

  // ---------- jobs dashboard ----------
  const cardRefs = new Map();

  // While any job runs, surface the shared SharePoint throttle state on its card with a
  // live countdown to the next retry; nothing is shown when we're not being throttled.
  let throttleTimer = null;
  function ensureThrottleTicker() {
    if (!throttleTimer) throttleTimer = setInterval(updateThrottleIndicators, 500);
  }
  function updateThrottleIndicators() {
    const st = window.SharePoint?.getThrottleState?.();
    let anyRunning = false;
    for (const [id, el] of cardRefs) {
      const t = el.querySelector('.job-throttle');
      if (!t) continue;
      if (!Transfer.isRunning(id)) { t.classList.add('hidden'); continue; }
      anyRunning = true;
      if (st && st.throttled) {
        t.classList.remove('hidden');
        t.textContent = st.phase === 'recovering'
          ? '⏳ Throttled by SharePoint — recovering…'
          : `⏳ Throttled by SharePoint — retrying in ${Math.ceil(st.waitMs / 1000)}s`;
      } else {
        t.classList.add('hidden');
      }
    }
    if (!anyRunning) { clearInterval(throttleTimer); throttleTimer = null; }
  }

  function renderJobs() {
    const list = Jobs.list();
    const wrap = $('jobsList');
    wrap.innerHTML = '';
    cardRefs.clear();
    $('jobsEmpty').classList.toggle('hidden', list.length > 0);
    let active = 0;
    for (const job of list) {
      if (Transfer.isRunning(job.id) || job.status === 'running' || job.status === 'enumerating') active++;
      wrap.appendChild(buildJobCard(job));
    }
    const badge = $('jobsBadge');
    if (active > 0) { badge.classList.remove('hidden'); badge.textContent = String(active); }
    else badge.classList.add('hidden');
  }

  function statusMeta(job) {
    const running = Transfer.isRunning(job.id);
    if (running && job.status === 'enumerating') return { cls: 'review', label: 'Scanning' };
    if (running) return { cls: 'review', label: 'Running' };
    switch (job.status) {
      case 'completed': return { cls: 'keep', label: 'Completed' };
      case 'completed_with_errors': return { cls: 'review', label: 'Completed with errors' };
      case 'cancelled': return { cls: 'noaccess', label: 'Cancelled' };
      case 'failed': return { cls: 'cleanup', label: 'Failed' };
      case 'running': case 'enumerating': return { cls: 'unknown', label: 'Interrupted' };
      default: return { cls: 'unknown', label: 'Ready' };
    }
  }

  function buildJobCard(job) {
    const el = document.createElement('div');
    el.className = 'job-card';
    el.dataset.id = job.id;
    const m = statusMeta(job);
    const s = job.stats || {};
    const pct = s.bytesTotal ? Math.round((s.bytesDone / s.bytesTotal) * 100) : (s.total ? Math.round((s.done / s.total) * 100) : 0);
    el.innerHTML = `
      <div class="job-head">
        <span class="badge ${m.cls} job-status">${esc(m.label)}</span>
        <span class="job-title">${esc(job.config.source.siteTitle)} <span class="muted">→</span> ${esc(job.config.target.siteTitle)}</span>
        <span class="job-mode">${job.config.mode === 'cross' ? 'cross-tenant' : 'same-tenant'} · ${esc(job.config.operation)}</span>
      </div>
      <div class="job-sub muted">${esc(job.config.source.folder)} <span>→</span> ${esc(job.config.target.folder)}</div>
      <div class="progress-box">
        <div class="progress-header"><span class="job-current"></span><span class="job-eta"></span></div>
        <div class="progress-bar"><div class="job-fill" style="width:${pct}%"></div></div>
        <div class="progress-footer"><span class="job-stats"></span><span class="job-speed"></span></div>
        <div class="job-throttle hidden"></div>
      </div>
      <div class="job-actions"></div>`;
    cardRefs.set(job.id, el);
    updateJobCard(job, null, el);
    return el;
  }

  function updateJobCard(job, snap, el) {
    el = el || cardRefs.get(job.id);
    if (!el) return;
    const s = snap || job.stats || {};
    const total = s.itemsTotal ?? job.stats.total;
    const done = s.itemsDone ?? job.stats.done;
    const failed = s.failed ?? job.stats.failed;
    const bytesTotal = s.bytesTotal ?? job.stats.bytesTotal;
    const bytesDone = s.bytesDone ?? job.stats.bytesDone;
    const pct = bytesTotal ? Math.round((bytesDone / bytesTotal) * 100) : (total ? Math.round((done / total) * 100) : 0);
    el.querySelector('.job-fill').style.width = `${pct}%`;
    const m = statusMeta(job);
    const st = el.querySelector('.job-status'); st.textContent = m.label; st.className = `badge ${m.cls} job-status`;
    const versionsTotal = job.stats?.versionsTotal || 0;
    el.querySelector('.job-stats').textContent = `${done}/${total} files${versionsTotal ? ` +${versionsTotal} versions` : ''}${failed ? ` · ${failed} failed` : ''}${bytesTotal ? ` · ${fmtBytes(bytesDone)}/${fmtBytes(bytesTotal)}` : ''}`;
    const speed = snap?.speedBps ? `${fmtBytes(snap.speedBps)}/s` : '';
    el.querySelector('.job-speed').textContent = speed;
    el.querySelector('.job-current').textContent = snap?.currentFile ? `Current: ${snap.currentFile}` : '';
    const eta = (snap?.etaSec != null) ? `ETA ${fmtDuration(snap.etaSec)}` : '';
    el.querySelector('.job-eta').textContent = eta;
    renderJobActions(job, el);
  }

  function renderJobActions(job, el) {
    const box = el.querySelector('.job-actions');
    const running = Transfer.isRunning(job.id);
    const s = job.stats || {};
    const buttons = [];
    if (running) {
      buttons.push(`<button class="ghost-btn" data-act="pause">Pause</button>`);
      buttons.push(`<button class="ghost-btn" data-act="cancel">Cancel</button>`);
    } else {
      const hasPending = (s.total || 0) - (s.done || 0) - (s.failed || 0) - (s.skipped || 0) > 0;
      if (job.status === 'draft') {
        buttons.push(`<button class="primary-btn" data-act="resume">Start</button>`);
      } else if (hasPending) {
        buttons.push(`<button class="primary-btn" data-act="resume">Resume</button>`);
      }
      if ((s.failed || 0) > 0) {
        buttons.push(`<button class="primary-btn" data-act="retry">Retry failed (${s.failed})</button>`);
      }
      const hasRun = ((s.done || 0) + (s.skipped || 0) + (s.deleted || 0)) > 0;
      if (job.status !== 'draft' && !hasPending && hasRun) {
        buttons.push(`<button class="ghost-btn" data-act="rerun">Re-run</button>`);
      }
      buttons.push(`<button class="ghost-btn" data-act="delete">Delete</button>`);
    }
    if (job.status !== 'draft' && (s.total || (job.items && job.items.length))) {
      buttons.push(`<button class="ghost-btn" data-act="report">Report</button>`);
    }
    box.innerHTML = buttons.join('');
    box.querySelectorAll('button').forEach(b => b.addEventListener('click', () => jobAction(job.id, b.dataset.act, b)));
  }

  let pausedJobs = new Set();
  function jobAction(id, act, btn) {
    if (act === 'resume') { startJobRun(id, { action: 'resume' }); }
    else if (act === 'retry') { startJobRun(id, { action: 'retry', autoReport: false }); }
    else if (act === 'pause') {
      if (pausedJobs.has(id)) { Transfer.resumeRun(id); pausedJobs.delete(id); btn.textContent = 'Pause'; }
      else { Transfer.pause(id); pausedJobs.add(id); btn.textContent = 'Paused - resume'; }
    }
    else if (act === 'cancel') {
      Transfer.cancel(id); pausedJobs.delete(id);
      const el = cardRefs.get(id);
      if (el) {
        const st = el.querySelector('.job-status');
        if (st) { st.textContent = 'Cancelling'; st.className = 'badge review job-status'; }
        el.querySelectorAll('.job-actions button').forEach(b => b.disabled = true);
      }
    }
    else if (act === 'report') { openJobReport(id); }
    else if (act === 'rerun') { showRerunChoice(id); }
    else if (act === 'delete') {
      const job = Jobs.get(id);
      if (Transfer.isRunning(id)) { toast('Cancel the job before deleting it.'); return; }
      if (job && ['completed', 'completed_with_errors'].includes(job.status) === false && (job.stats?.done || 0) > 0) {
        if (!confirm('Delete this job and its resume state? Files already transferred stay in the target.')) return;
      }
      Jobs.remove(id); renderJobs();
    }
  }

  function showRerunChoice(id) {
    const el = cardRefs.get(id); if (!el) return;
    const box = el.querySelector('.job-actions');
    box.innerHTML = `<span class="muted rerun-q">Re-run:</span>
      <button class="primary-btn" data-r="incremental" title="Transfer new and changed files; for copy jobs also remove target files that were deleted from the source.">Sync changes</button>
      <button class="ghost-btn" data-r="full" title="Re-transfer every file currently in the source.">Full re-run</button>
      <button class="ghost-btn" data-r="cancel">Cancel</button>`;
    box.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
      const mode = b.dataset.r;
      if (mode === 'cancel') { const job = Jobs.get(id); if (job) renderJobActions(job, el); return; }
      startRerun(id, mode);
    }));
  }

  function startRerun(id, mode) {
    const job = Jobs.get(id);
    if (!job) return;
    job.rerun = mode;
    job.status = 'enumerating';
    job.startedAt = null;
    job.finishedAt = null;
    Jobs.save(job);
    Log.info(`Re-run (${mode}) requested for job ${String(id).slice(0, 8)}: re-scanning source for ${mode === 'incremental' ? 'new / changed / deleted objects' : 'a full transfer'}.`);
    startJobRun(id, { action: 'rerun' });
  }

  async function ensureJobConnections(job) {
    await ensureRoleConnection('source', job.config.source);
    if (job.config.mode === 'cross') await ensureRoleConnection('target', job.config.target);
  }

  // Reconnect one role for an existing job: reuse a matching live connection, else silently
  // restore it from the MSAL cache, else prompt an interactive (tenant-pinned) sign-in via a
  // modal that spells out WHICH tenant to pick. Then ensure SharePoint consent (silent if it
  // already was).
  async function ensureRoleConnection(role, loc) {
    if (!Auth.connectionMatches(role, loc)) {
      const restored = await Auth.restoreConnection(role, loc);
      if (!restored) {
        await promptReconnect(role, loc);
      }
    }
    const conn = Auth.getConnection(role);
    if (conn && !conn.spoConsent) {
      const ok = await Auth.hasSharePointConsent(role);
      if (ok) conn.spoConsent = true;
      else {
        toast(`Approve SharePoint access for the ${role} tenant to continue...`);
        await Auth.grantSharePointConsent(role);
      }
    }
    updateUserBox();
    updateWizard();
  }

  // Show an explicit "which tenant" dialog, then open the sign-in popup from the button click
  // (keeps the user gesture so the popup is not blocked). Resolves once signed in; rejects on cancel.
  function promptReconnect(role, loc) {
    return new Promise((resolve, reject) => {
      const isSrc = role === 'source';
      const tenant = loc.tenantName || loc.tenantHost || 'the required tenant';
      $('reconnectRole').innerHTML = `<span class="badge ${isSrc ? 'src' : 'tgt'}">${isSrc ? 'SOURCE' : 'TARGET'} tenant</span>`;
      $('reconnectMsg').innerHTML =
        `Sign in to the <strong>${isSrc ? 'source' : 'target'}</strong> tenant for this job:<br>` +
        `<strong>${esc(tenant)}</strong>` +
        (loc.tenantHost ? ` <span class="muted">(${esc(loc.tenantHost)})</span>` : '') +
        (loc.upn ? `<br>Use the account <strong>${esc(loc.upn)}</strong>.` : '') +
        `<br><span class="muted">Choose an account in this tenant in the Microsoft sign-in window that opens next.</span>`;
      const errEl = $('reconnectErr'); errEl.classList.add('hidden'); errEl.textContent = '';
      const signInBtn = $('reconnectSignIn');
      const cancelBtn = $('reconnectCancel');
      signInBtn.textContent = `Sign in to ${isSrc ? 'source' : 'target'} tenant`;
      signInBtn.disabled = false;

      const cleanup = () => {
        signInBtn.onclick = null; cancelBtn.onclick = null;
        $('reconnectModal').removeEventListener('click', onBackdrop);
        document.removeEventListener('keydown', onKey);
        closeReconnect();
      };
      const onBackdrop = (e) => { if (e.target.dataset && e.target.dataset.close) { cleanup(); reject(new Error('Reconnect cancelled')); } };
      const onKey = (e) => { if (e.key === 'Escape') { cleanup(); reject(new Error('Reconnect cancelled')); } };

      signInBtn.onclick = async () => {
        signInBtn.disabled = true;
        try {
          await Auth.signIn(role, { loginHint: loc.upn, expectTenantId: loc.tenantId, expectTenantHost: loc.tenantHost });
          cleanup();
          resolve(true);
        } catch (e) {
          signInBtn.disabled = false;
          errEl.textContent = e.message;
          errEl.classList.remove('hidden');
        }
      };
      cancelBtn.onclick = () => { cleanup(); reject(new Error('Reconnect cancelled')); };
      $('reconnectModal').addEventListener('click', onBackdrop);
      document.addEventListener('keydown', onKey);

      $('reconnectModal').classList.remove('hidden');
      document.body.classList.add('modal-open');
    });
  }

  function closeReconnect() {
    $('reconnectModal').classList.add('hidden');
    document.body.classList.remove('modal-open');
  }

  async function startJobRun(id, opts = {}) {
    const pre = Jobs.get(id);
    if (!pre) return;
    // Make sure both tenants are connected before running so a resume/re-run after a reload
    // reconnects (silently from cache, or via popup) instead of failing with "not connected".
    const el0 = cardRefs.get(id);
    try {
      if (el0) el0.querySelector('.job-current').textContent = 'Checking tenant connections...';
      await ensureJobConnections(pre);
    } catch (e) {
      toast(`Cannot start: ${e.message}`);
      Log.err('Reconnect failed:', e);
      if (el0) el0.querySelector('.job-current').textContent = '';
      renderJobs();
      return;
    }
    if (pre && pre.status !== 'draft' && (opts.action === 'resume' || opts.action === 'retry')) {
      const s = pre.stats || {};
      const failed = s.failed || 0;
      const pending = Math.max(0, (s.total || 0) - (s.done || 0) - failed - (s.skipped || 0));
      const parts = [];
      if (failed) parts.push(`${failed} previously failed`);
      if (pending) parts.push(`${pending} not yet transferred`);
      const scope = parts.length ? parts.join(' + ') : 'no remaining items';
      Log.info(`${opts.action === 'retry' ? 'Retry' : 'Resume'} job ${String(id).slice(0, 8)}: processing ${scope}.`);
      if (opts.action === 'retry' && pending) {
        toast(`Retrying ${failed} failed and also ${pending} not-yet-transferred file(s) from the cancelled run.`);
      }
    }
    renderJobs();
    ensureThrottleTicker();
    Transfer.run(id, {
      onEnumerate: (n) => {
        const el = cardRefs.get(id); if (el) el.querySelector('.job-current').textContent = `Scanning source... ${n} files found`;
      },
      onVersions: (done, total) => {
        const el = cardRefs.get(id); if (el) el.querySelector('.job-current').textContent = `Reading version history... ${done}/${total}`;
      },
      onProgress: (snap, job) => updateJobCard(job, snap),
    }).then(() => { renderJobs(); if (opts.autoReport !== false) maybeAutoReport(id); refreshOpenReport(id); })
      .catch(e => { toast(`Job error: ${e.message}`); renderJobs(); refreshOpenReport(id); });
    // reflect running state on buttons immediately
    const job = Jobs.get(id); const el = cardRefs.get(id); if (job && el) renderJobActions(job, el);
  }

  // If the report modal is open for this job, reload its items and re-render in place
  // so a retry updates statuses/attempt counts without replacing the user's view.
  async function refreshOpenReport(id) {
    if (!reportState || reportState.jobId !== id) return;
    if ($('reportModal').classList.contains('hidden')) return;
    try { reportState.items = await Jobs.loadItems(id); reportState.job = Jobs.get(id) || reportState.job; } catch (e) { Log.warn('Report refresh failed:', e.message); }
    if (reportState && reportState.jobId === id && !$('reportModal').classList.contains('hidden')) renderReport();
  }

  // Auto-open the report when a job reaches a finished state (unless another
  // report is already open or the job was simply cancelled with nothing done).
  function maybeAutoReport(id) {
    const job = Jobs.get(id);
    if (!job) return;
    const done = ['completed', 'completed_with_errors', 'failed'].includes(job.status);
    if (!done) return;
    if (!$('reportModal').classList.contains('hidden')) return;
    if (!(job.stats?.total) && !(job.items && job.items.length)) return;
    openJobReport(id);
  }

  function fmtDuration(sec) {
    if (sec == null) return '';
    if (sec < 60) return `${sec}s`;
    if (sec < 3600) return `${Math.floor(sec / 60)}m ${sec % 60}s`;
    return `${Math.floor(sec / 3600)}h ${Math.floor((sec % 3600) / 60)}m`;
  }

  // ---------- job report ----------
  let reportState = null; // { jobId, filter, term, items, job }

  async function openJobReport(id) {
    const job = Jobs.get(id);
    if (!job) { toast('Job not found.'); return; }
    reportState = { jobId: id, filter: 'all', term: '', items: [], job };
    $('reportModal').classList.remove('hidden');
    document.body.classList.add('modal-open');
    $('reportBody').innerHTML = '<div class="muted" style="padding:20px">Loading report...</div>';
    try { reportState.items = await Jobs.loadItems(id); } catch (e) { Log.warn('Report load failed:', e.message); }
    if (reportState && reportState.jobId === id) renderReport();
  }

  function closeReport() { $('reportModal').classList.add('hidden'); document.body.classList.remove('modal-open'); reportState = null; }

  function donutSvg(done, failed, skipped, pending, deleted = 0) {
    const total = Math.max(1, done + failed + skipped + pending + deleted);
    const R = 52, C = 2 * Math.PI * R;
    const seg = (val, color, offset) => {
      const len = (val / total) * C;
      return `<circle r="${R}" cx="70" cy="70" fill="none" stroke="${color}" stroke-width="18" stroke-dasharray="${len} ${C - len}" stroke-dashoffset="${-offset}" transform="rotate(-90 70 70)"></circle>`;
    };
    let acc = 0;
    const parts = [];
    const add = (val, color) => { if (val > 0) { parts.push(seg(val, color, (acc / total) * C)); acc += val; } };
    add(done, 'var(--ok)');
    add(failed, 'var(--danger)');
    add(deleted, '#a855f7');
    add(skipped, '#3b82f6');
    add(pending, '#c1c9da');
    const pct = Math.round((done / total) * 100);
    return `<svg viewBox="0 0 140 140" width="150" height="150" role="img" aria-label="Result chart">
      <circle r="52" cx="70" cy="70" fill="none" stroke="var(--border)" stroke-width="18"></circle>
      ${parts.join('')}
      <text x="70" y="66" text-anchor="middle" font-size="24" font-weight="700" fill="var(--text)">${pct}%</text>
      <text x="70" y="88" text-anchor="middle" font-size="11" fill="var(--muted)">success</text>
    </svg>`;
  }

  function renderReport() {
    if (!reportState) return;
    const job = Jobs.get(reportState.jobId) || reportState.job;
    if (!job) { closeReport(); return; }
    const items = reportState.items || [];
    const s = job.stats || {};
    const total = s.total || items.length;
    const done = items.length ? items.filter(i => i.status === 'done').length : (s.done || 0);
    const failed = items.length ? items.filter(i => i.status === 'failed').length : (s.failed || 0);
    const skipped = items.length ? items.filter(i => i.status === 'skipped').length : (s.skipped || 0);
    const deleted = items.length ? items.filter(i => i.status === 'deleted').length : (s.deleted || 0);
    const pending = Math.max(0, total - done - failed - skipped - deleted);
    const m = statusMeta(job);
    const durationMs = (job.finishedAt && job.startedAt) ? (job.finishedAt - job.startedAt)
      : (job.startedAt ? Date.now() - job.startedAt : 0);
    const durationSec = durationMs / 1000;
    const skippedBytes = items.filter(i => i.status === 'skipped').reduce((a, b) => a + (b.size || 0) + (b.versionBytes || 0), 0);
    const transferredBytes = Math.max(0, (s.bytesDone || 0) - skippedBytes);
    const avgBps = durationSec > 0 ? transferredBytes / durationSec : 0;
    const retried = items.filter(i => (i.attempts || 0) > 1).length;
    const versionsTotal = items.reduce((a, b) => a + (b.versionCount || 0), 0) || (s.versionsTotal || 0);

    const summary = `
      <div class="report-summary">
        <div class="report-chart">${donutSvg(done, failed, skipped, pending, deleted)}</div>
        <div class="report-legend">
          <div class="rl-row"><span class="rl-dot ok"></span> Succeeded <strong>${done}</strong></div>
          <div class="rl-row"><span class="rl-dot bad"></span> Failed <strong>${failed}</strong></div>
          ${skipped ? `<div class="rl-row"><span class="rl-dot skip"></span> Skipped <strong>${skipped}</strong></div>` : ''}
          ${deleted ? `<div class="rl-row"><span class="rl-dot del"></span> Deleted on target <strong>${deleted}</strong></div>` : ''}
          <div class="rl-row"><span class="rl-dot pend"></span> Not transferred <strong>${pending}</strong></div>
        </div>
        <dl class="report-stats">
          <dt>Status</dt><dd><span class="badge ${m.cls}">${esc(m.label)}</span></dd>
          <dt>Operation</dt><dd>${esc(job.config.operation)} · ${job.config.mode === 'cross' ? 'cross-tenant' : 'same-tenant'}</dd>
          <dt>Files</dt><dd>${done} of ${total}${skipped ? ` · ${skipped} skipped` : ''}${deleted ? ` · ${deleted} deleted` : ''}</dd>
          <dt>Data</dt><dd>${fmtBytes(s.bytesDone || 0)} / ${fmtBytes(s.bytesTotal || 0)}${versionsTotal ? ' (incl. versions)' : ''}</dd>
          ${versionsTotal ? `<dt>Versions</dt><dd>${versionsTotal} version(s) included</dd>` : ''}
          <dt>Avg speed</dt><dd>${avgBps ? fmtBytes(avgBps) + '/s' : '-'}</dd>
          <dt>Duration</dt><dd>${durationMs ? fmtDuration(Math.round(durationSec)) : '-'}</dd>
          ${retried ? `<dt>Retried</dt><dd>${retried} file(s) took more than one try</dd>` : ''}
          <dt>Route</dt><dd>${esc(job.config.source.siteTitle)} &rarr; ${esc(job.config.target.siteTitle)}</dd>
        </dl>
      </div>`;

    const counts = { all: items.length, failed, skipped, done, pending, deleted };
    const filterKeys = ['all', 'failed'].concat(skipped ? ['skipped'] : []).concat(deleted ? ['deleted'] : []).concat(['done', 'pending']);
    const filters = filterKeys.map(f =>
      `<button type="button" class="rf-btn ${reportState.filter === f ? 'active' : ''}" data-filter="${f}">${f[0].toUpperCase() + f.slice(1)} (${counts[f] ?? 0})</button>`).join('');

    const controls = `
      <div class="report-controls">
        <div class="rf-group">${filters}</div>
        <input type="search" class="report-search" placeholder="Filter by file path..." value="${esc(reportState.term)}" />
        <button type="button" class="ghost-btn" id="reportExport">Export CSV</button>
      </div>`;

    $('reportBody').innerHTML = summary + controls + `<div id="reportTableWrap"></div>`;

    $('reportBody').querySelectorAll('.rf-btn').forEach(b => b.addEventListener('click', () => {
      reportState.filter = b.dataset.filter;
      $('reportBody').querySelectorAll('.rf-btn').forEach(x => x.classList.toggle('active', x.dataset.filter === reportState.filter));
      renderReportRows();
    }));
    const searchInput = $('reportBody').querySelector('.report-search');
    searchInput.addEventListener('input', () => { reportState.term = searchInput.value; renderReportRows(); });
    const exportBtn = $('reportBody').querySelector('#reportExport');
    if (exportBtn) exportBtn.addEventListener('click', () => exportReportCsv(job));
    renderReportRows();
  }

  function currentReportItems() {
    const items = (reportState && reportState.items) || [];
    const t = reportState.term.trim().toLowerCase();
    return items.filter(it => {
      const st = it.status || 'pending';
      if (reportState.filter === 'failed' && st !== 'failed') return false;
      if (reportState.filter === 'done' && st !== 'done') return false;
      if (reportState.filter === 'skipped' && st !== 'skipped') return false;
      if (reportState.filter === 'deleted' && st !== 'deleted') return false;
      if (reportState.filter === 'pending' && (st === 'done' || st === 'failed' || st === 'skipped' || st === 'deleted')) return false;
      if (t && !(it.subPath || '').toLowerCase().includes(t)) return false;
      return true;
    });
  }

  function renderReportRows() {
    if (!reportState) return;
    const wrap = document.getElementById('reportTableWrap');
    if (!wrap) return;
    const filtered = currentReportItems();
    if (!filtered.length) { wrap.innerHTML = '<div class="empty-state">No items match.</div>'; return; }
    const capped = filtered.slice(0, 500);
    const rows = capped.map(it => {
      const st = it.status || 'pending';
      const badge = st === 'done' ? 'keep' : st === 'failed' ? 'cleanup' : st === 'skipped' ? 'noaccess' : st === 'deleted' ? 'review' : 'unknown';
      const label = st === 'done' ? 'OK' : st === 'failed' ? 'Failed' : st === 'skipped' ? 'Skipped' : st === 'deleted' ? 'Deleted' : 'Pending';
      const tries = it.attempts || 0;
      const detail = st === 'skipped' ? (it.skipReason || 'Skipped') : st === 'deleted' ? 'Removed from target (deleted from source)' : (it.error || '');
      return `<tr>
        <td><span class="badge ${badge}">${label}</span></td>
        <td class="rt-path" title="${esc(it.subPath || '')}">${esc(it.subPath || '')}</td>
        <td class="rt-size">${fmtBytes(it.size || 0)}</td>
        <td class="rt-tries" title="${tries} attempt(s)">${tries || ''}</td>
        <td class="rt-err" title="${esc(detail)}">${esc(detail)}</td>
      </tr>`;
    }).join('');
    wrap.innerHTML = `
      ${filtered.length > capped.length ? `<div class="muted rt-cap">Showing first ${capped.length} of ${filtered.length}. Narrow with search or export CSV for the full list.</div>` : ''}
      <div class="report-table-scroll">
        <table class="report-table">
          <thead><tr><th>Status</th><th>File</th><th>Size</th><th>Tries</th><th>Details</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  }

  function exportReportCsv(job) {
    const items = (reportState && reportState.items) || [];
    const cell = (v) => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;
    const rows = [['status', 'file', 'size_bytes', 'versions', 'version_bytes', 'attempts', 'details']];
    items.forEach(it => rows.push([it.status || 'pending', it.subPath || '', it.size || 0, it.versionCount || 0, it.versionBytes || 0, it.attempts || 0, it.status === 'skipped' ? (it.skipReason || 'Skipped') : (it.error || '')]));
    const csv = rows.map(r => r.map(cell).join(',')).join('\r\n');
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `SPOMove-report-${String(job.id).slice(0, 8)}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  // ---------- boot ----------
  async function boot() {
    initChrome();
    initWizard();
    renderJobs();
    try {
      await Auth.init();
      updateUserBox();
      updateWizard();
      Log.info('SPOMove ready.');
    } catch (e) {
      const be = $('bootError');
      be.classList.remove('hidden');
      be.textContent = `Startup problem: ${e.message}`;
      Log.err('Boot failed:', e);
    }
  }

  window.addEventListener('DOMContentLoaded', boot);
})();
