// SPOMove UI: wizard, location pickers, jobs dashboard with live progress, reports and
// the record of site admin rights the tool added.
(() => {
  const { esc, toast, busy, h, icon, iconHtml } = Toolkit;
  const $ = (id) => document.getElementById(id);
  const fmtBytes = (n) => Transfer.fmtBytes(n);
  const plural = (n, word) => `${n} ${word}${n === 1 ? '' : 's'}`;

  const STEP_LABELS = {
    mode: 'Type', sourceConnect: 'Source login', sourceLoc: 'Source',
    targetConnect: 'Target login', targetLoc: 'Target', options: 'Options', review: 'Review',
  };
  const DEFAULT_OPTIONS = { operation: 'copy', includeVersions: false, includeMetadata: true, overwrite: false, onlyNewer: false, concurrency: 4 };

  const wiz = { step: 0, mode: null, sourcePicker: null, targetPicker: null, options: { ...DEFAULT_OPTIONS } };
  let activateTab = () => {};

  function steps() {
    return wiz.mode === 'cross'
      ? ['mode', 'sourceConnect', 'sourceLoc', 'targetConnect', 'targetLoc', 'options', 'review']
      : ['mode', 'sourceConnect', 'sourceLoc', 'targetLoc', 'options', 'review'];
  }
  const targetRole = () => (wiz.mode === 'cross' ? 'target' : 'source');
  const sameSite = (a, b) => SharePoint.trimUrl(a).toLowerCase() === SharePoint.trimUrl(b).toLowerCase();

  // A running job reads its tokens from these connections; switching one under it would
  // send its next calls to another account or tenant.
  function assertRoleFree(role) {
    if (Transfer.rolesInUse().has(role)) throw new Error(`A running job uses the ${role} connection. Let it finish or cancel it first.`);
  }

  function updateUserBox() { Toolkit.setUser(Auth.userInfo(onSignOut)); }

  // ---------- wizard ----------
  function initWizard() {
    document.querySelectorAll('.choice-card').forEach((c) => {
      c.addEventListener('click', () => {
        wiz.mode = c.dataset.mode;
        document.querySelectorAll('.choice-card').forEach((x) => {
          x.classList.toggle('selected', x === c);
          x.setAttribute('aria-pressed', x === c ? 'true' : 'false');
        });
        updateWizard();
      });
    });
    $('sourceSignInBtn').addEventListener('click', (e) => doSignIn('source', e.currentTarget));
    $('targetSignInBtn').addEventListener('click', (e) => doSignIn('target', e.currentTarget));
    $('wizBack').addEventListener('click', () => { if (wiz.step > 0) { wiz.step--; updateWizard(); } });
    $('wizNext').addEventListener('click', () => { if (canAdvance()) { wiz.step++; onEnterStep(); updateWizard(); } });
    $('wizStart').addEventListener('click', startTransfer);

    document.querySelectorAll('input[name="operation"]').forEach((r) =>
      r.addEventListener('change', (e) => { wiz.options.operation = e.target.value; refreshOptionWarning(); }));
    $('optMetadata').addEventListener('change', (e) => { wiz.options.includeMetadata = e.target.checked; });
    $('optVersions').addEventListener('change', (e) => { wiz.options.includeVersions = e.target.checked; refreshOptionWarning(); });
    $('optOverwrite').addEventListener('change', (e) => {
      wiz.options.overwrite = e.target.checked;
      const on = $('optOnlyNewer');
      on.disabled = !e.target.checked;
      if (!e.target.checked) { on.checked = false; wiz.options.onlyNewer = false; }
      refreshOptionWarning();
    });
    $('optOnlyNewer').addEventListener('change', (e) => { wiz.options.onlyNewer = e.target.checked; });
    $('optConcurrency').addEventListener('change', (e) => {
      wiz.options.concurrency = Math.max(1, Math.min(12, Number(e.target.value) || 4));
      e.target.value = wiz.options.concurrency;
    });
    resetWizard();
  }

  function resetWizard() {
    wiz.step = 0;
    wiz.mode = null;
    wiz.sourcePicker = null;
    wiz.targetPicker = null;
    document.querySelectorAll('.choice-card').forEach((x) => { x.classList.remove('selected'); x.setAttribute('aria-pressed', 'false'); });
    $('sourcePicker').replaceChildren();
    $('targetPicker').replaceChildren();
    wiz.options = { ...DEFAULT_OPTIONS };
    document.querySelector('input[name="operation"][value="copy"]').checked = true;
    $('optMetadata').checked = true;
    $('optVersions').checked = false;
    $('optOverwrite').checked = false;
    $('optConcurrency').value = 4;
    $('optOnlyNewer').checked = false;
    $('optOnlyNewer').disabled = true;
    updateWizard();
  }

  function onEnterStep() {
    const id = steps()[wiz.step];
    if (id === 'sourceLoc' && !wiz.sourcePicker) wiz.sourcePicker = createPicker($('sourcePicker'), 'source', updateNavState);
    if (id === 'targetLoc' && !wiz.targetPicker) wiz.targetPicker = createPicker($('targetPicker'), targetRole(), updateNavState);
    if (id === 'review') renderReview();
    if (id === 'options') refreshOptionWarning();
  }

  function renderStepper() {
    $('stepper').innerHTML = steps().map((s, i) => {
      const cls = `step${i === wiz.step ? ' active' : ''}${i < wiz.step ? ' done' : ''}`;
      const dot = i < wiz.step ? iconHtml('i-check', 'sm') : String(i + 1);
      return `<li class="${cls}"><span class="step-dot">${dot}</span><span class="step-label">${esc(STEP_LABELS[s])}</span></li>`;
    }).join('');
  }

  function updateWizard() {
    renderStepper();
    const cur = steps()[wiz.step];
    document.querySelectorAll('.wiz-step').forEach((el) => { el.hidden = el.dataset.step !== cur; });
    renderConnState('source', $('sourceConnState'), $('sourceSignInBtn'));
    renderConnState('target', $('targetConnState'), $('targetSignInBtn'));
    updateNavState();
  }

  function renderConnState(role, stateEl, btn) {
    const c = Auth.getConnection(role);
    const label = (text) => btn.replaceChildren(icon('i-login'), text);
    if (!c || !c.account) {
      stateEl.replaceChildren();
      label(role === 'source' ? 'Sign in to the source tenant' : 'Sign in to the target tenant');
      return;
    }
    let html = `<div class="notice notice-ok">${iconHtml('i-check-circle')}<div>Connected to <strong>${esc(c.tenantName || c.tenantHost || 'tenant')}</strong> as <strong>${esc(c.upn)}</strong>.</div></div>`;
    if (!c.spoConsent) {
      html += `<div class="notice notice-warn">${iconHtml('i-alert')}<div>This tenant has not approved SharePoint access for the app yet. That is needed to browse and move files; one prompt covers all SharePoint hosts of the tenant.
        <div class="row"><button class="btn btn-primary btn-small conn-grantspo" type="button">${iconHtml('i-key')}Grant SharePoint access</button></div></div></div>`;
    }
    stateEl.innerHTML = html;
    label('Switch account');
    const gb = stateEl.querySelector('.conn-grantspo');
    if (gb) gb.addEventListener('click', () => grantSpo(role, gb));
  }

  async function grantSpo(role, btn) {
    await busy(btn, async () => {
      try {
        await Auth.grantSharePointConsent(role);
        toast('SharePoint access approved.', 'ok');
        updateWizard();
        advanceFromConnect();
      } catch (e) {
        toast(`Consent failed or was cancelled: ${e.message}`, 'err');
      }
    });
  }

  function advanceFromConnect() {
    const cur = steps()[wiz.step];
    if ((cur === 'sourceConnect' || cur === 'targetConnect') && canAdvance()) { wiz.step++; onEnterStep(); updateWizard(); }
  }

  function currentOverlap() {
    const src = wiz.sourcePicker && wiz.sourcePicker.getSelection();
    const tgt = wiz.targetPicker && wiz.targetPicker.getSelection();
    if (!src || !tgt) return null;
    const sc = Auth.getConnection('source');
    const tc = Auth.getConnection(targetRole());
    return MoveRules.overlapError({ ...src, tenantId: sc && sc.tenantId }, { ...tgt, tenantId: tc && tc.tenantId });
  }

  function updateNavState() {
    $('wizBack').disabled = wiz.step === 0;
    const cur = steps()[wiz.step];
    const isReview = cur === 'review';
    $('wizNext').hidden = isReview;
    $('wizStart').hidden = !isReview;
    $('wizNext').disabled = !canAdvance();
    const overlap = cur === 'targetLoc' ? currentOverlap() : null;
    $('overlapNotice').hidden = !overlap;
    $('overlapText').textContent = overlap || '';
    if (isReview) renderReview();
    updateUserBox();
  }

  function canAdvance() {
    switch (steps()[wiz.step]) {
      case 'mode': return !!wiz.mode;
      case 'sourceConnect': return Auth.isConnected('source') && !!Auth.getConnection('source').spoConsent;
      case 'sourceLoc': return !!(wiz.sourcePicker && wiz.sourcePicker.getSelection());
      case 'targetConnect': return Auth.isConnected('target') && !!Auth.getConnection('target').spoConsent;
      case 'targetLoc': return !!(wiz.targetPicker && wiz.targetPicker.getSelection()) && !currentOverlap();
      case 'options': return true;
      default: return false;
    }
  }

  async function doSignIn(role, btn) {
    try { assertRoleFree(role); } catch (e) { toast(e.message, 'warn'); return; }
    await busy(btn, async () => {
      try {
        await Auth.signIn(role);
        // A new account invalidates locations picked with the old one.
        if (role === 'source') { wiz.sourcePicker = null; $('sourcePicker').replaceChildren(); if (wiz.mode !== 'cross') { wiz.targetPicker = null; $('targetPicker').replaceChildren(); } }
        if (role === 'target') { wiz.targetPicker = null; $('targetPicker').replaceChildren(); }
        updateWizard();
        advanceFromConnect();
      } catch (e) {
        toast(e.message, 'err');
        Log.err('Sign-in failed:', e);
        updateWizard();
      }
    });
  }

  function refreshOptionWarning() {
    const msgs = [];
    if (wiz.mode === 'cross') msgs.push('Between tenants every file goes through your browser. Large libraries are limited by your bandwidth and memory; keep this tab open until the job is done.');
    if (wiz.options.includeVersions) msgs.push('Previous versions multiply the data that is transferred and make the job a lot slower.');
    if (wiz.options.operation === 'move') msgs.push('Move sends each source file to the source recycle bin after it transferred. Try a small folder first.');
    $('optWarning').hidden = !msgs.length;
    $('optWarningText').innerHTML = msgs.map((m) => `<p>${esc(m)}</p>`).join('');
  }

  function locLabel(sel) {
    if (!sel) return '';
    return [sel.siteTitle, sel.library, sel.subLabel].filter(Boolean).join(' / ');
  }

  function renderReview() {
    const src = wiz.sourcePicker && wiz.sourcePicker.getSelection();
    const tgt = wiz.targetPicker && wiz.targetPicker.getSelection();
    const sc = Auth.getConnection('source');
    const tc = Auth.getConnection(targetRole());
    const o = wiz.options;
    const rows = [
      ['Type', wiz.mode === 'cross' ? 'Between two tenants' : 'Same tenant'],
      ['Operation', o.operation === 'move' ? 'Move (originals go to the source recycle bin)' : 'Copy'],
      ['Source tenant', `${sc.tenantName || sc.tenantHost || ''} (${sc.upn || ''})`],
      ['Source', locLabel(src)],
      ['Target tenant', `${tc.tenantName || tc.tenantHost || ''} (${tc.upn || ''})`],
      ['Target', locLabel(tgt)],
      ['Metadata', o.includeMetadata ? 'Keep the original' : 'Reset to you and now'],
      ['Previous versions', o.includeVersions ? 'Included' : 'Current version only'],
      ['Overwrite existing', o.overwrite ? (o.onlyNewer ? 'Yes, only when the source is newer' : 'Yes') : 'No, existing files are skipped'],
      ['Parallel files', String(o.concurrency)],
    ];
    $('reviewList').innerHTML = rows.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join('');
  }

  async function startTransfer() {
    const src = wiz.sourcePicker.getSelection();
    const tgt = wiz.targetPicker.getSelection();
    const sc = Auth.getConnection('source');
    const tc = Auth.getConnection(targetRole());
    const overlap = currentOverlap();
    if (overlap) { toast(overlap, 'err'); return; }
    if (wiz.options.operation === 'move') {
      const ok = await Toolkit.confirm({
        title: 'Move these files?',
        body: 'Each source file goes to the source site recycle bin once it is on the target. Files that fail stay where they are.',
        items: [`From: ${src.siteUrl} / ${src.folder}`, `To: ${tgt.siteUrl} / ${tgt.folder}`],
        okLabel: 'Start move',
        danger: true,
      });
      if (!ok) return;
    }
    const loc = (c, sel) => ({ tenantId: c.tenantId, tenantHost: c.tenantHost, tenantName: c.tenantName, upn: c.upn, siteUrl: sel.siteUrl, siteTitle: sel.siteTitle, library: sel.library, folder: sel.folder });
    const config = {
      mode: wiz.mode,
      operation: wiz.options.operation,
      includeVersions: wiz.options.includeVersions,
      includeMetadata: wiz.options.includeMetadata,
      overwrite: wiz.options.overwrite,
      onlyNewer: wiz.options.onlyNewer,
      concurrency: wiz.options.concurrency,
      source: loc(sc, src),
      target: loc(tc, tgt),
    };
    const elevations = Jobs.takePendingElevations((e) =>
      (e.tenantId === sc.tenantId && sameSite(e.siteUrl, src.siteUrl)) || (e.tenantId === tc.tenantId && sameSite(e.siteUrl, tgt.siteUrl)));
    const job = Jobs.create(config, { elevations });
    Log.info(`Created job ${job.id}${elevations.length ? ` with ${elevations.length} site admin grant(s) on record` : ''}`);
    resetWizard();
    activateTab('jobs');
    startJobRun(job.id);
  }

  // ---------- location picker ----------
  function createPicker(container, role, onChange) {
    const state = { siteUrl: null, siteTitle: null, library: null, libraries: [], folder: null, root: null };
    container.innerHTML = `
      <div class="picker-site">
        <div class="pk-scope" role="group" aria-label="Find the site by">
          <button type="button" class="chip-btn active" data-scope="sites">${iconHtml('i-search', 'sm')}Sites</button>
          <button type="button" class="chip-btn" data-scope="onedrive">${iconHtml('i-user', 'sm')}OneDrive (by user)</button>
          <button type="button" class="chip-btn" data-scope="all">${iconHtml('i-grid', 'sm')}All sites (admin)</button>
        </div>
        <div class="row pk-inputs">
          <input class="pk-search" type="search" placeholder="Search sites by name" aria-label="Search" />
          <span class="muted">or</span>
          <input class="pk-url" type="text" placeholder="https://tenant.sharepoint.com/sites/Name" aria-label="Site URL" />
          <button class="btn pk-load" type="button">Load URL</button>
        </div>
        <p class="pk-scope-hint muted small" hidden></p>
        <div class="pk-results"></div>
        <div class="pk-siteinfo" hidden></div>
        <div class="pk-access" hidden></div>
      </div>
      <div class="picker-lib" hidden>
        <label>Document library <select class="pk-libsel"></select></label>
      </div>
      <div class="picker-folder" hidden>
        <div class="pk-breadcrumb"></div>
        <div class="pk-folders"></div>
        <p class="table-hint">The highlighted folder is the one that is used. Click a subfolder to go deeper.</p>
      </div>`;

    const q = (sel) => container.querySelector(sel);
    const searchEl = q('.pk-search');
    const urlEl = q('.pk-url');
    const resultsEl = q('.pk-results');
    const siteInfoEl = q('.pk-siteinfo');
    const accessEl = q('.pk-access');
    const libWrap = q('.picker-lib');
    const libSel = q('.pk-libsel');
    const folderWrap = q('.picker-folder');
    const crumbEl = q('.pk-breadcrumb');
    const foldersEl = q('.pk-folders');
    const hintEl = q('.pk-scope-hint');
    const conn = () => Auth.getConnection(role);

    let scope = 'sites';
    let allSites = null;
    let allSitesNote = '';
    let searchTimer = null;

    const SCOPE_UI = {
      sites: { placeholder: 'Search sites by name', minLen: 2, hint: '' },
      onedrive: { placeholder: 'Search a user by name or UPN', minLen: 2, hint: 'Opens the OneDrive of the user you pick. The first search may ask for consent to read basic user profiles.' },
      all: { placeholder: 'Filter every site collection', minLen: 0, hint: 'Lists every site in the tenant, which needs the SharePoint Administrator role. Sites you cannot open offer Take ownership.' },
    };

    const notice = (kind, ico, html) => `<div class="notice notice-${kind}">${iconHtml(ico)}<div>${html}</div></div>`;
    const muted = (text) => `<p class="muted small">${esc(text)}</p>`;

    container.querySelectorAll('.pk-scope .chip-btn').forEach((btn) => btn.addEventListener('click', () => setScope(btn.dataset.scope)));

    function setScope(next) {
      scope = next;
      container.querySelectorAll('.pk-scope .chip-btn').forEach((b) => b.classList.toggle('active', b.dataset.scope === next));
      const ui = SCOPE_UI[next];
      searchEl.value = '';
      searchEl.placeholder = ui.placeholder;
      hintEl.textContent = ui.hint;
      hintEl.hidden = !ui.hint;
      resultsEl.replaceChildren();
      if (next === 'all') loadAllSites();
    }

    searchEl.addEventListener('input', () => {
      clearTimeout(searchTimer);
      const term = searchEl.value.trim();
      if (term.length < SCOPE_UI[scope].minLen && scope !== 'all') { resultsEl.replaceChildren(); return; }
      searchTimer = setTimeout(() => runSearch(term), scope === 'all' ? 120 : 350);
    });
    q('.pk-load').addEventListener('click', () => { const url = urlEl.value.trim(); if (url) loadSite(url); });
    urlEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') q('.pk-load').click(); });

    function runSearch(term) {
      if (scope === 'onedrive') return runUserSearch(term);
      if (scope === 'all') return renderAllSites(term);
      return runSiteSearch(term);
    }

    function resultRow(title, sub, onClick) {
      return h('button.pk-result', { type: 'button', on: { click: onClick } }, [h('strong', { text: title }), h('span.muted.small', { text: sub })]);
    }
    function siteResultRow(displayName, webUrl) {
      return resultRow(displayName, webUrl, () => { resultsEl.replaceChildren(); searchEl.value = ''; loadSite(webUrl); });
    }

    async function runSiteSearch(term) {
      if (term.length < 2) return;
      resultsEl.innerHTML = muted('Searching...');
      try {
        const sites = await Graph.searchSites(role, term);
        if (!sites.length) { resultsEl.innerHTML = muted('No sites found.'); return; }
        resultsEl.replaceChildren(...sites.map((s) => siteResultRow(s.displayName, s.webUrl)));
      } catch (e) { resultsEl.innerHTML = notice('warn', 'i-alert', `Search failed: ${esc(e.message)}`); }
    }

    async function runUserSearch(term) {
      if (term.length < 2) return;
      resultsEl.innerHTML = muted('Searching users...');
      try {
        const users = await Graph.searchUsers(role, term);
        if (!users.length) { resultsEl.innerHTML = muted('No users found.'); return; }
        resultsEl.replaceChildren(...users.map((u) => {
          const row = resultRow(u.displayName, u.userPrincipalName || u.mail || '', () => openUserOneDrive(u, row));
          return row;
        }));
      } catch (e) {
        const msg = e.status === 403 ? 'You are not allowed to look up users in this tenant.' : e.message;
        resultsEl.innerHTML = notice('warn', 'i-alert', `User search failed: ${esc(msg)}`);
      }
    }

    async function openUserOneDrive(user, row) {
      await busy(row, async () => {
        try {
          const siteUrl = await Graph.getUserOneDriveSiteUrl(role, user);
          resultsEl.replaceChildren();
          searchEl.value = '';
          loadSite(siteUrl);
        } catch (e) {
          resultsEl.insertAdjacentHTML('afterbegin', notice('warn', 'i-alert', esc(e.message)));
        }
      });
    }

    async function loadAllSites() {
      if (allSites) { renderAllSites(searchEl.value.trim()); return; }
      resultsEl.innerHTML = muted('Loading all site collections (admin)...');
      try {
        allSites = await SharePoint.enumerateAllSites(role, {
          onProgress: (n) => { resultsEl.innerHTML = muted(`Loading all site collections (admin)... ${n} found`); },
        });
        allSitesNote = '';
        renderAllSites(searchEl.value.trim());
      } catch (e) {
        allSites = null;
        const denied = e.status === 401 || e.status === 403;
        const c = conn();
        // Without the admin role, fall back to the sites this account can see through Graph.
        if (denied && c.spoConsent) {
          try {
            resultsEl.innerHTML = muted('Admin site list not available, listing the sites your account can see...');
            const gsites = await Graph.listAllSites(role, (n) => { resultsEl.innerHTML = muted(`Listing sites your account can see... ${n} found`); });
            if (gsites.length) {
              allSites = gsites;
              allSitesNote = 'The tenant admin API is not available (it needs the SharePoint Administrator role), so this only lists sites your account can see.';
              renderAllSites(searchEl.value.trim());
              return;
            }
          } catch (ge) { Log.dbg(`Graph site fallback failed: ${ge.message}`); }
        }
        // A 403 after consent is a missing admin role, not missing consent: no grant button then.
        const needConsent = denied && !c.spoConsent;
        const msg = !denied ? esc(e.message)
          : needConsent ? 'Listing every site needs SharePoint access for the app in this tenant. Grant it and try again.'
            : `Listing every site needs the SharePoint or Global Administrator role. <strong>${esc(c.upn)}</strong> does not seem to have it (the admin API answered ${esc(e.status)}). Use the Sites tab and paste a site URL, or sign in with an admin account.`;
        resultsEl.innerHTML = notice('warn', 'i-alert', `${msg}${needConsent ? `<div class="row"><button class="btn btn-primary btn-small pk-grantspo" type="button">${iconHtml('i-key')}Grant SharePoint access</button></div>` : ''}`);
        const gb = resultsEl.querySelector('.pk-grantspo');
        if (gb) gb.addEventListener('click', () => busy(gb, async () => {
          try { await Auth.grantSharePointConsent(role); allSites = null; loadAllSites(); }
          catch (err) { toast(`Consent failed or was cancelled: ${err.message}`, 'err'); }
        }));
      }
    }

    function renderAllSites(term) {
      if (!allSites) return;
      const t = String(term || '').toLowerCase();
      const matched = t ? allSites.filter((s) => s.title.toLowerCase().includes(t) || s.url.toLowerCase().includes(t)) : allSites;
      if (!matched.length) { resultsEl.innerHTML = muted('No sites match.'); return; }
      const shown = matched.slice(0, 200);
      resultsEl.replaceChildren();
      if (allSitesNote) resultsEl.insertAdjacentHTML('beforeend', notice('info', 'i-info', esc(allSitesNote)));
      if (matched.length > shown.length) resultsEl.insertAdjacentHTML('beforeend', muted(`${matched.length} sites, showing the first ${shown.length}. Narrow the filter.`));
      shown.forEach((s) => resultsEl.appendChild(siteResultRow(s.title, s.url)));
    }

    async function loadSite(siteUrl) {
      accessEl.hidden = true;
      accessEl.replaceChildren();
      libWrap.hidden = true;
      folderWrap.hidden = true;
      state.library = null;
      state.folder = null;
      siteInfoEl.hidden = false;
      let url;
      try { url = new URL(siteUrl); } catch { url = null; }
      const c = conn();
      if (!url || url.protocol !== 'https:' || !MoveRules.isTenantSpoHost(url.host, c.tenantHost)) {
        siteInfoEl.innerHTML = notice('warn', 'i-alert', `That is not a SharePoint site of ${esc(c.tenantName || c.tenantHost)}. Use a URL on ${esc(c.tenantHost)} or its OneDrive host.`);
        state.siteUrl = null;
        fire();
        return;
      }
      siteInfoEl.innerHTML = muted('Loading site...');
      state.siteUrl = SharePoint.trimUrl(`${url.origin}${url.pathname}`);
      try {
        const web = await SharePoint.getWeb(role, state.siteUrl);
        state.siteTitle = web.Title || state.siteUrl;
        siteInfoEl.innerHTML = notice('info', 'i-cloud', `<strong>${esc(state.siteTitle)}</strong> <span class="muted small">${esc(state.siteUrl)}</span>`);
        await loadLibraries();
      } catch (e) {
        if (e.status === 401 || e.status === 403) showAccessDenied();
        else siteInfoEl.innerHTML = notice('warn', 'i-alert', `Could not open the site: ${esc(e.message)}`);
      }
      fire();
    }

    async function loadLibraries() {
      const libs = await SharePoint.getDocumentLibraries(role, state.siteUrl);
      state.libraries = libs;
      if (!libs.length) {
        libWrap.hidden = true;
        siteInfoEl.insertAdjacentHTML('beforeend', notice('warn', 'i-alert', `No document libraries found on this site. If this tenant has not approved SharePoint access for the app, grant it and try again.
          <div class="row"><button class="btn btn-primary btn-small pk-grantspo" type="button">${iconHtml('i-key')}Grant SharePoint access</button></div>`));
        const gb = siteInfoEl.querySelector('.pk-grantspo');
        if (gb) gb.addEventListener('click', () => busy(gb, async () => {
          try { await Auth.grantSharePointConsent(role); toast('SharePoint access granted. Reloading the site...', 'ok'); await loadSite(state.siteUrl); }
          catch (e) { toast(`Consent failed or was cancelled: ${e.message}`, 'err'); }
        }));
        return;
      }
      libSel.innerHTML = libs.map((l, i) => `<option value="${i}">${esc(l.title)} (${esc(l.itemCount)} items)</option>`).join('');
      libWrap.hidden = false;
      libSel.onchange = () => selectLibrary(Number(libSel.value));
      selectLibrary(0);
    }

    function selectLibrary(i) {
      state.library = state.libraries[i];
      state.root = state.library.serverRelativeUrl;
      openFolder(state.root);
    }

    async function openFolder(serverRel) {
      folderWrap.hidden = false;
      foldersEl.innerHTML = muted('Loading folders...');
      state.folder = serverRel;
      renderCrumb();
      fire();
      try {
        const children = await SharePoint.getFolderChildren(role, state.siteUrl, serverRel);
        if (state.folder !== serverRel) return;
        if (!children.folders.length) { foldersEl.innerHTML = muted('No subfolders. This folder is used.'); return; }
        foldersEl.replaceChildren(...children.folders.map((f) => h('button.pk-folder', { type: 'button', on: { click: () => openFolder(f.serverRelativeUrl) } },
          [icon('i-folder', 'sm'), h('span', { text: f.name }), h('span.muted.small', { text: `(${f.itemCount})` })])));
      } catch (e) { foldersEl.innerHTML = notice('warn', 'i-alert', `Could not list folders: ${esc(e.message)}`); }
    }

    function renderCrumb() {
      const crumbs = [{ name: state.library.title, path: state.root }];
      const rel = state.folder.startsWith(state.root) ? state.folder.substring(state.root.length) : '';
      let acc = state.root;
      rel.split('/').filter(Boolean).forEach((seg) => {
        acc += `/${seg}`;
        let name = seg;
        try { name = decodeURIComponent(seg); } catch { /* keep raw */ }
        crumbs.push({ name, path: acc });
      });
      crumbEl.replaceChildren();
      crumbs.forEach((c, i) => {
        if (i) crumbEl.appendChild(icon('i-chevron-right', 'sm'));
        crumbEl.appendChild(h('button', { type: 'button', class: `crumb${i === crumbs.length - 1 ? ' current' : ''}`, text: c.name, on: { click: () => openFolder(c.path) } }));
      });
    }

    function showAccessDenied() {
      const c = conn();
      siteInfoEl.hidden = true;
      accessEl.hidden = false;
      accessEl.innerHTML = notice('warn', 'i-lock', `<p><strong>${esc(c.upn)}</strong> has no access to the content of this site. A Global or SharePoint Administrator role does not give content access by itself.</p>
        <p>With the SharePoint Administrator role you can add yourself as site collection admin. SPOMove keeps a record of it and offers <strong>Remove my site admin rights</strong> on the job and in the Jobs tab; it stays in place until you remove it.</p>
        <div class="row"><button class="btn btn-primary btn-small pk-takeown" type="button">${iconHtml('i-shield')}Take ownership</button></div>
        <p class="pk-takeown-err small" hidden></p>`);
      const errEl = accessEl.querySelector('.pk-takeown-err');
      const btn = accessEl.querySelector('.pk-takeown');
      btn.addEventListener('click', () => takeOwnership(btn, errEl));
    }

    async function takeOwnership(btn, errEl) {
      const c = conn();
      const siteUrl = state.siteUrl;
      const ok = await Toolkit.confirm({
        title: 'Make yourself site collection admin?',
        body: 'SPOMove adds this account as site collection administrator through the SharePoint admin API. This is logged in the audit log and stays until you remove it with Remove my site admin rights.',
        items: [`Site: ${siteUrl}`, `Account: ${c.upn}`, `Tenant: ${c.tenantName || c.tenantHost}`],
        okLabel: 'Add me as site admin',
      });
      if (!ok) return;
      errEl.hidden = true;
      await busy(btn, async () => {
        // On record before the call, so a closed tab never leaves an unknown grant behind.
        const entry = Jobs.addPendingElevation({ siteUrl, upn: c.upn, tenantId: c.tenantId, tenantHost: c.tenantHost, tenantName: c.tenantName, role });
        try {
          await SharePoint.setSiteAdmin(role, siteUrl, SharePoint.buildClaimLoginName(c.upn), true);
          Log.info(`Added ${c.upn} as site collection admin of ${siteUrl}`);
          toast('You are site admin now. Reloading the site...', 'ok');
          renderJobs();
          await loadSite(siteUrl);
        } catch (e) {
          Jobs.dropPendingElevation(entry.id);
          errEl.hidden = false;
          errEl.textContent = e.status === 401 || e.status === 403
            ? `The tenant admin API refused it. This needs the SharePoint Administrator role and SharePoint access for the app in this tenant. (${e.message})`
            : `Take ownership failed: ${e.message}`;
        }
      });
    }

    function subLabel() {
      if (!state.folder || !state.root || state.folder === state.root) return '';
      const rel = state.folder.substring(state.root.length).replace(/^\/+/, '');
      try { return decodeURIComponent(rel); } catch { return rel; }
    }

    function getSelection() {
      if (!state.siteUrl || !state.library || !state.folder) return null;
      return { siteUrl: state.siteUrl, siteTitle: state.siteTitle, library: state.library.title, folder: state.folder, subLabel: subLabel() };
    }

    function fire() { if (onChange) onChange(); }

    return { getSelection };
  }

  // ---------- jobs dashboard ----------
  const cardRefs = new Map();
  const pausedJobs = new Set();

  let throttleTimer = null;
  function ensureThrottleTicker() {
    if (!throttleTimer) throttleTimer = setInterval(updateThrottleIndicators, 500);
  }
  function updateThrottleIndicators() {
    const st = SharePoint.getThrottleState();
    let anyRunning = false;
    for (const [id, el] of cardRefs) {
      const t = el.querySelector('.job-throttle');
      if (!t) continue;
      const show = Transfer.isRunning(id) && st.throttled;
      if (Transfer.isRunning(id)) anyRunning = true;
      t.hidden = !show;
      if (show) {
        t.querySelector('span').textContent = st.phase === 'recovering'
          ? 'Throttled by SharePoint, recovering...'
          : `Throttled by SharePoint, retrying in ${Math.ceil(st.waitMs / 1000)}s`;
      }
    }
    if (!anyRunning) { clearInterval(throttleTimer); throttleTimer = null; }
  }

  function renderJobs() {
    const list = Jobs.list();
    const wrap = $('jobsList');
    wrap.replaceChildren();
    cardRefs.clear();
    $('jobsEmpty').hidden = list.length > 0;
    $('clearAllBtn').disabled = list.length === 0;
    let active = 0;
    for (const job of list) {
      if (Transfer.isRunning(job.id)) active++;
      wrap.appendChild(buildJobCard(job));
    }
    const badge = $('jobsBadge');
    badge.hidden = active === 0;
    badge.textContent = String(active);
    renderPendingElevations();
  }

  function statusMeta(job) {
    const running = Transfer.isRunning(job.id);
    if (running && job.status === 'enumerating') return { cls: 'info', label: 'Scanning' };
    if (running) return { cls: 'info', label: pausedJobs.has(job.id) ? 'Paused' : 'Running' };
    switch (job.status) {
      case 'completed': return { cls: 'ok', label: 'Completed' };
      case 'completed_with_errors': return { cls: 'warn', label: 'Completed with errors' };
      case 'cancelled': return { cls: 'neutral', label: 'Cancelled' };
      case 'failed': return { cls: 'danger', label: 'Failed' };
      case 'running': case 'enumerating': return { cls: 'neutral', label: 'Interrupted' };
      default: return { cls: 'neutral', label: 'Ready' };
    }
  }

  function buildJobCard(job) {
    const c = job.config;
    const el = document.createElement('article');
    el.className = 'job-card';
    el.dataset.id = job.id;
    const unlisted = job.unlistedFolders || [];
    const elev = Jobs.activeElevations(job);
    el.innerHTML = `
      <div class="job-head">
        <span class="badge dot job-status"></span>
        <span class="job-title">${esc(c.source.siteTitle)} ${iconHtml('i-arrow-right', 'sm')} ${esc(c.target.siteTitle)}</span>
        <span class="job-mode muted small">${c.mode === 'cross' ? 'Between tenants' : 'Same tenant'} · ${c.operation === 'move' ? 'move' : 'copy'}</span>
      </div>
      <div class="job-sub muted small">${esc(c.source.folder)} ${iconHtml('i-arrow-right', 'sm')} ${esc(c.target.folder)}</div>
      ${job.lastError && job.status === 'failed' ? `<div class="notice notice-err">${iconHtml('i-error')}<div>${esc(job.lastError)}</div></div>` : ''}
      ${unlisted.length ? `<div class="notice notice-warn">${iconHtml('i-alert')}<div>${esc(plural(unlisted.length, 'source folder'))} could not be listed, so their files were not part of the last run and sync deleted nothing. See the report.</div></div>` : ''}
      ${elev.length ? `<div class="notice notice-warn">${iconHtml('i-shield')}<div>You are site collection admin on ${esc(plural(elev.length, 'site'))} through SPOMove for this job: ${esc(elev.map((e) => e.siteUrl).join(', '))}</div></div>` : ''}
      <div class="progress-box">
        <div class="progress-header"><span class="job-current"></span><span class="job-eta"></span></div>
        <div class="progress-bar"><div class="job-fill"></div></div>
        <div class="progress-footer"><span class="job-stats"></span><span class="job-speed"></span></div>
        <div class="job-throttle" hidden>${iconHtml('i-clock', 'sm')}<span></span></div>
      </div>
      <div class="job-actions"></div>`;
    cardRefs.set(job.id, el);
    updateJobCard(job, null, el);
    return el;
  }

  function updateJobCard(job, snap, el) {
    el = el || cardRefs.get(job.id);
    if (!el) return;
    const s = snap || {};
    const js = job.stats || {};
    const total = s.itemsTotal ?? js.total ?? 0;
    const done = s.itemsDone ?? js.done ?? 0;
    const failed = s.failed ?? js.failed ?? 0;
    const skipped = s.skipped ?? js.skipped ?? 0;
    const bytesTotal = s.bytesTotal ?? js.bytesTotal ?? 0;
    const bytesDone = s.bytesDone ?? js.bytesDone ?? 0;
    const pct = bytesTotal ? Math.round((bytesDone / bytesTotal) * 100) : (total ? Math.round(((done + skipped) / total) * 100) : 0);
    el.querySelector('.job-fill').style.width = `${Math.min(100, pct)}%`;
    const m = statusMeta(job);
    const st = el.querySelector('.job-status');
    st.textContent = m.label;
    st.className = `badge dot ${m.cls} job-status`;
    const versionsTotal = js.versionsTotal || 0;
    el.querySelector('.job-stats').textContent = `${done}/${total} files${versionsTotal ? ` +${versionsTotal} versions` : ''}${skipped ? ` · ${skipped} skipped` : ''}${failed ? ` · ${failed} failed` : ''}${bytesTotal ? ` · ${fmtBytes(bytesDone)} / ${fmtBytes(bytesTotal)}` : ''}`;
    el.querySelector('.job-speed').textContent = s.speedBps ? `${fmtBytes(s.speedBps)}/s` : '';
    el.querySelector('.job-current').textContent = s.currentFile ? `Current: ${s.currentFile}` : '';
    let eta = s.etaSec != null ? `about ${fmtDuration(s.etaSec)} left` : '';
    if (!eta && snap && Transfer.isRunning(job.id) && job.startedAt) eta = Concurrency.eta(job.startedAt, done + failed + skipped, total);
    el.querySelector('.job-eta').textContent = eta;
    renderJobActions(job, el);
  }

  function actionBtn(act, label, ico, variant) {
    return `<button type="button" class="btn btn-small${variant ? ` btn-${variant}` : ''}" data-act="${act}">${ico ? iconHtml(ico) : ''}${esc(label)}</button>`;
  }

  function renderJobActions(job, el) {
    const box = el.querySelector('.job-actions');
    if (box.dataset.mode === 'rerun' && !Transfer.isRunning(job.id)) return;
    box.dataset.mode = '';
    const running = Transfer.isRunning(job.id);
    const s = job.stats || {};
    const b = [];
    if (running) {
      b.push(pausedJobs.has(job.id) ? actionBtn('pause', 'Continue', 'i-play') : actionBtn('pause', 'Pause', 'i-pause'));
      b.push(actionBtn('cancel', 'Cancel', 'i-stop'));
    } else {
      const hasPending = (s.total || 0) - (s.done || 0) - (s.failed || 0) - (s.skipped || 0) - (s.deleted || 0) > 0;
      if (job.status === 'draft') b.push(actionBtn('resume', 'Start', 'i-play', 'primary'));
      else if (hasPending) b.push(actionBtn('resume', 'Resume', 'i-play', 'primary'));
      if ((s.failed || 0) > 0) b.push(actionBtn('retry', `Retry failed (${s.failed})`, 'i-refresh', 'primary'));
      const hasRun = ((s.done || 0) + (s.skipped || 0) + (s.deleted || 0)) > 0;
      if (job.status !== 'draft' && !hasPending && hasRun) b.push(actionBtn('rerun', 'Re-run', 'i-refresh'));
      const elev = Jobs.activeElevations(job).length;
      if (elev) b.push(actionBtn('revoke', `Remove my site admin rights (${elev})`, 'i-shield'));
      b.push(actionBtn('delete', 'Delete job', 'i-trash'));
    }
    if (job.status !== 'draft' && (s.total || 0) > 0) b.push(actionBtn('report', 'Report', 'i-chart'));
    box.innerHTML = b.join('');
  }

  function showRerunChoice(id) {
    const el = cardRefs.get(id);
    if (!el) return;
    const box = el.querySelector('.job-actions');
    box.dataset.mode = 'rerun';
    const job = Jobs.get(id);
    const isMove = job && job.config.operation === 'move';
    box.innerHTML = `<span class="muted small">Re-run:</span>
      <button type="button" class="btn btn-small btn-primary" data-act="sync" title="Transfer new and changed files.${isMove ? '' : ' Files gone from the source are listed for deletion and only deleted after you confirm.'}">${iconHtml('i-refresh')}Sync changes</button>
      <button type="button" class="btn btn-small" data-act="full" title="Transfer every file that is in the source now.">Full re-run</button>
      <button type="button" class="btn btn-small btn-ghost" data-act="rerun-cancel">Cancel</button>`;
  }

  function onJobsClick(e) {
    const btn = e.target.closest('button[data-act]');
    if (!btn) return;
    const card = btn.closest('.job-card');
    if (!card) return;
    jobAction(card.dataset.id, btn.dataset.act, btn);
  }

  async function jobAction(id, act, btn) {
    const job = Jobs.get(id);
    if (!job) { renderJobs(); return; }
    const el = cardRefs.get(id);
    switch (act) {
      case 'resume': return startJobRun(id, { action: 'resume' });
      case 'retry': return startJobRun(id, { action: 'retry', autoReport: false });
      case 'pause':
        if (pausedJobs.has(id)) { Transfer.resumeRun(id); pausedJobs.delete(id); } else { Transfer.pause(id); pausedJobs.add(id); }
        return updateJobCard(job, null);
      case 'cancel':
        Transfer.cancel(id);
        pausedJobs.delete(id);
        if (el) {
          const st = el.querySelector('.job-status');
          st.textContent = 'Cancelling';
          st.className = 'badge dot warn job-status';
          el.querySelectorAll('.job-actions button').forEach((b) => { b.disabled = true; });
        }
        return undefined;
      case 'report': return openJobReport(id);
      case 'rerun': return showRerunChoice(id);
      case 'rerun-cancel':
        if (el) { el.querySelector('.job-actions').dataset.mode = ''; renderJobActions(job, el); }
        return undefined;
      case 'sync': return startRerun(id, 'incremental');
      case 'full': return startRerun(id, 'full');
      case 'revoke': return busy(btn, () => revokeElevations(Jobs.activeElevations(job), id));
      case 'delete': return deleteJob(job);
      default: return undefined;
    }
  }

  async function deleteJob(job) {
    if (Transfer.isRunning(job.id)) { toast('Cancel the job before deleting it.', 'warn'); return; }
    const elev = Jobs.activeElevations(job);
    const ok = await Toolkit.confirm({
      title: 'Delete this job?',
      body: `Removes the job and its resume data from this browser. Files already transferred stay on the target.${elev.length ? ' The site admin rights it added stay on record in the Jobs tab, so you can still remove them.' : ''}`,
      items: [`${job.config.source.siteTitle} -> ${job.config.target.siteTitle}`, `${job.config.source.folder} -> ${job.config.target.folder}`],
      okLabel: 'Delete job',
      danger: true,
    });
    if (!ok) return;
    Jobs.remove(job.id);
    renderJobs();
  }

  function startRerun(id, mode) {
    const job = Jobs.get(id);
    if (!job) return;
    job.rerun = mode;
    job.status = 'enumerating';
    job.startedAt = null;
    job.finishedAt = null;
    Jobs.save(job);
    Log.info(`Re-run (${mode}) for job ${String(id).slice(0, 8)}: re-scanning the source.`);
    startJobRun(id, { action: 'rerun' });
  }

  // ---------- connections for jobs ----------
  async function ensureJobConnections(job) {
    const c = job.config;
    await ensureRoleConnection('source', c.source);
    if (c.mode === 'cross') await ensureRoleConnection('target', c.target);
    // Someone signed in with another account of the same tenant: the job now runs as that one.
    let changed = false;
    for (const [key, role] of [['source', 'source'], ['target', c.mode === 'cross' ? 'target' : 'source']]) {
      const upn = Auth.getConnection(role).upn;
      if (c[key].upn && upn && c[key].upn.toLowerCase() !== upn.toLowerCase()) {
        Log.warn(`Job ${String(job.id).slice(0, 8)} ${key} account changed from ${c[key].upn} to ${upn}.`);
        c[key].upn = upn;
        changed = true;
      }
    }
    if (changed) Jobs.save(job);
  }

  // Reuse a matching live connection, else re-attach the exact account (tenant + UPN) this tab
  // signed in, else ask for a sign-in with the job's account as hint and check the tenant.
  async function ensureRoleConnection(role, loc, { requireUpn = true } = {}) {
    if (!Auth.connectionMatches(role, loc, { requireUpn })) {
      assertRoleFree(role);
      const restored = await Auth.restoreConnection(role, loc);
      if (!restored) await promptReconnect(role, loc);
    }
    const conn = Auth.getConnection(role);
    if (conn && !conn.spoConsent) {
      if (await Auth.hasSharePointConsent(role)) conn.spoConsent = true;
      else {
        toast(`Approve SharePoint access for the ${role} tenant to continue.`, 'info');
        await Auth.grantSharePointConsent(role);
      }
    }
    updateUserBox();
    updateWizard();
  }

  // Explicit "which tenant" dialog; the sign-in popup opens from its button click.
  async function promptReconnect(role, loc) {
    const isSrc = role === 'source';
    const tenant = loc.tenantName || loc.tenantHost || 'the tenant of this job';
    let error = '';
    for (;;) {
      const body = h('div.reconnect', null, [
        h('p', null, h('span', { class: `badge ${isSrc ? 'info' : 'ok'}`, text: isSrc ? 'Source tenant' : 'Target tenant' })),
        h('p', null, ['Sign in to ', h('strong', { text: tenant }), loc.tenantHost ? h('span.muted', { text: ` (${loc.tenantHost})` }) : null, '.']),
        loc.upn ? h('p', null, ['The job ran as ', h('strong', { text: loc.upn }), '. The sign-in window opens with that account filled in.']) : null,
        error ? h('div.notice.notice-err', null, [icon('i-error'), h('div', { text: error })]) : null,
      ]);
      const choice = await Toolkit.modal({
        title: 'Reconnect tenant',
        body,
        actions: [
          { label: 'Cancel', value: null },
          { label: `Sign in to the ${isSrc ? 'source' : 'target'} tenant`, value: 'go', variant: 'primary', icon: 'i-login' },
        ],
      });
      if (choice !== 'go') throw new Error('Reconnect cancelled');
      try {
        await Auth.signIn(role, { loginHint: loc.upn, expectTenantId: loc.tenantId, expectTenantHost: loc.tenantHost });
        return;
      } catch (e) {
        error = e.message;
      }
    }
  }

  async function confirmSyncDeletes(job, paths, share) {
    const t = job.config.target;
    const pct = Math.round(share * 100);
    const body = h('div', null, [
      h('p', { text: `This job copied these files to the target before and they are no longer in the source. Sync moves them to the recycle bin of ${t.siteTitle} (${t.tenantName || t.tenantHost}).` }),
      share >= 0.25 ? h('div.notice.notice-warn', null, [icon('i-alert'), h('div', { text: `That is ${pct}% of what this job put on the target. If the source folder was moved or renamed, or you lost access to part of it, keep them.` })]) : null,
      h('p.muted.small', { text: 'Keep them: new and changed files are still transferred and nothing is deleted.' }),
    ]);
    activateTab('jobs');
    return Toolkit.confirm({
      title: `Delete ${plural(paths.length, 'file')} from the target?`,
      body, items: paths, itemsMax: 1000, wide: true, danger: true,
      okLabel: `Delete ${paths.length}`, cancelLabel: 'Keep them',
    });
  }

  async function startJobRun(id, opts = {}) {
    const pre = Jobs.get(id);
    if (!pre) return;
    const el0 = cardRefs.get(id);
    try {
      if (el0) el0.querySelector('.job-current').textContent = 'Checking tenant connections...';
      await ensureJobConnections(pre);
    } catch (e) {
      toast(`Cannot start: ${e.message}`, 'err');
      Log.err('Reconnect failed:', e);
      renderJobs();
      return;
    }
    if (pre.status !== 'draft' && (opts.action === 'resume' || opts.action === 'retry')) {
      const s = pre.stats || {};
      const failed = s.failed || 0;
      const pending = Math.max(0, (s.total || 0) - (s.done || 0) - failed - (s.skipped || 0) - (s.deleted || 0));
      const parts = [];
      if (failed) parts.push(`${failed} failed before`);
      if (pending) parts.push(`${pending} not transferred yet`);
      Log.info(`${opts.action === 'retry' ? 'Retry' : 'Resume'} job ${String(id).slice(0, 8)}: ${parts.length ? parts.join(' + ') : 'nothing left'}.`);
      if (opts.action === 'retry' && pending) toast(`Retrying ${failed} failed file(s) and ${pending} not transferred yet.`, 'info');
    }
    Transfer.run(id, {
      onEnumerate: (n) => { const el = cardRefs.get(id); if (el) el.querySelector('.job-current').textContent = `Scanning the source... ${n} files found`; },
      onVersions: (done, total) => { const el = cardRefs.get(id); if (el) el.querySelector('.job-current').textContent = `Reading version history... ${done}/${total}`; },
      onProgress: (snap, job) => updateJobCard(job, snap),
      confirmDeletes: confirmSyncDeletes,
    }).then(() => {
      pausedJobs.delete(id);
      renderJobs();
      afterJob(id);
      if (opts.autoReport !== false) maybeAutoReport(id);
      refreshOpenReport(id);
    }).catch((e) => {
      pausedJobs.delete(id);
      if (!/^Cancelled$/.test(e.message || '')) toast(`Job error: ${e.message}`, 'err');
      renderJobs();
      afterJob(id);
      refreshOpenReport(id);
    });
    renderJobs();
    ensureThrottleTicker();
  }

  function afterJob(id) {
    const job = Jobs.get(id);
    const n = job ? Jobs.activeElevations(job).length : 0;
    if (n) toast(`You are still site admin on ${plural(n, 'site')} for this job. Use Remove my site admin rights when you are done.`, 'warn', { timeout: 10000 });
  }

  // ---------- site admin rights added by SPOMove ----------
  async function roleForTenant(entry) {
    for (const r of ['source', 'target']) if (Auth.connectionMatches(r, entry, { requireUpn: false })) return r;
    const role = entry.role === 'target' ? 'target' : 'source';
    await ensureRoleConnection(role, entry, { requireUpn: false });
    return role;
  }

  async function revokeElevations(entries, jobId) {
    const active = (entries || []).filter((e) => !e.revokedAt);
    if (!active.length) return;
    if (jobId && Transfer.isRunning(jobId)) { toast('Wait for the job to finish first.', 'warn'); return; }
    const ok = await Toolkit.confirm({
      title: 'Remove your site admin rights?',
      body: 'SPOMove removes the site collection admin rights it added on these sites. Nothing else changes.',
      items: active.map((e) => `${e.siteUrl}  (${e.upn}, ${e.tenantName || e.tenantHost})`),
      okLabel: `Remove from ${plural(active.length, 'site')}`,
    });
    if (!ok) return;
    let done = 0;
    let failed = 0;
    for (const e of active) {
      try {
        const role = await roleForTenant(e);
        await SharePoint.setSiteAdmin(role, e.siteUrl, SharePoint.buildClaimLoginName(e.upn), false);
        Jobs.markElevationRevoked(e.id, jobId);
        Log.info(`Removed ${e.upn} as site collection admin of ${e.siteUrl}`);
        done++;
      } catch (err) {
        failed++;
        Log.err(`Could not remove site admin rights on ${e.siteUrl}:`, err);
        if (/Reconnect cancelled/.test(err.message || '')) break;
      }
    }
    toast(failed ? `Removed on ${done} site(s), ${failed} failed. See the debug log.` : `Removed your site admin rights on ${plural(done, 'site')}.`, failed ? 'warn' : 'ok');
    renderJobs();
    refreshOpenReport(jobId);
  }

  function renderPendingElevations() {
    const list = Jobs.pendingElevations();
    $('elevCard').hidden = !list.length;
    if (!list.length) return;
    $('elevTable').querySelector('tbody').innerHTML = list.map((e) => `<tr>
      <td class="url-cell">${esc(e.siteUrl)}</td>
      <td>${esc(e.upn)}</td>
      <td>${esc(e.tenantName || e.tenantHost || '')}</td>
      <td class="nowrap">${esc(new Date(e.grantedAt).toLocaleString())}</td>
    </tr>`).join('');
  }

  // ---------- sign-out and stored data ----------
  async function onSignOut() {
    if (Transfer.rolesInUse().size) { toast('Cancel or finish the running jobs before signing out.', 'warn'); return; }
    const choice = await Toolkit.modal({
      title: 'Sign out',
      body: 'Signs out the source and target connections in this tab. Job history (settings, site URLs, file lists) stays in this browser unless you clear it as well.',
      actions: [
        { label: 'Cancel', value: null },
        { label: 'Sign out and clear job data', value: 'clear', variant: 'danger', icon: 'i-trash' },
        { label: 'Sign out', value: 'out', variant: 'primary', icon: 'i-logout' },
      ],
    });
    if (!choice) return;
    if (choice === 'clear') clearJobData();
    await Auth.signOutAll();
    resetWizard();
    updateUserBox();
    toast('Signed out.', 'ok');
  }

  function clearJobData() {
    const n = Jobs.list().length;
    Jobs.clearAll();
    renderJobs();
    Log.info(`Cleared ${plural(n, 'job')} from this browser.`);
  }

  async function onClearAll() {
    const jobs = Jobs.list();
    if (!jobs.length) return;
    if (jobs.some((j) => Transfer.isRunning(j.id))) { toast('Cancel the running jobs first.', 'warn'); return; }
    const ok = await Toolkit.confirm({
      title: 'Clear all job data?',
      body: 'Removes every job, its file list and resume data from this browser. Resuming is not possible afterwards. Site admin rights SPOMove added and you did not remove yet stay on record.',
      items: jobs.map((j) => `${j.config.source.siteTitle} -> ${j.config.target.siteTitle} (${statusMeta(j).label})`),
      okLabel: 'Clear all',
      danger: true,
    });
    if (ok) clearJobData();
  }

  function fmtDuration(sec) {
    if (sec == null) return '';
    sec = Math.round(sec);
    if (sec < 60) return `${sec}s`;
    if (sec < 3600) return `${Math.floor(sec / 60)}m ${sec % 60}s`;
    return `${Math.floor(sec / 3600)}h ${Math.floor((sec % 3600) / 60)}m`;
  }

  // ---------- job report ----------
  let reportState = null; // { jobId, filter, term, items, job, node }

  async function openJobReport(id) {
    const job = Jobs.get(id);
    if (!job) { toast('Job not found.', 'err'); return; }
    const node = h('div.report', null, h('p.muted', { text: 'Loading report...' }));
    reportState = { jobId: id, filter: 'all', term: '', items: [], job, node };
    const closed = Toolkit.modal({ title: 'Transfer report', body: node, wide: true });
    try { reportState.items = await Jobs.loadItems(id); } catch (e) { Log.warn('Report load failed:', e.message); }
    if (reportState && reportState.node === node) renderReport();
    await closed;
    if (reportState && reportState.node === node) reportState = null;
  }

  async function refreshOpenReport(id) {
    if (!reportState || reportState.jobId !== id || !reportState.node.isConnected) return;
    try { reportState.items = await Jobs.loadItems(id); reportState.job = Jobs.get(id) || reportState.job; }
    catch (e) { Log.warn('Report refresh failed:', e.message); }
    if (reportState && reportState.jobId === id && reportState.node.isConnected) renderReport();
  }

  function maybeAutoReport(id) {
    const job = Jobs.get(id);
    if (!job || !['completed', 'completed_with_errors', 'failed'].includes(job.status)) return;
    if (reportState && reportState.node.isConnected) return;
    if (!(job.stats && job.stats.total)) return;
    openJobReport(id);
  }

  function donutSvg(parts) {
    const total = Math.max(1, parts.reduce((a, p) => a + p.value, 0));
    const R = 52;
    const C = 2 * Math.PI * R;
    let acc = 0;
    const segs = parts.filter((p) => p.value > 0).map((p) => {
      const len = (p.value / total) * C;
      const seg = `<circle class="${p.cls}" r="${R}" cx="70" cy="70" fill="none" stroke-width="18" stroke-dasharray="${len} ${C - len}" stroke-dashoffset="${-acc}" transform="rotate(-90 70 70)"></circle>`;
      acc += len;
      return seg;
    });
    const ok = parts.find((p) => p.cls === 'd-ok');
    const pct = Math.round(((ok ? ok.value : 0) / total) * 100);
    return `<svg class="donut" viewBox="0 0 140 140" width="150" height="150" role="img" aria-label="${pct}% succeeded">
      <circle class="d-track" r="${R}" cx="70" cy="70" fill="none" stroke-width="18"></circle>
      ${segs.join('')}
      <text class="d-pct" x="70" y="68" text-anchor="middle">${pct}%</text>
      <text class="d-label" x="70" y="88" text-anchor="middle">success</text>
    </svg>`;
  }

  const ST_LABEL = { done: 'OK', failed: 'Failed', skipped: 'Skipped', deleted: 'Deleted', pending: 'Pending' };
  const ST_BADGE = { done: 'ok', failed: 'danger', skipped: 'neutral', deleted: 'violet', pending: 'neutral' };
  const itemStatus = (it) => (ST_LABEL[it.status] ? it.status : 'pending');
  const itemDetail = (it) => {
    const st = itemStatus(it);
    if (st === 'skipped') return it.skipReason || 'Skipped';
    if (st === 'deleted') return 'Moved to the target recycle bin (gone from the source)';
    if (it.op === 'delete' && st === 'pending') return 'Sync delete, not done yet';
    return it.error || '';
  };

  function renderReport() {
    if (!reportState) return;
    const job = Jobs.get(reportState.jobId) || reportState.job;
    if (!job) return;
    const items = reportState.items || [];
    const s = job.stats || {};
    const count = (st) => items.filter((i) => itemStatus(i) === st).length;
    const total = items.length || s.total || 0;
    const done = items.length ? count('done') : (s.done || 0);
    const failed = items.length ? count('failed') : (s.failed || 0);
    const skipped = items.length ? count('skipped') : (s.skipped || 0);
    const deleted = items.length ? count('deleted') : (s.deleted || 0);
    const pending = Math.max(0, total - done - failed - skipped - deleted);
    const m = statusMeta(job);
    const durationMs = job.finishedAt && job.startedAt ? job.finishedAt - job.startedAt : (job.startedAt ? Date.now() - job.startedAt : 0);
    const durationSec = durationMs / 1000;
    const skippedBytes = items.filter((i) => i.status === 'skipped').reduce((a, b) => a + (b.size || 0) + (b.versionBytes || 0), 0);
    const avgBps = durationSec > 0 ? Math.max(0, (s.bytesDone || 0) - skippedBytes) / durationSec : 0;
    const retried = items.filter((i) => (i.attempts || 0) > 1).length;
    const versionsTotal = items.reduce((a, b) => a + (b.versionCount || 0), 0) || (s.versionsTotal || 0);
    const unlisted = job.unlistedFolders || [];
    const elev = Jobs.activeElevations(job);
    const legend = (cls, label, n) => `<div class="rl-row"><span class="rl-dot ${cls}"></span>${esc(label)} <strong>${esc(n)}</strong></div>`;

    const parts = [];
    parts.push(`<div class="report-summary">
      <div class="report-chart">${donutSvg([
        { cls: 'd-ok', value: done }, { cls: 'd-fail', value: failed }, { cls: 'd-del', value: deleted },
        { cls: 'd-skip', value: skipped }, { cls: 'd-pend', value: pending }])}</div>
      <div class="report-legend">
        ${legend('d-ok', 'Succeeded', done)}${legend('d-fail', 'Failed', failed)}
        ${skipped ? legend('d-skip', 'Skipped', skipped) : ''}${deleted ? legend('d-del', 'Deleted on target', deleted) : ''}
        ${legend('d-pend', 'Not transferred', pending)}
      </div>
      <dl class="report-stats">
        <dt>Status</dt><dd><span class="badge dot ${m.cls}">${esc(m.label)}</span></dd>
        <dt>Operation</dt><dd>${esc(job.config.operation)} · ${job.config.mode === 'cross' ? 'between tenants' : 'same tenant'}</dd>
        <dt>Files</dt><dd>${esc(done)} of ${esc(total)}${skipped ? ` · ${esc(skipped)} skipped` : ''}${deleted ? ` · ${esc(deleted)} deleted` : ''}</dd>
        <dt>Data</dt><dd>${fmtBytes(s.bytesDone || 0)} / ${fmtBytes(s.bytesTotal || 0)}${versionsTotal ? ' (incl. versions)' : ''}</dd>
        ${versionsTotal ? `<dt>Versions</dt><dd>${esc(versionsTotal)} included</dd>` : ''}
        <dt>Avg speed</dt><dd>${avgBps ? `${fmtBytes(avgBps)}/s` : '-'}</dd>
        <dt>Duration</dt><dd>${durationMs ? fmtDuration(durationSec) : '-'}</dd>
        ${retried ? `<dt>Retried</dt><dd>${esc(retried)} file(s) took more than one try</dd>` : ''}
        <dt>Route</dt><dd>${esc(job.config.source.siteTitle)} ${iconHtml('i-arrow-right', 'sm')} ${esc(job.config.target.siteTitle)}</dd>
      </dl>
    </div>`);
    if (unlisted.length) {
      parts.push(`<div class="notice notice-warn">${iconHtml('i-alert')}<div><p><strong>${esc(plural(unlisted.length, 'source folder'))} could not be listed.</strong> Their files were not transferred and a sync deleted nothing on the target. Fix access and run Sync changes again.</p>
        <ul class="unlisted">${unlisted.slice(0, 50).map((f) => `<li><code>${esc(f.folder)}</code> <span class="muted small">${esc(f.error || '')}</span></li>`).join('')}</ul>
        ${unlisted.length > 50 ? `<p class="muted small">+ ${esc(unlisted.length - 50)} more in the debug log</p>` : ''}</div></div>`);
    }
    if (elev.length) {
      parts.push(`<div class="notice notice-warn">${iconHtml('i-shield')}<div><p>You are site collection admin on ${esc(plural(elev.length, 'site'))} because SPOMove added you for this job: ${esc(elev.map((e) => e.siteUrl).join(', '))}</p>
        <div class="row"><button type="button" class="btn btn-small btn-primary" data-rep="revoke">${iconHtml('i-shield')}Remove my site admin rights</button></div></div></div>`);
    }
    const counts = { all: items.length, failed, skipped, deleted, done, pending };
    const keys = ['all', 'failed'].concat(skipped ? ['skipped'] : []).concat(deleted ? ['deleted'] : []).concat(['done', 'pending']);
    parts.push(`<div class="report-controls">
      <div class="row tight rf-group">${keys.map((f) => `<button type="button" class="chip-btn${reportState.filter === f ? ' active' : ''}" data-filter="${f}">${f[0].toUpperCase() + f.slice(1)} (${esc(counts[f] || 0)})</button>`).join('')}</div>
      <input type="search" class="report-search" placeholder="Filter on file path" aria-label="Filter on file path" />
      <button type="button" class="btn btn-small" data-rep="xlsx">${iconHtml('i-table')}Excel</button>
      <button type="button" class="btn btn-small" data-rep="csv">${iconHtml('i-download')}CSV</button>
    </div>
    <div class="report-rows"></div>`);

    const node = reportState.node;
    node.innerHTML = parts.join('');
    const search = node.querySelector('.report-search');
    search.value = reportState.term;
    search.addEventListener('input', () => { reportState.term = search.value; renderReportRows(); });
    node.querySelectorAll('[data-filter]').forEach((b) => b.addEventListener('click', () => {
      reportState.filter = b.dataset.filter;
      node.querySelectorAll('[data-filter]').forEach((x) => x.classList.toggle('active', x === b));
      renderReportRows();
    }));
    node.querySelectorAll('[data-rep]').forEach((b) => b.addEventListener('click', () => {
      const act = b.dataset.rep;
      if (act === 'csv') exportReport(job, 'csv');
      else if (act === 'xlsx') busy(b, () => exportReport(job, 'xlsx'));
      else if (act === 'revoke') busy(b, () => revokeElevations(Jobs.activeElevations(Jobs.get(job.id) || job), job.id));
    }));
    renderReportRows();
  }

  function currentReportItems() {
    const items = (reportState && reportState.items) || [];
    const t = reportState.term.trim().toLowerCase();
    return items.filter((it) => {
      const st = itemStatus(it);
      if (reportState.filter !== 'all' && st !== reportState.filter) return false;
      return !t || String(it.subPath || '').toLowerCase().includes(t);
    });
  }

  function renderReportRows() {
    if (!reportState) return;
    const wrap = reportState.node.querySelector('.report-rows');
    if (!wrap) return;
    const filtered = currentReportItems();
    if (!filtered.length) { wrap.innerHTML = `<div class="empty-state">${iconHtml('i-search')}<p>No items match.</p></div>`; return; }
    const capped = filtered.slice(0, 500);
    const rows = capped.map((it) => {
      const st = itemStatus(it);
      const detail = itemDetail(it);
      return `<tr>
        <td><span class="badge ${ST_BADGE[st]}">${ST_LABEL[st]}</span></td>
        <td class="rt-path" title="${esc(it.subPath || '')}">${esc(it.subPath || '')}</td>
        <td class="nowrap num">${fmtBytes(it.size || 0)}</td>
        <td class="num">${esc(it.attempts || '')}</td>
        <td class="rt-detail${st === 'failed' ? ' is-err' : ''}" title="${esc(detail)}">${esc(detail)}</td>
      </tr>`;
    }).join('');
    wrap.innerHTML = `
      ${filtered.length > capped.length ? `<p class="table-hint">Showing the first ${capped.length} of ${filtered.length}. Narrow the filter or export for the full list.</p>` : ''}
      <div class="table-wrap report-table-scroll">
        <table class="data-table">
          <thead><tr><th>Status</th><th>File</th><th class="num">Size</th><th class="num">Tries</th><th>Details</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  }

  const EXPORT_COLUMNS = [
    { key: 'status', label: 'Status', format: (v, r) => ST_LABEL[itemStatus(r)] },
    { key: 'subPath', label: 'File' },
    { key: 'size', label: 'Size (bytes)', format: (v) => v || 0 },
    { key: 'versionCount', label: 'Versions', format: (v) => v || 0 },
    { key: 'versionBytes', label: 'Version bytes', format: (v) => v || 0 },
    { key: 'attempts', label: 'Attempts', format: (v) => v || 0 },
    { key: 'error', label: 'Details', format: (v, r) => itemDetail(r) },
  ];

  async function exportReport(job, kind) {
    const rows = (reportState && reportState.items) || [];
    if (!rows.length) { toast('Nothing to export.', 'warn'); return; }
    const name = `SPOMove-report-${String(job.id).slice(0, 8)}-${Export.stamp()}`;
    try {
      if (kind === 'xlsx') await Export.xlsx([{ name: 'Files', rows, columns: EXPORT_COLUMNS }], `${name}.xlsx`);
      else Export.csv(rows, { columns: EXPORT_COLUMNS, filename: `${name}.csv` });
    } catch (e) { toast(`Export failed: ${e.message}`, 'err'); }
  }

  // ---------- boot ----------
  async function boot() {
    if (ToolkitAuth.isAuthPopup()) return;
    activateTab = Toolkit.tabs('#tabs', { onChange: (id) => { if (id === 'jobs') renderJobs(); } });
    initWizard();
    $('newJobBtn').addEventListener('click', () => { resetWizard(); activateTab('wizard'); });
    $('clearFinishedBtn').addEventListener('click', () => { Jobs.clearFinished(); renderJobs(); });
    $('clearAllBtn').addEventListener('click', onClearAll);
    $('revokePendingBtn').addEventListener('click', (e) => busy(e.currentTarget, () => revokeElevations(Jobs.pendingElevations(), null)));
    $('jobsList').addEventListener('click', onJobsClick);
    renderJobs();
    try {
      await Auth.init();
      updateUserBox();
      updateWizard();
      Log.info('SPOMove ready.');
    } catch (e) {
      $('bootError').hidden = false;
      $('bootErrorText').textContent = `Start-up problem: ${e.message}`;
      Log.err('Start-up failed:', e);
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
})();
