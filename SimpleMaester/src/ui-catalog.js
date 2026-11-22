// Catalog UI: two modes
//  - default: a grid of category tiles ("Authentication Methods", "Conditional
//    Access", ...) plus a "Custom selection" tile.
//  - custom:  the existing checkbox catalog for power users.
// The standalone "Test catalog" tab is the read-only browser.
(() => {
  let _all = [];
  let _selected = new Set();
  const LS_LAST_SELECTION = 'SimpleMaester:lastCustomSelection';

  function _loadLastSelection() {
    try {
      const raw = localStorage.getItem(LS_LAST_SELECTION);
      if (!raw) return null;
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : null;
    } catch { return null; }
  }
  function _saveLastSelection(ids) {
    try { localStorage.setItem(LS_LAST_SELECTION, JSON.stringify(ids)); }
    catch (e) { Log.warn('Could not persist last custom selection:', e.message); }
  }

  // Per-category description + icon. Anything we don't have an explicit entry
  // for falls back to a generic shield icon and a generated description.
  // Category names match Maester's upstream tag taxonomy verbatim
  // (see https://maester.dev/docs/tests/ - the "Tag" column).
  const CATEGORY_META = {
    'CA':             { icon: 'i-lock', desc: 'Conditional Access. Block legacy auth, require MFA, exclude break-glass, scope by risk and device.' },
    'App':            { icon: 'i-code', desc: 'Applications and service principals. Owners, secrets, high-risk permissions, broad assignments.' },
    'Privileged':     { icon: 'i-shield', desc: 'Privileged Identity Management and admin role hygiene. PIM eligibility, GA count, cloud-only admins.' },
    'Authentication': { icon: 'i-key', desc: 'Authentication methods. Phishing-resistant MFA, FIDO2, MS Authenticator, weak methods like SMS/voice.' },
    'Entra':          { icon: 'i-user', desc: 'General Entra ID hygiene. Stale guests, administrative units, tenant settings, group creation.' },
    'General':        { icon: 'i-check-circle', desc: 'EIDSCA - Entra ID Security Config Analyzer. 44 baseline configuration controls.' },
    'CISA':           { icon: 'i-book', desc: 'CISA SCuBA baseline. Entra ID controls, mail DNS records (SPF, DKIM, DMARC) and SharePoint sharing.' },
    'CIS':            { icon: 'i-book', desc: 'CIS Microsoft 365 Foundations Benchmark v6.0.1 - 38 controls across identity, mail, devices and apps.' },
    // The following Maester tags exist upstream but we have no implemented tests for them yet.
    // They will only show up if/when tests are added.
    'Teams':          { icon: 'i-user', desc: 'Microsoft Teams meeting and collaboration controls.' },
    'Exchange':       { icon: 'i-file',       desc: 'Exchange Online mail flow, modern auth, MailTips, transport rules.' },
    'Intune':         { icon: 'i-settings', desc: 'Intune device management, compliance and connector health.' },
    'Defender':       { icon: 'i-shield', desc: 'Microsoft Defender for Endpoint policy baselines.' },
    'AIAgent':        { icon: 'i-sparkle', desc: 'Copilot Studio AI agent security and access control.' },
    'Governance':     { icon: 'i-eye', desc: 'Entitlement management, access packages, catalogs.' },
    'Azure':          { icon: 'i-cloud',       desc: 'Azure subscription, management groups, RBAC.' },
    'Group':          { icon: 'i-user', desc: 'Microsoft 365 group lifecycle and creation restrictions.' },
    'Purview':        { icon: 'i-file', desc: 'Microsoft Purview information protection, such as published sensitivity labels.' },
    'XSPM':           { icon: 'i-search', desc: 'Exposure Security Posture Management - device exposure surface.' },
    'Backup':         { icon: 'i-folder', desc: 'Recovery Services Vault and backup hardening.' },
  };

  async function init() {
    _all = await Registry.getAll();
    await renderCategoryTiles();
    renderFullCatalog();
    populateCategoryDropdown();
    populateCustomCategoryDropdown();

    document.getElementById('catalogFilterCategory').addEventListener('change', renderRunCatalog);
    document.getElementById('catalogFilterSeverity').addEventListener('change', renderRunCatalog);
    document.getElementById('catalogFilterStatus').addEventListener('change', renderRunCatalog);
    document.getElementById('selectAllBtn').addEventListener('click', () => { for (const t of visibleRun()) if (t.implemented) _selected.add(t.id); renderRunCatalog(); });
    document.getElementById('selectNoneBtn').addEventListener('click', () => { _selected.clear(); renderRunCatalog(); });
    document.getElementById('restoreLastBtn').addEventListener('click', () => {
      const last = _loadLastSelection();
      if (!last || !last.length) { Toolkit.toast('No previous custom selection found.', 'info'); return; }
      _selected = new Set(last);
      renderRunCatalog();
    });
    document.getElementById('backToCategoriesBtn').addEventListener('click', showCategoryView);

    document.getElementById('catalogFullSearch').addEventListener('input', renderFullCatalog);
    document.getElementById('catalogFullCategory').addEventListener('change', renderFullCatalog);
    document.getElementById('catalogFullStatus').addEventListener('change', renderFullCatalog);
  }

  async function renderCategoryTiles() {
    const cats = await Registry.runCategories();
    const grid = document.getElementById('categoryTiles');
    grid.innerHTML = '';

    const allImpl = _all.filter(t => t.implemented).length;
    grid.appendChild(makeTile({
      icon: 'i-play', name: 'Run everything',
      desc: "Run every check that's implemented in the browser. Takes a couple of minutes.",
      count: allImpl,
      onClick: () => window.App.runTests(_all.filter(t => t.implemented), 'Everything'),
    }));

    for (const cat of cats) {
      const meta = CATEGORY_META[cat.name] || { icon: 'i-settings', desc: `${cat.count} check(s) in this category.` };
      grid.appendChild(makeTile({
        icon: meta.icon, name: cat.name, desc: meta.desc, count: cat.count,
        onClick: () => window.App.runTests(cat.tests, cat.name),
      }));
    }

    grid.appendChild(makeTile({
      icon: 'i-filter', name: 'Custom selection', custom: true,
      desc: 'Hand-pick tests from the full catalog.',
      countLabel: `${_all.filter(t=>t.implemented).length} available`,
      onClick: showCustomView,
    }));
  }

  function makeTile({ icon, name, desc, count, countLabel, onClick, custom }) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'category-tile' + (custom ? ' custom' : '');
    btn.innerHTML = `
      <span class="ct-icon">${Toolkit.iconHtml(icon)}</span>
      <h3>${escape(name)}</h3>
      <p>${escape(desc)}</p>
      <span class="ct-meta">
        <span>${escape(countLabel || (count + ' check' + (count === 1 ? '' : 's')))}</span>
        ${Toolkit.iconHtml('i-chevron-right', 'sm')}
      </span>`;
    btn.addEventListener('click', onClick);
    return btn;
  }

  function showCustomView() {
    document.getElementById('categoryPickerCard').hidden = true;
    document.getElementById('customPickerCard').hidden = false;
    // Default to the selection from the user's last custom run, so re-running is one click.
    // First time visiting (no history yet), pre-select all implemented Maester-native tests
    // as a sensible starter set.
    if (_selected.size === 0) {
      const last = _loadLastSelection();
      if (last && last.length) {
        const valid = new Set(_all.filter(t => t.implemented).map(t => t.id));
        _selected = new Set(last.filter(id => valid.has(id)));
      }
      if (_selected.size === 0) {
        for (const t of _all) if (t.implemented && t.category === 'Maester') _selected.add(t.id);
      }
    }
    renderRunCatalog();
  }
  function showCategoryView() {
    document.getElementById('customPickerCard').hidden = true;
    document.getElementById('categoryPickerCard').hidden = false;
  }

  function populateCustomCategoryDropdown() {
    // Use the runCategory grouping that the tile view also uses, so the dropdown
    // matches Maester's upstream taxonomy (CA, App, Privileged, ...).
    const cats = Array.from(new Set(_all.filter(t => t.implemented).map(t => t.runCategory || t.category))).sort();
    const sel = document.getElementById('catalogFilterCategory');
    sel.replaceChildren(new Option('All categories', ''), ...cats.map(c => new Option(c, c)));
  }

  function populateCategoryDropdown() {
    const cats = Array.from(new Set(_all.map(t => t.category))).sort();
    const sel = document.getElementById('catalogFullCategory');
    sel.replaceChildren(new Option('All categories', ''), ...cats.map(c => new Option(c, c)));
  }

  function visibleRun() {
    const cat = document.getElementById('catalogFilterCategory').value;
    const sev = document.getElementById('catalogFilterSeverity').value;
    const status = document.getElementById('catalogFilterStatus').value || 'implemented';
    return _all.filter(t => {
      if (status === 'implemented' && !t.implemented) return false;
      if (cat && (t.runCategory || t.category) !== cat) return false;
      if (sev && t.severity !== sev) return false;
      return true;
    });
  }

  function renderRunCatalog() {
    const list = document.getElementById('catalogList');
    list.innerHTML = '';
    const items = visibleRun();
    for (const t of items) {
      const div = document.createElement('label');
      div.className = 'catalog-item' + (t.implemented ? '' : ' disabled');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = _selected.has(t.id);
      cb.disabled = !t.implemented;
      cb.addEventListener('change', () => { if (cb.checked) _selected.add(t.id); else _selected.delete(t.id); updateStats(); });
      const body = document.createElement('div');
      body.className = 'body';
      body.innerHTML = `<div class="ci-title">${escape(t.id)} - ${escape(t.title)}</div>
        <div class="ci-meta">
          ${SMUi.severityBadge(t.severity)}
          ${SMUi.tagBadge(t.runCategory || t.category)}
          ${t.implemented ? '' : SMUi.badge('neutral', 'Not implemented')}
        </div>`;
      div.appendChild(cb);
      div.appendChild(body);
      list.appendChild(div);
    }
    document.getElementById('catalogStats').textContent = `${items.length} test(s) shown of ${_all.length} in catalog`;
    updateStats();
  }

  function updateStats() {
    document.getElementById('selectionStats').textContent = `${_selected.size} test(s) selected`;
    document.getElementById('runBtn').disabled = _selected.size === 0;
  }

  function renderFullCatalog() {
    const q = (document.getElementById('catalogFullSearch').value || '').toLowerCase();
    const cat = document.getElementById('catalogFullCategory').value;
    const stat = document.getElementById('catalogFullStatus').value;
    const items = _all.filter(t => {
      if (cat && t.category !== cat) return false;
      if (stat === 'implemented' && !t.implemented) return false;
      if (stat === 'stub' && t.implemented) return false;
      if (q) {
        const blob = `${t.id} ${t.title} ${t.severity} ${t.category}`.toLowerCase();
        if (!blob.includes(q)) return false;
      }
      return true;
    });
    const list = document.getElementById('catalogFullList');
    list.innerHTML = '';
    for (const t of items) {
      const div = document.createElement('div');
      div.className = 'catalog-item';
      div.innerHTML = `<div class="body">
        <div class="ci-title">${escape(t.id)} - ${escape(t.title)}</div>
        <div class="ci-meta">
          ${SMUi.severityBadge(t.severity)}
          ${SMUi.tagBadge(t.runCategory || t.category)}
          ${t.implemented ? SMUi.badge('ok', 'Implemented') : SMUi.badge('neutral', 'Not implemented')}
          ${t.docUrl ? `<a class="ci-doc" href="${escape(Toolkit.safeUrl(t.docUrl))}" target="_blank" rel="noopener noreferrer">Docs ${Toolkit.iconHtml('i-external', 'sm')}</a>` : ''}
        </div></div>`;
      list.appendChild(div);
    }
  }

  const escape = (s) => Toolkit.esc(s);

  function getSelectedTests() {
    return _all.filter(t => _selected.has(t.id) && t.implemented);
  }

  // Persist the current selection so the user can re-run the same set with one click.
  function recordCustomRun() {
    _saveLastSelection(Array.from(_selected));
  }

  window.UICatalog = { init, getSelectedTests, showCustomView, showCategoryView, recordCustomRun };
})();
