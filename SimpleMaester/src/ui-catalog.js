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
    'CA':             { icon: '\uD83D\uDEA6', desc: 'Conditional Access. Block legacy auth, require MFA, exclude break-glass, scope by risk and device.' },
    'App':            { icon: '\uD83E\uDDE9', desc: 'Applications and service principals. Owners, secrets, high-risk permissions, broad assignments.' },
    'Privileged':     { icon: '\uD83D\uDC51', desc: 'Privileged Identity Management and admin role hygiene. PIM eligibility, GA count, cloud-only admins.' },
    'Authentication': { icon: '\uD83D\uDD11', desc: 'Authentication methods. Phishing-resistant MFA, FIDO2, MS Authenticator, weak methods like SMS/voice.' },
    'Entra':          { icon: '\uD83C\uDFE2', desc: 'General Entra ID hygiene. Stale guests, administrative units, tenant settings, group creation.' },
    'General':        { icon: '\uD83D\uDEE1', desc: 'EIDSCA - Entra ID Security Config Analyzer. 44 baseline configuration controls.' },
    'CIS':            { icon: '\uD83D\uDEE1', desc: 'CIS Microsoft 365 Foundations Benchmark v6.0.1 - 38 controls across identity, mail, devices and apps.' },
    // The following Maester tags exist upstream but we have no implemented tests for them yet.
    // They will only show up if/when tests are added.
    'Teams':          { icon: '\uD83D\uDC65', desc: 'Microsoft Teams meeting and collaboration controls.' },
    'Exchange':       { icon: '\u2709',       desc: 'Exchange Online mail flow, modern auth, MailTips, transport rules.' },
    'Intune':         { icon: '\uD83D\uDCF1', desc: 'Intune device management, compliance and connector health.' },
    'Defender':       { icon: '\uD83D\uDD12', desc: 'Microsoft Defender for Endpoint policy baselines.' },
    'AIAgent':        { icon: '\uD83E\uDD16', desc: 'Copilot Studio AI agent security and access control.' },
    'Governance':     { icon: '\uD83D\uDCDC', desc: 'Entitlement management, access packages, catalogs.' },
    'Azure':          { icon: '\u2601',       desc: 'Azure subscription, management groups, RBAC.' },
    'Group':          { icon: '\uD83D\uDC65', desc: 'Microsoft 365 group lifecycle and creation restrictions.' },
    'XSPM':           { icon: '\uD83D\uDD0D', desc: 'Exposure Security Posture Management - device exposure surface.' },
    'Backup':         { icon: '\uD83D\uDCBE', desc: 'Recovery Services Vault and backup hardening.' },
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
      if (!last || !last.length) { window.App?.toast?.('No previous custom selection found.'); return; }
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
      icon: '\uD83C\uDFAF', name: 'Run everything',
      desc: "Run every check that's implemented in the browser. Takes a couple of minutes.",
      count: allImpl,
      onClick: () => window.App.runTests(_all.filter(t => t.implemented), 'Everything'),
    }));

    for (const cat of cats) {
      const meta = CATEGORY_META[cat.name] || { icon: '\u2699', desc: `${cat.count} check(s) in this category.` };
      grid.appendChild(makeTile({
        icon: meta.icon, name: cat.name, desc: meta.desc, count: cat.count,
        onClick: () => window.App.runTests(cat.tests, cat.name),
      }));
    }

    grid.appendChild(makeTile({
      icon: '\uD83C\uDFA8', name: 'Custom selection', custom: true,
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
      <div class="ct-icon">${icon}</div>
      <h3>${escape(name)}</h3>
      <p>${escape(desc)}</p>
      <div class="ct-meta">
        <span>${countLabel || (count + ' check' + (count === 1 ? '' : 's'))}</span>
        <span>&rsaquo;</span>
      </div>`;
    btn.addEventListener('click', onClick);
    return btn;
  }

  function showCustomView() {
    document.getElementById('categoryPickerCard').classList.add('hidden');
    document.getElementById('customPickerCard').classList.remove('hidden');
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
    document.getElementById('customPickerCard').classList.add('hidden');
    document.getElementById('categoryPickerCard').classList.remove('hidden');
  }

  function populateCustomCategoryDropdown() {
    // Use the runCategory grouping that the tile view also uses, so the dropdown
    // matches Maester's upstream taxonomy (CA, App, Privileged, ...).
    const cats = Array.from(new Set(_all.filter(t => t.implemented).map(t => t.runCategory || t.category))).sort();
    const sel = document.getElementById('catalogFilterCategory');
    sel.innerHTML = '<option value="">All categories</option>' + cats.map(c => `<option value="${escape(c)}">${escape(c)}</option>`).join('');
  }

  function populateCategoryDropdown() {
    const cats = Array.from(new Set(_all.map(t => t.category))).sort();
    const sel = document.getElementById('catalogFullCategory');
    sel.innerHTML = '<option value="">All categories</option>' + cats.map(c => `<option value="${c}">${c}</option>`).join('');
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
      const div = document.createElement('div');
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
          <span class="badge ${(t.severity||'info').toLowerCase()}">${t.severity||'-'}</span>
          <span class="badge tag">${escape(t.runCategory || t.category)}</span>
          ${t.implemented ? '' : '<span class="badge stub">Not implemented</span>'}
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
          <span class="badge ${(t.severity||'info').toLowerCase()}">${t.severity||'-'}</span>
          <span class="badge tag">${escape(t.runCategory || t.category)}</span>
          ${t.implemented ? '<span class="badge passed">Implemented</span>' : '<span class="badge stub">Not implemented</span>'}
          ${t.docUrl ? `<a href="${t.docUrl}" target="_blank" rel="noopener" style="margin-left:auto;font-size:11px">docs &rarr;</a>` : ''}
        </div></div>`;
      list.appendChild(div);
    }
  }

  function escape(s) { return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

  function getSelectedTests() {
    return _all.filter(t => _selected.has(t.id) && t.implemented);
  }

  // Persist the current selection so the user can re-run the same set with one click.
  function recordCustomRun() {
    _saveLastSelection(Array.from(_selected));
  }

  window.UICatalog = { init, getSelectedTests, showCustomView, showCategoryView, recordCustomRun };
})();
