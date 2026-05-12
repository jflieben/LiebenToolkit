(() => {
  const state = {
    tenantPolicies: [],
    importRows: [],
    resolverMap: {},
    needsResolver: [],
    writeConsentReady: false,
    importSource: null,
  };

  function q(id) { return document.getElementById(id); }

  function esc(v) {
    return String(v == null ? '' : v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function formatDate(v) {
    if (!v) return '';
    const d = new Date(v);
    if (Number.isNaN(d.getTime())) return '';
    return d.toLocaleString();
  }

  function setConnectedUi(account) {
    const connected = !!account;
    q('connectCard').classList.toggle('hidden', connected);
    q('connectedCard').classList.toggle('hidden', !connected);
    q('userBox').classList.toggle('hidden', !connected);
    q('refreshPoliciesBtn').disabled = !connected;
    q('exportAllBtn').disabled = !connected;
    q('loadStarterBtn').disabled = !connected;
    if (account) q('userName').textContent = account.username || account.name || 'Signed in';
  }

  function setSummary(el, items) {
    el.classList.remove('hidden');
    el.innerHTML = items.map((x) => `<span>${x}</span>`).join('');
  }

  function switchTab(name) {
    document.querySelectorAll('.tab').forEach((t) => t.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach((p) => p.classList.remove('active'));
    const tab = document.querySelector(`.tab[data-tab="${name}"]`);
    if (tab) tab.classList.add('active');
    const panel = q(`tab-${name}`);
    if (panel) panel.classList.add('active');
  }

  function savePrefs() {
    const prefs = {
      prefix: q('namePrefix').value,
      suffix: q('nameSuffix').value,
      overwriteExisting: q('overwriteExisting').checked,
      theme: document.documentElement.getAttribute('data-theme') || '',
    };
    localStorage.setItem('ca-tool-prefs', JSON.stringify(prefs));
  }

  function loadPrefs() {
    try {
      const prefs = JSON.parse(localStorage.getItem('ca-tool-prefs') || '{}');
      q('namePrefix').value = prefs.prefix || '';
      q('nameSuffix').value = prefs.suffix || '';
      q('overwriteExisting').checked = !!prefs.overwriteExisting;
      if (prefs.theme === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
    } catch {}
  }

  async function readTextFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ''));
      reader.onerror = () => reject(new Error('Failed to read file.'));
      reader.readAsText(file, 'utf-8');
    });
  }

  function downloadJson(filename, payload) {
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  function toImportRows(policies, sourceLabel) {
    return (policies || []).map((policy, idx) => ({
      rowId: `${sourceLabel}-${idx}-${Math.random().toString(36).slice(2, 8)}`,
      sourcePolicy: policy,
      selected: true,
      sourceDisplayName: policy.displayName || '(no name)',
      finalDisplayName: policy.displayName || '(no name)',
      targetState: policy.state || 'enabledForReportingButNotEnforced',
      existingMatch: null,
    }));
  }

  function applyNamingAndMatching(rows) {
    const prefix = q('namePrefix').value || '';
    const suffix = q('nameSuffix').value || '';
    const byName = new Map((state.tenantPolicies || []).map((p) => [(p.displayName || '').trim().toLowerCase(), p]));

    for (const row of rows) {
      row.finalDisplayName = `${prefix}${row.sourceDisplayName}${suffix}`;
      row.existingMatch = byName.get(row.finalDisplayName.trim().toLowerCase()) || null;
    }
  }

  function renderImportPreview(targetId, rows) {
    const root = q(targetId);
    if (!rows.length) {
      root.innerHTML = '<p class="muted">No policies loaded.</p>';
      root.classList.remove('hidden');
      return;
    }

    const html = [];
    html.push('<div class="table-wrap"><table class="data-table"><thead><tr>');
    html.push('<th>Import</th><th>Source name</th><th>Final name</th><th>State</th><th>Exists</th>');
    html.push('</tr></thead><tbody>');

    for (const row of rows) {
      const exists = row.existingMatch ? '<span class="pill warn">Yes</span>' : '<span class="pill ok">No</span>';
      html.push('<tr>');
      html.push(`<td><input type="checkbox" data-role="pick" data-row="${esc(row.rowId)}" ${row.selected ? 'checked' : ''}></td>`);
      html.push(`<td>${esc(row.sourceDisplayName)}</td>`);
      html.push(`<td>${esc(row.finalDisplayName)}</td>`);
      html.push('<td>');
      html.push(`<select data-role="state" data-row="${esc(row.rowId)}">`);
      html.push(`<option value="disabled" ${row.targetState === 'disabled' ? 'selected' : ''}>disabled</option>`);
      html.push(`<option value="enabledForReportingButNotEnforced" ${row.targetState === 'enabledForReportingButNotEnforced' ? 'selected' : ''}>report-only</option>`);
      html.push(`<option value="enabled" ${row.targetState === 'enabled' ? 'selected' : ''}>active</option>`);
      html.push('</select>');
      html.push('</td>');
      html.push(`<td>${exists}</td>`);
      html.push('</tr>');
    }
    html.push('</tbody></table></div>');

    root.innerHTML = html.join('');
    root.classList.remove('hidden');

    root.querySelectorAll('input[data-role="pick"]').forEach((el) => {
      el.addEventListener('change', () => {
        const row = rows.find((r) => r.rowId === el.dataset.row);
        if (row) row.selected = el.checked;
        updateApplySection();
      });
    });
    root.querySelectorAll('select[data-role="state"]').forEach((el) => {
      el.addEventListener('change', () => {
        const row = rows.find((r) => r.rowId === el.dataset.row);
        if (row) row.targetState = el.value;
      });
    });
  }

  function allSelectedRows() {
    return state.importRows.filter((r) => r.selected);
  }

  function renderActivePreview() {
    const importWrap = q('importPreviewWrap');
    const starterWrap = q('starterPreviewWrap');

    if (state.importSource === 'starter') {
      renderImportPreview('starterPreviewWrap', state.importRows);
      importWrap.classList.add('hidden');
      importWrap.innerHTML = '';
      return;
    }

    if (state.importSource === 'import') {
      renderImportPreview('importPreviewWrap', state.importRows);
      starterWrap.classList.add('hidden');
      starterWrap.innerHTML = '';
      return;
    }

    importWrap.classList.add('hidden');
    starterWrap.classList.add('hidden');
  }

  function collectUnknownPrincipalIds(rows) {
    const allIds = new Set();
    for (const row of rows) {
      for (const id of PolicyUtils.extractPrincipalIds(row.sourcePolicy)) {
        allIds.add(id);
      }
    }
    return Array.from(allIds);
  }

  async function buildResolver(rows) {
    const ids = collectUnknownPrincipalIds(rows);
    if (!ids.length) {
      state.needsResolver = [];
      renderResolver();
      return;
    }
    const existing = await Graph.getDirectoryObjectsByIds(ids);
    const existingSet = new Set(existing.map((x) => String(x.id || '').toLowerCase()));
    state.needsResolver = ids.filter((id) => !existingSet.has(id));
    renderResolver();
  }

  function renderResolver() {
    const card = q('resolverCard');
    const root = q('resolverWrap');
    if (!state.needsResolver.length) {
      card.classList.add('hidden');
      root.innerHTML = '<p class="muted">No missing user/group references detected.</p>';
      updateApplySection();
      return;
    }

    card.classList.remove('hidden');
    const rows = state.needsResolver
      .map((id) => {
        const mapped = state.resolverMap[id] || '';
        const valid = mapped && PolicyUtils.isGuid(mapped);
        return `<tr>
          <td>${esc(id)}</td>
          <td><input type="text" data-role="resolve" data-src="${esc(id)}" value="${esc(mapped)}" placeholder="Replacement object ID" /></td>
          <td>${valid ? '<span class="pill ok">Ready</span>' : '<span class="pill err">Missing</span>'}</td>
          <td><button class="ghost-btn" data-role="clear-map" data-src="${esc(id)}" type="button">Clear</button></td>
        </tr>`;
      })
      .join('');

    root.innerHTML = `<div class="table-wrap"><table class="data-table"><thead><tr><th>Missing object ID</th><th>Replacement object ID</th><th>Status</th><th></th></tr></thead><tbody>${rows}</tbody></table></div>`;

    root.querySelectorAll('input[data-role="resolve"]').forEach((el) => {
      el.addEventListener('input', () => {
        const src = String(el.dataset.src || '').toLowerCase();
        const val = String(el.value || '').trim();
        if (val) state.resolverMap[src] = val;
        else delete state.resolverMap[src];
        renderResolver();
      });
    });
    root.querySelectorAll('button[data-role="clear-map"]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const src = String(btn.dataset.src || '').toLowerCase();
        delete state.resolverMap[src];
        renderResolver();
      });
    });
    updateApplySection();
  }

  function unresolvedCount() {
    return state.needsResolver.filter((id) => {
      const mapped = state.resolverMap[id];
      return !mapped || !PolicyUtils.isGuid(mapped);
    }).length;
  }

  function updateApplySection() {
    const selected = allSelectedRows();
    const card = q('applyCard');
    const hint = q('applyHint');
    const btn = q('applyImportBtn');

    if (!selected.length) {
      card.classList.add('hidden');
      return;
    }

    card.classList.remove('hidden');
    const unresolved = unresolvedCount();
    const overwrite = q('overwriteExisting').checked;
    hint.textContent = `${selected.length} selected policy/policies. Overwrite mode: ${overwrite ? 'on' : 'off'}. Unresolved references: ${unresolved}.`;
    btn.disabled = unresolved > 0;
  }

  async function refreshTenantPolicies() {
    Log.info('Loading tenant Conditional Access policies...');
    const list = await Graph.listConditionalAccessPolicies();
    state.tenantPolicies = list.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));

    const wrap = q('exportTableWrap');
    const rows = state.tenantPolicies
      .map((p) => `<tr><td>${esc(p.displayName || '')}</td><td>${esc(p.state || '')}</td><td>${formatDate(p.createdDateTime)}</td><td>${formatDate(p.modifiedDateTime)}</td></tr>`)
      .join('');
    wrap.innerHTML = `<table class="data-table"><thead><tr><th>Name</th><th>State</th><th>Created</th><th>Modified</th></tr></thead><tbody>${rows}</tbody></table>`;
    wrap.classList.toggle('hidden', state.tenantPolicies.length === 0);

    setSummary(q('exportSummary'), [
      `${state.tenantPolicies.length} policy/policies loaded`,
      `Tenant read completed`,
    ]);
    q('copyExportSummaryBtn').disabled = state.tenantPolicies.length === 0;
    Log.info(`Loaded ${state.tenantPolicies.length} policies.`);
    if (state.importRows.length) {
      applyNamingAndMatching(state.importRows);
      renderActivePreview();
      updateApplySection();
    }
  }

  function makeExportPackage(policies) {
    return {
      tool: 'ConditionalAccess',
      schemaVersion: 1,
      exportedAtUtc: new Date().toISOString(),
      policyCount: policies.length,
      metadata: {
        user: (Auth.getAccount() && (Auth.getAccount().username || Auth.getAccount().name)) || '',
        source: 'Graph /identity/conditionalAccess/policies',
      },
      policies,
    };
  }

  async function loadImportFromFile() {
    const file = q('importFile').files && q('importFile').files[0];
    if (!file) throw new Error('Pick an import file first.');
    const text = await readTextFile(file);
    const payload = PolicyUtils.parseImportPayload(text);
    state.importRows = toImportRows(payload.policies || [], 'import');
    state.importSource = 'import';
    state.resolverMap = {};
    applyNamingAndMatching(state.importRows);
    renderActivePreview();
    await buildResolver(state.importRows);
    updateApplySection();
    Log.info(`Loaded ${state.importRows.length} import policy rows.`);
  }

  async function loadStarterSet() {
    const starter = PolicyUtils.buildStarterPolicies();
    state.importRows = toImportRows(starter, 'starter');
    state.importSource = 'starter';
    state.resolverMap = {};
    applyNamingAndMatching(state.importRows);
    renderActivePreview();
    await buildResolver(state.importRows);
    updateApplySection();
    Log.info(`Starter set loaded (${starter.length} policies).`);
  }

  async function ensureWriteConsent() {
    if (state.writeConsentReady) return;
    Log.info('Requesting write consent...');
    const acct = await Auth.signIn(true);
    if (acct) {
      state.writeConsentReady = true;
      setConnectedUi(acct);
    } else {
      throw new Error('Write consent flow started via redirect. Continue after sign-in.');
    }
  }

  function buildMappingObject() {
    const out = {};
    for (const [k, v] of Object.entries(state.resolverMap)) {
      if (PolicyUtils.isGuid(v)) out[String(k).toLowerCase()] = String(v).trim();
    }
    return out;
  }

  async function applyImport() {
    const selected = allSelectedRows();
    if (!selected.length) return;
    if (unresolvedCount() > 0) {
      alert('Resolve all missing references first.');
      return;
    }

    const overwrite = q('overwriteExisting').checked;
    if (!confirm(`Apply ${selected.length} policy/policies? Overwrite existing: ${overwrite ? 'yes' : 'no'}.`)) return;

    await ensureWriteConsent();

    const results = [];
    const byName = new Map((state.tenantPolicies || []).map((p) => [(p.displayName || '').trim().toLowerCase(), p]));
    const mapping = buildMappingObject();
    q('applyImportBtn').disabled = true;

    for (const row of selected) {
      const existing = byName.get(row.finalDisplayName.trim().toLowerCase()) || null;
      if (existing && !overwrite) {
        results.push({ ok: false, action: 'skip', name: row.finalDisplayName, message: 'Exists and overwrite is disabled.' });
        continue;
      }

      try {
        let policy = PolicyUtils.applyPrincipalMapping(row.sourcePolicy, mapping);
        policy = PolicyUtils.normalizePolicyForCreateOrUpdate(policy, row.finalDisplayName, row.targetState);
        if (existing) {
          await Graph.updateConditionalAccessPolicy(existing.id, policy);
          results.push({ ok: true, action: 'update', name: row.finalDisplayName, message: 'Updated existing policy.' });
        } else {
          await Graph.createConditionalAccessPolicy(policy);
          results.push({ ok: true, action: 'create', name: row.finalDisplayName, message: 'Created new policy.' });
        }
      } catch (e) {
        results.push({ ok: false, action: existing ? 'update' : 'create', name: row.finalDisplayName, message: e.message });
      }
    }

    const ok = results.filter((r) => r.ok).length;
    const fail = results.length - ok;
    const html = [
      `<p><strong>${ok}</strong> succeeded, <strong>${fail}</strong> failed.</p>`,
      '<ul class="result-list">',
      ...results.map((r) => `<li><span class="pill ${r.ok ? 'ok' : 'err'}">${esc(r.action)}</span> <strong>${esc(r.name)}</strong><br>${esc(r.message)}</li>`),
      '</ul>',
    ].join('');
    q('applyResults').innerHTML = html;
    q('applyImportBtn').disabled = false;

    await refreshTenantPolicies();
    Log.info(`Import apply finished: ${ok} ok, ${fail} failed.`);
  }

  async function runDiff() {
    const fOld = q('diffFileOld').files && q('diffFileOld').files[0];
    const fNew = q('diffFileNew').files && q('diffFileNew').files[0];
    if (!fOld || !fNew) return;

    const [oldText, newText] = await Promise.all([readTextFile(fOld), readTextFile(fNew)]);
    const oldPayload = PolicyUtils.parseImportPayload(oldText);
    const newPayload = PolicyUtils.parseImportPayload(newText);
    const diff = PolicyUtils.comparePolicySets(oldPayload.policies || [], newPayload.policies || []);

    setSummary(q('diffSummary'), [
      `Added: ${diff.added.length}`,
      `Removed: ${diff.removed.length}`,
      `Changed: ${diff.changed.length}`,
      `Unchanged: ${diff.unchanged.length}`,
    ]);

    const lines = [];
    if (diff.added.length) {
      lines.push('<h3>Added</h3><ul class="result-list">');
      lines.push(...diff.added.map((n) => `<li><span class="pill ok">added</span> ${esc(n)}</li>`));
      lines.push('</ul>');
    }
    if (diff.removed.length) {
      lines.push('<h3>Removed</h3><ul class="result-list">');
      lines.push(...diff.removed.map((n) => `<li><span class="pill err">removed</span> ${esc(n)}</li>`));
      lines.push('</ul>');
    }
    if (diff.changed.length) {
      lines.push('<h3>Changed</h3><ul class="result-list">');
      lines.push(...diff.changed.map((c) => `<li><span class="pill warn">changed</span> <strong>${esc(c.name)}</strong><br>${esc(c.paths.slice(0, 20).join(', '))}${c.paths.length > 20 ? ' ...' : ''}</li>`));
      lines.push('</ul>');
    }
    if (!lines.length) lines.push('<p class="muted">No differences found.</p>');

    q('diffWrap').innerHTML = lines.join('');
    Log.info('Diff finished.', diff);
  }

  function wireTabsAndTheme() {
    document.querySelectorAll('.tab').forEach((tab) => {
      tab.addEventListener('click', () => switchTab(tab.dataset.tab));
    });

    q('darkToggle').addEventListener('click', () => {
      const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
      if (isDark) document.documentElement.removeAttribute('data-theme');
      else document.documentElement.setAttribute('data-theme', 'dark');
      savePrefs();
    });
  }

  function wireUi() {
    q('verboseLog').addEventListener('change', (e) => Log.setVerbose(e.target.checked));
    q('clearLogBtn').addEventListener('click', () => Log.clear());

    q('signInBtn').addEventListener('click', async () => {
      try {
        const acct = await Auth.signIn(false);
        if (acct) {
          setConnectedUi(acct);
          await refreshTenantPolicies();
        }
      } catch (e) {
        Log.err('Sign-in failed:', e.message);
        alert('Sign-in failed: ' + e.message);
      }
    });

    q('signOutBtn').addEventListener('click', async () => {
      await Auth.signOut();
      state.writeConsentReady = false;
      setConnectedUi(null);
    });

    q('refreshPoliciesBtn').addEventListener('click', async () => {
      try { await refreshTenantPolicies(); } catch (e) { Log.err(e.message); alert(e.message); }
    });

    q('exportAllBtn').addEventListener('click', async () => {
      try {
        if (!state.tenantPolicies.length) await refreshTenantPolicies();
        const payload = makeExportPackage(state.tenantPolicies);
        const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
        downloadJson(`conditional-access-export-${stamp}.json`, payload);
        setSummary(q('exportSummary'), [
          `${state.tenantPolicies.length} policy/policies exported`,
          `Created ${stamp}`,
        ]);
      } catch (e) {
        Log.err('Export failed:', e.message);
        alert('Export failed: ' + e.message);
      }
    });

    q('copyExportSummaryBtn').addEventListener('click', async () => {
      try {
        const lines = state.tenantPolicies.map((p) => `${p.displayName || ''}\t${p.state || ''}`).join('\n');
        await navigator.clipboard.writeText(lines);
        Log.info('Export summary copied to clipboard.');
      } catch (e) {
        Log.warn('Clipboard copy failed:', e.message);
      }
    });

    q('importFile').addEventListener('change', () => {
      q('loadImportBtn').disabled = !(q('importFile').files && q('importFile').files[0]);
    });

    q('loadImportBtn').addEventListener('click', async () => {
      try {
        await loadImportFromFile();
      } catch (e) {
        Log.err('Load import failed:', e.message);
        alert('Failed to load import file: ' + e.message);
      }
    });

    q('loadStarterBtn').addEventListener('click', async () => {
      try {
        await loadStarterSet();
      } catch (e) {
        Log.err('Load starter set failed:', e.message);
        alert('Failed to load starter set: ' + e.message);
      }
    });

    ['namePrefix', 'nameSuffix', 'overwriteExisting'].forEach((id) => {
      q(id).addEventListener('change', () => {
        savePrefs();
        if (state.importRows.length) {
          applyNamingAndMatching(state.importRows);
          renderActivePreview();
          updateApplySection();
        }
      });
    });

    q('applyImportBtn').addEventListener('click', async () => {
      try {
        await applyImport();
      } catch (e) {
        Log.err('Apply import failed:', e.message);
        alert('Apply import failed: ' + e.message);
      }
    });

    const diffReady = () => {
      const ready = (q('diffFileOld').files && q('diffFileOld').files[0]) && (q('diffFileNew').files && q('diffFileNew').files[0]);
      q('runDiffBtn').disabled = !ready;
    };
    q('diffFileOld').addEventListener('change', diffReady);
    q('diffFileNew').addEventListener('change', diffReady);
    q('runDiffBtn').addEventListener('click', async () => {
      try {
        await runDiff();
      } catch (e) {
        Log.err('Diff failed:', e.message);
        alert('Diff failed: ' + e.message);
      }
    });
  }

  async function init() {
    wireTabsAndTheme();
    wireUi();
    loadPrefs();

    try {
      const acct = await Auth.init();
      setConnectedUi(acct);
      if (acct) await refreshTenantPolicies();
    } catch (e) {
      Log.err('Initialization failed:', e.message);
      alert('Initialization failed: ' + e.message);
    }
  }

  init();
})();