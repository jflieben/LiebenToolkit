// Main orchestration & UI wiring
(() => {
  const state = {
    sites: [],         // discovered sites
    selected: new Set(), // selected site IDs
    results: [],       // analysis results
    cancelToken: null,
    sortSites:   { key: null, dir: 1 },
    sortResults: { key: null, dir: 1 },
    // Excel-style column filters: per-column, set of values to HIDE.
    // Empty set = show all (default).
    colFilters: { type: new Set(), advice: new Set() },
  };

  // ---------- Generic table sort ----------
  // Reads `data-sort` (key) and optional `data-sort-type` ('date'|'number'|'string')
  // from a <th>, then re-renders the bound table.
  function wireSortHeaders(tableSelector, sortState, rerender) {
    const ths = document.querySelectorAll(`${tableSelector} thead th[data-sort]`);
    ths.forEach(th => {
      th.classList.add('sortable-col');
      th.addEventListener('click', () => {
        const key = th.dataset.sort;
        if (sortState.key === key) sortState.dir = -sortState.dir;
        else { sortState.key = key; sortState.dir = 1; sortState.type = th.dataset.sortType || 'string'; }
        sortState.type = th.dataset.sortType || 'string';
        ths.forEach(x => { x.classList.remove('sort-asc', 'sort-desc'); });
        th.classList.add(sortState.dir === 1 ? 'sort-asc' : 'sort-desc');
        rerender();
      });
    });
  }
  function applySort(items, sortState, valueOf) {
    if (!sortState.key) return items;
    const dir = sortState.dir;
    const t = sortState.type || 'string';
    const arr = items.slice();
    arr.sort((a, b) => {
      let av = valueOf(a, sortState.key);
      let bv = valueOf(b, sortState.key);
      if (av == null && bv == null) return 0;
      if (av == null) return 1;            // nulls last
      if (bv == null) return -1;
      if (t === 'date') {
        av = av instanceof Date ? av.getTime() : new Date(av).getTime();
        bv = bv instanceof Date ? bv.getTime() : new Date(bv).getTime();
        if (isNaN(av) && isNaN(bv)) return 0;
        if (isNaN(av)) return 1;
        if (isNaN(bv)) return -1;
      } else if (t === 'number') {
        av = Number(av); bv = Number(bv);
        if (isNaN(av) && isNaN(bv)) return 0;
        if (isNaN(av)) return 1;
        if (isNaN(bv)) return -1;
      } else {
        av = String(av).toLowerCase();
        bv = String(bv).toLowerCase();
      }
      return av < bv ? -dir : av > bv ? dir : 0;
    });
    return arr;
  }
  function populateSelectOptions(selectEl, values, currentValue) {
    const cur = currentValue ?? selectEl.value;
    // Keep the first <option> (the "All ..." entry); replace the rest
    while (selectEl.options.length > 1) selectEl.remove(1);
    [...new Set(values.filter(v => v != null && v !== ''))].sort().forEach(v => {
      const o = document.createElement('option');
      o.value = v; o.textContent = v;
      selectEl.appendChild(o);
    });
    if (cur && [...selectEl.options].some(o => o.value === cur)) selectEl.value = cur;
  }

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
  const saved = localStorage.getItem('spotrim-theme');
  if (saved === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
  document.getElementById('darkToggle').addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme') === 'dark' ? '' : 'dark';
    if (cur === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
    else document.documentElement.removeAttribute('data-theme');
    localStorage.setItem('spotrim-theme', cur);
  });

  // ---------- Verbose log ----------
  document.getElementById('verboseLog').addEventListener('change', (e) => Log.setVerbose(e.target.checked));
  document.getElementById('clearLogBtn').addEventListener('click', () => Log.clear());

  // ---------- Toast ----------
  let toastTimer = null;
  function toast(msg, ms = 4000) {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.remove('hidden');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.add('hidden'), ms);
  }

  // ---------- Auth ----------
  function setSignedInUI(account) {
    if (account) {
      document.getElementById('connectCard').classList.add('hidden');
      document.getElementById('discoveryCard').classList.remove('hidden');
      document.getElementById('userBox').classList.remove('hidden');
      document.getElementById('userName').textContent = account.username || account.name || 'Signed in';
    } else {
      document.getElementById('connectCard').classList.remove('hidden');
      document.getElementById('discoveryCard').classList.add('hidden');
      document.getElementById('userBox').classList.add('hidden');
    }
  }

  document.getElementById('signInBtn').addEventListener('click', async () => {
    try {
      const acc = await Auth.signIn();
      if (acc) {
        Log.info(`Signed in as ${acc.username}`);
        setSignedInUI(acc);
      }
    } catch (e) {
      Log.err('Sign-in failed:', e);
      toast(`Sign-in failed: ${e.message}`);
    }
  });

  document.getElementById('signOutBtn').addEventListener('click', async () => {
    await Auth.signOut();
    setSignedInUI(null);
  });

  // ---------- Discovery ----------
  document.getElementById('discoverBtn').addEventListener('click', discover);
  document.getElementById('analyzeBtn').addEventListener('click', analyze);
  document.getElementById('cancelBtn').addEventListener('click', () => {
    if (state.cancelToken) { state.cancelToken.cancel(); toast('Cancelling...'); }
  });
  document.getElementById('selectAllBtn').addEventListener('click', () => {
    state.sites.forEach(s => state.selected.add(s.id));
    renderSites();
  });
  document.getElementById('selectNoneBtn').addEventListener('click', () => {
    state.selected.clear();
    renderSites();
  });
  document.getElementById('selAll').addEventListener('change', (e) => {
    if (e.target.checked) filteredSites().forEach(s => state.selected.add(s.id));
    else                  filteredSites().forEach(s => state.selected.delete(s.id));
    renderSites();
  });
  document.getElementById('siteFilter').addEventListener('input', renderSites);
  document.getElementById('typeFilter').addEventListener('change', renderSites);
  document.getElementById('includeOneDrive').addEventListener('change', renderSites);
  document.getElementById('includeSharePoint').addEventListener('change', renderSites);
  wireSortHeaders('#siteTable', state.sortSites, renderSites);

  function classify(site) { return Exporter.classify(site); }
  function classifyType(site) { return Exporter.classifyType(site); }

  function filteredSites() {
    const q = (document.getElementById('siteFilter').value || '').toLowerCase();
    const tf = document.getElementById('typeFilter').value;
    const incOd = document.getElementById('includeOneDrive').checked;
    const incSp = document.getElementById('includeSharePoint').checked;
    return state.sites.filter(s => {
      const cat = classify(s);
      const t = classifyType(s);
      if (cat === 'OneDrive' && !incOd) return false;
      if (cat === 'SharePoint' && !incSp) return false;
      if (tf && t !== tf) return false;
      if (q && !((s.displayName || '').toLowerCase().includes(q) || (s.webUrl || '').toLowerCase().includes(q))) return false;
      return true;
    });
  }

  async function discover() {
    document.getElementById('discoverBtn').disabled = true;
    document.getElementById('analyzeBtn').disabled = true;
    state.cancelToken = new Concurrency.CancelToken();
    const incOd = document.getElementById('includeOneDrive').checked;
    const incSp = document.getElementById('includeSharePoint').checked;
    if (!incOd && !incSp) { toast('Select at least one site type.'); document.getElementById('discoverBtn').disabled = false; return; }
    showProgress('Discovering sites...', null);
    try {
      const collected = [];
      const seen = new Set(); // by normalized URL

      // 1. SharePoint sites (and any personal sites the search index decides to surface) via SP Search.
      if (incSp || incOd) {
        Log.info(`Discovering sites via SharePoint Search (include SP=${incSp}, include OD=${incOd})...`);
        try {
          const searchSites = await SharePoint.searchAllSites({
            includeOneDrive: incOd,
            includeSharePoint: incSp,
            onProgress: (n) => updateProgress(n, n, 0),
          });
          for (const s of searchSites) {
            const k = (s.webUrl || '').toLowerCase().replace(/\/+$/, '');
            if (!k || seen.has(k)) continue;
            seen.add(k);
            collected.push(s);
          }
          Log.info(`SharePoint Search returned ${searchSites.length} sites (SP=${searchSites.filter(s => !s.isPersonalSite).length}, OD=${searchSites.filter(s => s.isPersonalSite).length}).`);
        } catch (e) {
          Log.warn(`SharePoint Search discovery failed: ${e.message}`);
        }
      }

      // 2. OneDrive sites - SharePoint Search trims results to what the user can access,
      //    so it never returns *other users'* OneDrives. We use the SharePoint Admin
      //    Tenant REST API which works directly from the browser (CORS-OK) and requires
      //    the SharePoint Administrator role.
      if (incOd) {
        let odSites = [];
        Log.info('Discovering OneDrive sites via SharePoint Admin Tenant API (requires SP Admin role)...');
        try {
          odSites = await SharePoint.listOneDriveSitesViaTenantApi();
          Log.info(`SP Admin tenant API returned ${odSites.length} OneDrive sites.`);
        } catch (e) {
          Log.warn(`SP Admin tenant API failed: ${e.message}`);
          toast(`OneDrive enumeration failed (SharePoint Administrator role required): ${e.message}`, 10000);
        }
        let added = 0;
        for (const s of odSites) {
          const k = (s.webUrl || '').toLowerCase().replace(/\/+$/, '');
          if (!k || seen.has(k)) continue;
          seen.add(k);
          collected.push(s);
          added++;
        }
        Log.info(`OneDrive enumeration: ${odSites.length} returned, ${added} added.`);
      }

      state.sites = collected;
      const spCount = collected.filter(s => !s.isPersonalSite).length;
      const odCount = collected.filter(s => s.isPersonalSite).length;
      Log.info(`Total discovered: ${collected.length} sites (SharePoint=${spCount}, OneDrive=${odCount}).`);
      toast(`Discovered ${collected.length} sites (${spCount} SP, ${odCount} OD).`);
      document.getElementById('siteListWrap').classList.remove('hidden');
      document.getElementById('selectAllBtn').disabled = false;
      document.getElementById('selectNoneBtn').disabled = false;
      renderSites();
    } catch (e) {
      Log.err('Discovery failed:', e);
      toast(`Discovery failed: ${e.message}`);
    } finally {
      document.getElementById('discoverBtn').disabled = false;
      hideProgress();
    }
  }

  function renderSites() {
    const tbody = document.querySelector('#siteTable tbody');
    tbody.innerHTML = '';
    // Refresh type-filter options from the full discovered set (so the user can pick any present type)
    populateSelectOptions(document.getElementById('typeFilter'),
      state.sites.map(classifyType));
    const list = applySort(filteredSites(), state.sortSites, (s, k) => {
      switch (k) {
        case 'title':    return s.displayName || s.name || '';
        case 'url':      return s.webUrl || '';
        case 'type':     return classifyType(s);
        case 'template': return (s.root && s.root['@odata.type']) || '';
        case 'created':  return s.createdDateTime || null;
        default: return '';
      }
    });
    for (const s of list) {
      const tr = document.createElement('tr');
      const cat = classify(s);
      const t = classifyType(s);
      const checked = state.selected.has(s.id) ? 'checked' : '';
      tr.innerHTML = `
        <td><input type="checkbox" data-id="${esc(s.id)}" ${checked}></td>
        <td>${esc(s.displayName || s.name || '')}</td>
        <td class="url-cell" title="${esc(s.webUrl || '')}"><a href="${esc(s.webUrl || '')}" target="_blank">${esc(s.webUrl || '')}</a></td>
        <td><span class="badge ${cat === 'OneDrive' ? 'od' : 'spo'}">${esc(t)}</span></td>
        <td>${esc((s.root && s.root['@odata.type']) || '')}</td>
        <td>${fmtDate(s.createdDateTime)}</td>`;
      tbody.appendChild(tr);
    }
    tbody.querySelectorAll('input[type=checkbox]').forEach(cb => {
      cb.addEventListener('change', (e) => {
        const id = e.target.dataset.id;
        if (e.target.checked) state.selected.add(id);
        else state.selected.delete(id);
        updateSelCount();
      });
    });
    updateSelCount();
  }

  function updateSelCount() {
    document.getElementById('selCount').textContent = state.selected.size;
    document.getElementById('analyzeBtn').disabled = state.selected.size === 0;
  }

  // ---------- Analysis ----------
  async function analyze() {
    const selected = state.sites.filter(s => state.selected.has(s.id));
    if (!selected.length) return;
    const concurrency = parseInt(document.getElementById('concurrency').value, 10) || 8;
    const staleDays   = parseInt(document.getElementById('staleDays').value, 10) || 365;
    const forceScan   = document.getElementById('forceScan').checked;

    document.getElementById('analyzeBtn').disabled = true;
    document.getElementById('discoverBtn').disabled = true;
    state.cancelToken = new Concurrency.CancelToken();

    // Force-scan elevation: temporarily make ourselves a site collection admin
    // on sites we don't already own. Track per site whether WE added the
    // grant so we only revoke what we added (never touching pre-existing access).
    const elevated = []; // [{ site, claim }]
    if (forceScan) {
      const acct = Auth.getAccount();
      const upn = acct && (acct.username || acct.name);
      if (!upn) {
        toast('Force scan: cannot determine your UPN, aborting.');
        document.getElementById('analyzeBtn').disabled = false;
        document.getElementById('discoverBtn').disabled = false;
        return;
      }
      const claim = SharePoint.buildClaimLoginName(upn);
      Log.info(`Force scan enabled. Checking site collection admin status for ${selected.length} site(s) as ${upn}...`);
      showProgress('Force scan: checking and granting temporary access...', selected.length);
      let done = 0, granted = 0, alreadyAdmin = 0, failed = 0;
      await Concurrency.pmap(selected, async (site) => {
        if (state.cancelToken.cancelled) return;
        try {
          const status = await SharePoint.getCurrentUserAdminStatus(site.webUrl);
          if (status.isSiteAdmin) { alreadyAdmin++; return; }
          // Either no access or non-admin access: try to grant.
          await SharePoint.setSiteAdmin(site.webUrl, claim, true);
          elevated.push({ site, claim });
          granted++;
        } catch (e) {
          failed++;
          Log.warn(`Force scan grant failed for ${site.webUrl}: ${e.message}`);
        } finally {
          done++;
          updateProgress(done, selected.length, 0);
        }
      }, { concurrency, cancelToken: state.cancelToken });
      Log.info(`Force scan: ${alreadyAdmin} already admin, ${granted} granted, ${failed} failed.`);
      if (failed > 0) toast(`Force scan: failed to grant access to ${failed} site(s). Check the debug log.`);
    }

    // Per-site engagement signals (analytics, drive activity, page edits) are
    // fetched in parallel inside Analysis.analyzeSite via Graph.getSiteEngagement.
    // No tenant-wide bulk fetch - everything works directly from the browser
    // against graph.microsoft.com (CORS-friendly).

    // Step 2: per-site analysis
    showProgress(`Analyzing ${selected.length} sites...`, selected.length);
    const start = Date.now();

    const { results } = await Concurrency.pmap(selected, async (site) => {
      state.cancelToken.throwIfCancelled();
      try {
        return await Analysis.analyzeSite(site, { staleDays });
      } catch (e) {
        Log.err(`Analysis failed for ${site.webUrl}:`, e.message);
        return { site, signals: {}, lastModified: null, lastViewed: null, advice: 'Unknown', reason: `Analysis error: ${e.message}`, errors: [e.message] };
      }
    }, {
      concurrency,
      cancelToken: state.cancelToken,
      onProgress: (done, total) => {
        const elapsed = (Date.now() - start) / 1000;
        const rate = done / Math.max(elapsed, 0.1);
        const remain = total - done;
        const eta = remain > 0 && rate > 0 ? Math.ceil(remain / rate) : 0;
        updateProgress(done, total, eta);
      },
    });

    state.results = results.filter(r => r && !r.__error);
    Log.info(`Analysis complete: ${state.results.length} results.`);

    // Force-scan cleanup: revoke every grant we (and only we) added.
    if (elevated.length > 0) {
      showProgress(`Removing temporary site collection admin grants from ${elevated.length} site(s)...`, elevated.length);
      let cleaned = 0, failed = 0, doneRev = 0;
      await Concurrency.pmap(elevated, async ({ site, claim }) => {
        try {
          await SharePoint.setSiteAdmin(site.webUrl, claim, false);
          cleaned++;
        } catch (e) {
          failed++;
          Log.err(`Force scan REVOKE failed for ${site.webUrl}: ${e.message}. Remove manually if needed.`);
        } finally {
          doneRev++;
          updateProgress(doneRev, elevated.length, 0);
        }
      }, { concurrency });
      Log.info(`Force scan cleanup: ${cleaned} revoked, ${failed} failed.`);
      if (failed > 0) toast(`Force scan cleanup: ${failed} grant(s) NOT revoked. Check the debug log and remove manually.`, 12000);
    }

    hideProgress();
    document.getElementById('analyzeBtn').disabled = false;
    document.getElementById('discoverBtn').disabled = false;
    document.getElementById('exportXlsxBtn').disabled = state.results.length === 0;
    document.getElementById('exportCsvBtn').disabled  = state.results.length === 0;
    renderResults();
    // Switch to results tab
    document.querySelector('.tab[data-tab="results"]').click();
  }

  // ---------- Results ----------
  document.getElementById('resultFilter').addEventListener('input', renderResults);
  document.getElementById('clearFiltersBtn').addEventListener('click', () => {
    state.colFilters.type.clear();
    state.colFilters.advice.clear();
    renderResults();
  });
  document.getElementById('exportXlsxBtn').addEventListener('click', () => Exporter.exportXlsx(state.results));
  document.getElementById('exportCsvBtn').addEventListener('click', () => Exporter.exportCsv(state.results));
  wireSortHeaders('#resultTable', state.sortResults, renderResults);
  wireColumnFilterHeaders();
  wireDetailModal();

  function renderResults() {
    const tbody = document.querySelector('#resultTable tbody');
    tbody.innerHTML = '';
    renderResultNotices();
    refreshColumnFilterIndicators();
    const q = (document.getElementById('resultFilter').value || '').toLowerCase();
    let list = state.results.filter(r => {
      const t = classifyType(r.site);
      if (state.colFilters.type.has(t)) return false;
      if (state.colFilters.advice.has(r.advice || '')) return false;
      if (q) {
        const hay = `${r.site.displayName || ''} ${r.site.webUrl || ''}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
    list = applySort(list, state.sortResults, (r, k) => {
      switch (k) {
        case 'title':        return r.site.displayName || r.site.name || '';
        case 'url':          return r.site.webUrl || '';
        case 'type':         return classifyType(r.site);
        case 'owners':       return (r.owners && r.owners[0] ? (r.owners[0].title || r.owners[0].email) : '') || '';
        case 'lastModified': return r.lastModified || null;
        case 'lastViewed':   return r.lastViewed || null;
        case 'daysIdle':     return r.daysSinceModified == null ? null : Number(r.daysSinceModified);
        case 'pageViews':    return r.pageViews == null ? null : Number(r.pageViews);
        case 'storageMB':    return r.storageMB == null ? null : Number(r.storageMB);
        case 'advice':       return r.advice || '';
        default: return '';
      }
    });

    for (let i = 0; i < list.length; i++) {
      const r = list[i];
      const cat = classify(r.site);
      const t = classifyType(r.site);
      const tr = document.createElement('tr');
      tr.classList.add('clickable-row');
      tr.dataset.idx = String(state.results.indexOf(r));
      const lastViewedLabel = r.lastViewed
        ? (r.lastViewedApprox ? `≤7d (${fmtDate(r.lastViewed)})` : fmtDate(r.lastViewed))
        : '-';
      const fullUrl = r.site.webUrl || '';
      const shortUrl = trimSiteUrlPrefix(fullUrl);
      tr.innerHTML = `
        <td><span class="badge ${badgeClass(r.advice)}">${esc(r.advice)}</span></td>
        <td>${esc(r.site.displayName || r.site.name || '')}</td>
        <td class="url-cell" title="${esc(fullUrl)}"><a href="${esc(fullUrl)}" target="_blank" rel="noopener" onclick="event.stopPropagation()">${esc(shortUrl)}</a></td>
        <td><span class="badge ${cat === 'OneDrive' ? 'od' : 'spo'}">${esc(t)}</span></td>
        <td>${fmtDate(r.lastModified)}</td>
        <td>${lastViewedLabel}</td>
        <td>${r.daysSinceModified ?? '-'}</td>
        <td>${r.storageMB ?? '-'}</td>`;
      tr.addEventListener('click', (e) => {
        // Don't open modal when clicking the URL link or the owners-popover trigger.
        if (e.target.closest('a, .owners-more, .owners-cell')) return;
        showDetailModal(r);
      });
      tbody.appendChild(tr);
    }
    wireOwnersPopovers();
  }

  function trimSiteUrlPrefix(url) {
    if (!url) return '';
    // Strip the protocol+host so the (long, repeated) prefix isn't shown.
    // The full URL stays in href + title for click/copy.
    try {
      const u = new URL(url);
      let p = u.pathname || '/';
      if (p.length > 1) p = p.replace(/\/+$/, '');
      return p || '/';
    } catch {
      return url.replace(/^https?:\/\/[^\/]+/i, '') || url;
    }
  }

  function badgeClass(advice) {
    return ({
      'Clean up':  'cleanup',
      'Review':    'review',
      'Keep':      'keep',
      'No access': 'noaccess',
      'Unknown':   'unknown',
    })[advice] || 'unknown';
  }

  function renderResultNotices() {
    const host = document.getElementById('resultNotices');
    if (!host) return;
    host.innerHTML = '';
    const notices = [];
    // Per-site permission failures
    const denied = state.results.filter(r => r && r.advice === 'No access');
    if (denied.length > 0) {
      const sample = denied.slice(0, 3).map(r => r.site.webUrl || r.site.displayName).filter(Boolean);
      const more = denied.length > sample.length ? ` and ${denied.length - sample.length} more` : '';
      notices.push({
        kind: 'info',
        text: `${denied.length} site${denied.length === 1 ? '' : 's'} could not be analyzed because you lack permission (HTTP 401/403). Filter the table to "No access" to see them. Examples: ${sample.join(', ')}${more}.`,
      });
    }
    for (const n of notices) {
      const div = document.createElement('div');
      div.className = `notice notice-${n.kind}`;
      div.textContent = n.text;
      host.appendChild(div);
    }
  }

  function buildWhyTooltip(r) {
    const lines = [];
    lines.push(`Recommendation: ${r.advice}`);
    lines.push('');
    lines.push(r.reason);
    lines.push('');
    lines.push('--- Raw signals ---');
    for (const [k, v] of Object.entries(r.signals)) {
      lines.push(`${k}: ${v instanceof Date ? v.toISOString() : (v ?? '-')}`);
    }
    if (r.errors && r.errors.length) {
      lines.push('');
      lines.push('--- Errors ---');
      r.errors.forEach(e => lines.push(`• ${e}`));
    }
    return lines.join('\n');
  }

  // ---------- Owners cell + popover ----------
  function renderOwnersCell(owners) {
    if (!owners || owners.length === 0) return '<span class="muted">-</span>';
    const first = owners[0];
    const label = esc(first.title || first.email || first.loginName || '');
    if (owners.length === 1) {
      return `<span class="owners-cell" title="${esc(first.email || first.loginName || '')}">${label}</span>`;
    }
    const more = owners.length - 1;
    const data = encodeURIComponent(JSON.stringify(owners));
    return `<span class="owners-cell">${label} <button type="button" class="owners-more" data-owners="${data}">+${more}</button></span>`;
  }
  function wireOwnersPopovers() {
    document.querySelectorAll('.owners-more').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const list = JSON.parse(decodeURIComponent(btn.dataset.owners));
        showOwnersPopover(btn, list);
      });
    });
    document.addEventListener('click', (e) => {
      const pop = document.getElementById('ownersPopover');
      if (!pop || pop.classList.contains('hidden')) return;
      if (!pop.contains(e.target) && !e.target.closest('.owners-more')) pop.classList.add('hidden');
    });
  }
  function showOwnersPopover(anchor, owners) {
    const pop = document.getElementById('ownersPopover');
    pop.innerHTML = '<div class="popover-title">Owners</div><ul class="owners-list">' +
      owners.map(o => {
        const t = esc(o.title || o.email || o.loginName || '');
        const em = o.email ? `<a href="mailto:${esc(o.email)}">${esc(o.email)}</a>` : '';
        return `<li><strong>${t}</strong>${em ? '<br><span class="muted">' + em + '</span>' : ''}</li>`;
      }).join('') + '</ul>';
    positionPopover(pop, anchor);
  }
  function positionPopover(pop, anchor) {
    pop.classList.remove('hidden');
    const r = anchor.getBoundingClientRect();
    const pr = pop.getBoundingClientRect();
    let left = r.left;
    let top  = r.bottom + 4 + window.scrollY;
    if (left + pr.width > window.innerWidth - 12) left = Math.max(8, window.innerWidth - pr.width - 12);
    pop.style.left = `${left}px`;
    pop.style.top  = `${top}px`;
  }

  // ---------- Excel-style column filter dropdowns ----------
  function wireColumnFilterHeaders() {
    document.querySelectorAll('#resultTable thead th[data-filter]').forEach(th => {
      th.classList.add('filterable-col');
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'col-filter-btn';
      btn.title = 'Filter';
      btn.textContent = '▾';
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        openColumnFilter(th.dataset.filter, btn);
      });
      th.appendChild(btn);
    });
    document.addEventListener('click', (e) => {
      const pop = document.getElementById('columnFilterPopover');
      if (!pop || pop.classList.contains('hidden')) return;
      if (!pop.contains(e.target) && !e.target.closest('.col-filter-btn')) pop.classList.add('hidden');
    });
  }
  function refreshColumnFilterIndicators() {
    document.querySelectorAll('#resultTable thead th[data-filter]').forEach(th => {
      const k = th.dataset.filter;
      const active = state.colFilters[k] && state.colFilters[k].size > 0;
      th.classList.toggle('filter-active', !!active);
    });
  }
  function openColumnFilter(colKey, anchor) {
    const pop = document.getElementById('columnFilterPopover');
    const valueOf = colKey === 'type'
      ? (r) => classifyType(r.site)
      : (r) => r.advice || '';
    const counts = new Map();
    for (const r of state.results) {
      const v = valueOf(r) || '';
      counts.set(v, (counts.get(v) || 0) + 1);
    }
    const values = [...counts.keys()].sort();
    const hidden = state.colFilters[colKey];
    const allChecked = values.every(v => !hidden.has(v));
    pop.innerHTML = `
      <div class="popover-title">Filter ${colKey === 'type' ? 'type' : 'recommendation'}</div>
      <label class="col-filter-all"><input type="checkbox" id="cfAll" ${allChecked ? 'checked' : ''} /> (Select all)</label>
      <div class="col-filter-list">${values.map((v, i) => {
        const checked = !hidden.has(v);
        return `<label><input type="checkbox" class="cf-item" data-val="${esc(v)}" ${checked ? 'checked' : ''} /> ${esc(v) || '<em>(empty)</em>'} <span class="muted">(${counts.get(v)})</span></label>`;
      }).join('')}</div>`;
    positionPopover(pop, anchor);
    pop.querySelector('#cfAll').addEventListener('change', (e) => {
      const on = e.target.checked;
      pop.querySelectorAll('.cf-item').forEach(cb => { cb.checked = on; });
      applyColumnFilter(colKey, pop);
    });
    pop.querySelectorAll('.cf-item').forEach(cb => {
      cb.addEventListener('change', () => applyColumnFilter(colKey, pop));
    });
  }
  function applyColumnFilter(colKey, pop) {
    const hidden = state.colFilters[colKey];
    hidden.clear();
    pop.querySelectorAll('.cf-item').forEach(cb => {
      if (!cb.checked) hidden.add(cb.dataset.val);
    });
    renderResults();
  }

  // ---------- Site detail modal ----------
  function wireDetailModal() {
    const modal = document.getElementById('detailModal');
    if (!modal) return;
    modal.querySelector('.modal-backdrop').addEventListener('click', hideDetailModal);
    modal.querySelector('#detailCloseBtn').addEventListener('click', hideDetailModal);
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !modal.classList.contains('hidden')) hideDetailModal();
    });
  }
  function hideDetailModal() {
    const modal = document.getElementById('detailModal');
    if (modal) modal.classList.add('hidden');
  }
  function showDetailModal(r) {
    const modal = document.getElementById('detailModal');
    if (!modal) return;
    document.getElementById('detailTitle').textContent = r.site.displayName || r.site.name || '(untitled site)';
    const body = document.getElementById('detailBody');
    const cat = classify(r.site);
    const t = classifyType(r.site);
    const fullUrl = r.site.webUrl || '';
    const ownersHtml = r.owners && r.owners.length
      ? '<ul class="owners-list">' + r.owners.map(o =>
          `<li><strong>${esc(o.title || o.email || '')}</strong>${o.email ? '<br><span class="muted"><a href="mailto:' + esc(o.email) + '">' + esc(o.email) + '</a></span>' : ''}</li>`
        ).join('') + '</ul>'
      : '<span class="muted">No owners detected</span>';
    const rows = [
      ['URL', `<a href="${esc(fullUrl)}" target="_blank" rel="noopener">${esc(fullUrl)}</a>`],
      ['Type', `<span class="badge ${cat === 'OneDrive' ? 'od' : 'spo'}">${esc(t)}</span>`],
      ['Recommendation', `<span class="badge ${badgeClass(r.advice)}">${esc(r.advice)}</span>`],
      ['Why', esc(r.reason || '')],
      ['Owners', ownersHtml],
      ['Last modified', fmtDate(r.lastModified)],
      ['Last viewed', r.lastViewed ? (r.lastViewedApprox ? `≤7d (${fmtDate(r.lastViewed)})` : fmtDate(r.lastViewed)) : '-'],
      ['Days since modified', r.daysSinceModified ?? '-'],
      ['Days since viewed', r.daysSinceViewed ?? '-'],
      ['Page views (all-time)', r.pageViews ?? '-'],
      ['Views (last 7 days)',  r.recentViews ?? '-'],
      ['Views (last 30 days)', r.last30Views ?? '-'],
      ['Storage (MB)', r.storageMB ?? '-'],
    ];
    let html = '<dl class="detail-list">';
    for (const [k, v] of rows) html += `<dt>${esc(k)}</dt><dd>${v}</dd>`;
    html += '</dl>';
    html += '<details class="detail-signals"><summary>Raw signals</summary><dl class="detail-list small">';
    for (const [k, v] of Object.entries(r.signals || {})) {
      const display = v instanceof Date ? v.toISOString() : (v == null ? '-' : String(v));
      html += `<dt>${esc(k)}</dt><dd>${esc(display)}</dd>`;
    }
    html += '</dl></details>';
    if (r.errors && r.errors.length) {
      html += '<details class="detail-errors" open><summary>Errors</summary><ul>' +
        r.errors.map(e => `<li>${esc(e)}</li>`).join('') + '</ul></details>';
    }
    body.innerHTML = html;
    modal.classList.remove('hidden');
  }

  // ---------- Tooltips ----------
  let activeTip = null;
  function wireTooltips() {
    document.querySelectorAll('.why-cell').forEach(el => {
      el.addEventListener('mouseenter', showTip);
      el.addEventListener('mouseleave', hideTip);
      el.addEventListener('click', showTip);
    });
  }
  function showTip(e) {
    hideTip();
    const text = e.currentTarget.dataset.tip;
    if (!text) return;
    const tip = document.createElement('div');
    tip.className = 'tooltip';
    tip.textContent = text;
    document.body.appendChild(tip);
    const rect = e.currentTarget.getBoundingClientRect();
    const tipRect = tip.getBoundingClientRect();
    let left = rect.right + 8;
    let top  = rect.top + window.scrollY;
    if (left + tipRect.width > window.innerWidth - 12) left = rect.left - tipRect.width - 8;
    if (left < 8) left = 8;
    tip.style.left = `${left}px`;
    tip.style.top  = `${top}px`;
    activeTip = tip;
  }
  function hideTip() { if (activeTip) { activeTip.remove(); activeTip = null; } }

  // ---------- Progress UI ----------
  function showProgress(label, total) {
    document.getElementById('progressBox').classList.remove('hidden');
    document.getElementById('progressLabel').textContent = label;
    document.getElementById('progressEta').textContent = '';
    if (total === null) {
      document.getElementById('progressFill').style.width = '20%';
      document.getElementById('progressFill').style.transition = 'width 1s';
      document.getElementById('progressCount').textContent = '';
    } else {
      updateProgress(0, total, 0);
    }
  }
  function updateProgress(done, total, etaSec) {
    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
    document.getElementById('progressFill').style.width = `${pct}%`;
    document.getElementById('progressCount').textContent = `${done}/${total} (${pct}%)`;
    if (etaSec > 0) {
      const m = Math.floor(etaSec / 60);
      const s = etaSec % 60;
      document.getElementById('progressEta').textContent = `ETA: ${m > 0 ? m + 'm ' : ''}${s}s`;
    } else {
      document.getElementById('progressEta').textContent = '';
    }
  }
  function hideProgress() {
    document.getElementById('progressBox').classList.add('hidden');
  }

  // ---------- Helpers ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function fmtDate(d) {
    if (!d) return '-';
    if (typeof d === 'string') d = new Date(d);
    if (!(d instanceof Date) || isNaN(d.getTime())) return '-';
    return d.toISOString().substring(0, 10);
  }

  // ---------- Bootstrap ----------
  (async function init() {
    Log.info('SPOTrim starting...');
    try {
      const acc = await Auth.init();
      const cfg = Auth.getAuthConfig();
      const redir = Auth.getRedirectUri();
      // Surface the active redirect URI so users can verify their app registration
      const note = document.createElement('p');
      note.style.fontSize = '12px';
      note.style.color = 'var(--muted)';
      note.innerHTML = `Active redirect URI: <code>${redir}</code> &nbsp; · &nbsp; Client ID: <code>${cfg.clientId}</code>`;
      const card = document.getElementById('connectCard');
      if (card && !card.querySelector('.redir-note')) {
        note.classList.add('redir-note');
        card.appendChild(note);
      }
      if (acc) {
        Log.info(`Restored session: ${acc.username}`);
        setSignedInUI(acc);
      } else {
        setSignedInUI(null);
      }
    } catch (e) {
      Log.err('Init failed:', e);
      toast(`Init failed: ${e.message}`, 10000);
      const card = document.getElementById('connectCard');
      if (card) {
        const err = document.createElement('p');
        err.style.color = 'var(--danger)';
        err.textContent = `Could not initialize: ${e.message}`;
        card.appendChild(err);
      }
    }
  })();
})();
