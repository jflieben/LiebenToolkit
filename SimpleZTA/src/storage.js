(() => {
  const SCANS_KEY = 'SimpleZTA:scans:v1';
  const LAST_SELECTION_KEY = 'SimpleZTA:lastSelection:v1';
  const PROFILES_KEY = 'SimpleZTA:profiles:v1';

  // Stored API traces are capped: they hold raw tenant data and are the bulk of a scan.
  const TRACE_MAX_ENTRIES = 25;
  const TRACE_PREVIEW_MAX = 1000;
  const TRACE_REQUEST_MAX = 500;

  function loadScans() {
    try {
      const raw = localStorage.getItem(SCANS_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    } catch {
      return [];
    }
  }

  function saveScans(scans) {
    localStorage.setItem(SCANS_KEY, JSON.stringify(scans));
  }

  function isQuotaExceeded(err) {
    if (!err) return false;
    const name = `${err.name || ''}`;
    const code = Number(err.code || 0);
    if (name === 'QuotaExceededError' || name === 'NS_ERROR_DOM_QUOTA_REACHED') return true;
    if (code === 22 || code === 1014) return true;
    return /quota/i.test(`${err.message || ''}`);
  }

  function trimText(value, maxLen) {
    const text = `${value || ''}`;
    if (text.length <= maxLen) return text;
    return `${text.slice(0, maxLen)}...`;
  }

  function capText(value, max) {
    const text = value == null ? '' : `${value}`;
    return text.length <= max ? text : `${text.slice(0, max)}\n...truncated (${text.length} chars)...`;
  }

  function capTrace(trace) {
    if (!Array.isArray(trace)) return [];
    const kept = trace.slice(0, TRACE_MAX_ENTRIES).map(e => ({
      method: e?.method,
      url: e?.url,
      status: e?.status,
      ok: e?.ok,
      durationMs: e?.durationMs,
      requestPreview: capText(e?.requestPreview, TRACE_REQUEST_MAX),
      responsePreview: capText(e?.responsePreview, TRACE_PREVIEW_MAX),
      nextLink: e?.nextLink || null,
    }));
    if (trace.length > TRACE_MAX_ENTRIES) {
      kept.push({ method: '', url: `${trace.length - TRACE_MAX_ENTRIES} more call(s) not stored`, status: '', ok: true, durationMs: 0, requestPreview: '', responsePreview: '', nextLink: null });
    }
    return kept;
  }

  function compactResult(result, level) {
    const base = {
      id: result?.id,
      title: result?.title,
      status: result?.status,
      message: result?.message,
      durationMs: result?.durationMs,
      pillar: result?.pillar,
      risk: result?.risk,
      category: result?.category,
      implementationLevel: result?.implementationLevel,
      implementationNotes: result?.implementationNotes,
      documentationUrl: result?.documentationUrl,
      documentationSearchUrl: result?.documentationSearchUrl,
      evidence: result?.evidence,
      apiTrace: capTrace(result?.apiTrace),
    };

    if (level >= 1) {
      // API traces are the largest payload and can be safely dropped from persistence.
      base.apiTrace = [];
    }
    if (level >= 2) {
      // Drop potentially large evidence payloads.
      base.evidence = null;
      base.message = trimText(base.message, 800);
    }
    if (level >= 3) {
      // Keep only summary fields needed for history/result lists.
      return {
        id: base.id,
        title: base.title,
        status: base.status,
        message: trimText(base.message, 300),
        durationMs: base.durationMs,
        pillar: base.pillar,
        risk: base.risk,
        category: base.category,
      };
    }

    return base;
  }

  function compactScan(scan, level) {
    return {
      ...scan,
      results: (scan?.results || []).map(r => compactResult(r, level)),
    };
  }

  function saveScan(scan) {
    const existing = loadScans();
    const levels = [0, 1, 2, 3];
    const retentionCandidates = Array.from(new Set([
      existing.length,
      Math.min(existing.length, 40),
      Math.min(existing.length, 20),
      Math.min(existing.length, 10),
      Math.min(existing.length, 5),
      Math.min(existing.length, 2),
      0,
    ])).sort((a, b) => b - a);

    let lastErr = null;
    for (const level of levels) {
      const scanCandidate = compactScan(scan, level);
      for (const keep of retentionCandidates) {
        const candidate = [...existing.slice(-keep), scanCandidate];
        try {
          saveScans(candidate);
          const degraded = level > 0;
          const pruned = keep < existing.length;
          return {
            saved: true,
            degraded,
            pruned,
            warning: (degraded || pruned)
              ? 'Scan results were partially compacted/pruned to fit browser storage quota.'
              : '',
          };
        } catch (err) {
          lastErr = err;
          if (!isQuotaExceeded(err)) throw err;
        }
      }
    }

    return {
      saved: false,
      degraded: true,
      pruned: true,
      warning: 'Could not persist scan results due to browser storage quota. Export results if needed and clear history.',
      error: lastErr ? `${lastErr.message || lastErr}` : 'quota exceeded',
    };
  }

  function clearScans() {
    try { localStorage.removeItem(SCANS_KEY); } catch { /* storage blocked */ }
  }

  function storedBytes() {
    try { return (localStorage.getItem(SCANS_KEY) || '').length * 2; } catch { return 0; }
  }

  // One line for confirm dialogs: how many scans, which period, how big.
  function describeHistory(scans) {
    const list = scans || loadScans();
    if (!list.length) return 'No stored scans';
    const dates = list.map(x => `${x.startedAt || ''}`.slice(0, 10)).filter(Boolean).sort();
    const tenants = new Set(list.map(x => x.tenantDomain || x.tenantId).filter(Boolean));
    const kb = Math.round(storedBytes() / 1024);
    return `${list.length} scan${list.length === 1 ? '' : 's'} of ${tenants.size || 1} tenant${tenants.size > 1 ? 's' : ''}, ${dates[0]} to ${dates[dates.length - 1]}, about ${kb} KB`;
  }

  function getScan(id) {
    return loadScans().find(s => s.id === id) || null;
  }

  function rememberSelection(ids) {
    try { localStorage.setItem(LAST_SELECTION_KEY, JSON.stringify(ids || [])); } catch { /* storage full or blocked */ }
  }

  function loadSelection() {
    try {
      const raw = localStorage.getItem(LAST_SELECTION_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    } catch {
      return [];
    }
  }

  function computeSummary(results) {
    const s = { total: results.length, passed: 0, failed: 0, skipped: 0, error: 0 };
    for (const r of results) {
      if (r.status === 'Passed') s.passed += 1;
      else if (r.status === 'Failed') s.failed += 1;
      else if (r.status === 'Skipped') s.skipped += 1;
      else if (r.status === 'Error') s.error += 1;
    }
    return s;
  }

  function getAverageDurationByTest() {
    const scans = loadScans();
    const map = new Map();

    for (const scan of scans) {
      for (const result of (scan.results || [])) {
        const ms = Number(result?.durationMs || 0);
        if (!ms || Number.isNaN(ms)) continue;
        const id = `${result.id || ''}`;
        if (!id) continue;
        if (!map.has(id)) map.set(id, { total: 0, count: 0 });
        const cur = map.get(id);
        cur.total += ms;
        cur.count += 1;
      }
    }

    const out = {};
    for (const [id, stat] of map.entries()) {
      out[id] = Math.round(stat.total / stat.count);
    }
    return out;
  }

  function loadProfiles() {
    try {
      const raw = localStorage.getItem(PROFILES_KEY);
      const obj = raw ? JSON.parse(raw) : {};
      if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return {};
      const out = {};
      for (const [k, v] of Object.entries(obj)) if (cleanProfileName(k) === k && Array.isArray(v)) out[k] = v.map(x => `${x}`);
      return out;
    } catch {
      return {};
    }
  }

  // Profile names become object keys and option labels: plain text, short, no prototype keys.
  function cleanProfileName(name) {
    const clean = `${name ?? ''}`.replace(/[\u0000-\u001f\u007f]/g, '').trim().slice(0, 80);
    return ['__proto__', 'constructor', 'prototype'].includes(clean) ? '' : clean;
  }

  function saveProfile(name, ids) {
    const key = cleanProfileName(name);
    if (!key) throw new Error('Invalid profile name');
    const profiles = loadProfiles();
    profiles[key] = Array.isArray(ids) ? ids.map(x => `${x}`) : [];
    localStorage.setItem(PROFILES_KEY, JSON.stringify(profiles));
    return profiles;
  }

  function deleteProfile(name) {
    const profiles = loadProfiles();
    delete profiles[name];
    localStorage.setItem(PROFILES_KEY, JSON.stringify(profiles));
    return profiles;
  }

  window.Store = {
    loadScans,
    saveScan,
    clearScans,
    getScan,
    computeSummary,
    getAverageDurationByTest,
    rememberSelection,
    loadSelection,
    loadProfiles,
    saveProfile,
    deleteProfile,
    cleanProfileName,
    capTrace,
    describeHistory,
    storedBytes,
    TRACE_MAX_ENTRIES,
    TRACE_PREVIEW_MAX,
  };
})();
