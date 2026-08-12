(() => {
  const SCANS_KEY = 'SimpleZTA:scans:v1';
  const LAST_SELECTION_KEY = 'SimpleZTA:lastSelection:v1';
  const PROFILES_KEY = 'SimpleZTA:profiles:v1';

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
      apiTrace: result?.apiTrace,
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
    localStorage.removeItem(SCANS_KEY);
  }

  function getScan(id) {
    return loadScans().find(s => s.id === id) || null;
  }

  function rememberSelection(ids) {
    localStorage.setItem(LAST_SELECTION_KEY, JSON.stringify(ids || []));
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

  function toCsv(scan) {
    const rows = [
      ['scanId', 'startedAt', 'tenantId', 'tenantDomain', 'testId', 'title', 'status', 'risk', 'pillar', 'durationMs', 'message'],
    ];
    for (const r of scan.results) {
      rows.push([
        scan.id,
        scan.startedAt,
        scan.tenantId || '',
        scan.tenantDomain || '',
        r.id,
        r.title,
        r.status,
        r.risk || '',
        r.pillar || '',
        r.durationMs,
        (r.message || '').replace(/\r?\n/g, ' '),
      ]);
    }
    return rows.map(cols => cols.map(c => `"${`${c ?? ''}`.replace(/"/g, '""')}"`).join(',')).join('\n');
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
      return obj && typeof obj === 'object' ? obj : {};
    } catch {
      return {};
    }
  }

  function saveProfile(name, ids) {
    const profiles = loadProfiles();
    profiles[name] = Array.isArray(ids) ? ids.map(x => `${x}`) : [];
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
    toCsv,
    getAverageDurationByTest,
    rememberSelection,
    loadSelection,
    loadProfiles,
    saveProfile,
    deleteProfile,
  };
})();
