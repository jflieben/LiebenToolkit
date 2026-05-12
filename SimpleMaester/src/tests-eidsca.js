// EIDSCA test runner. The test definitions live in /data/eidsca-tests.json
// which is generated from the original Maester PowerShell module by
// build/extract-eidsca.ps1 - this keeps SimpleMaester in lock-step with
// the upstream EIDSCA baseline by re-running the build script after each
// `Install-MaesterTests` refresh.
(() => {
  let _defs = null;

  async function loadDefs() {
    if (_defs) return _defs;
    const res = await fetch('./data/eidsca-tests.json', { cache: 'no-store' });
    if (!res.ok) throw new Error('Could not load eidsca-tests.json (' + res.status + ')');
    _defs = await res.json();
    Log.info(`Loaded ${_defs.length} EIDSCA test definitions`);
    return _defs;
  }

  // Convert a recommended value + operator into a human-readable expectation.
  function describeExpectation(op, expected) {
    if (Array.isArray(expected)) {
      const list = expected.map(v => `'${v}'`).join(', ');
      return op === 'in' ? `one of [${list}]` : `not in [${list}]`;
    }
    const map = { eq: 'equal to', ne: 'not equal to', gt: 'greater than', lt: 'less than', ge: 'greater than or equal to', le: 'less than or equal to' };
    return `${map[op] || op} '${expected}'`;
  }

  function compare(actual, op, expected) {
    if (actual == null) return false;
    const actS = String(actual).toLowerCase();
    if (Array.isArray(expected)) {
      const set = expected.map(v => String(v).toLowerCase());
      if (op === 'in')    return set.includes(actS);
      if (op === 'notin') return !set.includes(actS);
      return false;
    }
    const expS = String(expected).toLowerCase();
    switch (op) {
      case 'eq': return actS === expS;
      case 'ne': return actS !== expS;
      case 'gt': return Number(actual) >  Number(expected);
      case 'lt': return Number(actual) <  Number(expected);
      case 'ge': return Number(actual) >= Number(expected);
      case 'le': return Number(actual) <= Number(expected);
      default: return false;
    }
  }

  function makeTitle(def) {
    // Convert "Checks if X is set to 'Y'" into "X" for cleaner table display.
    const t = def.title || def.checkId;
    return t.replace(/^Checks if\s*/i, '').replace(/\s+is set to.+$/i, '');
  }

  // Run a single EIDSCA control. Returns a result row.
  async function runOne(def, ctx) {
    const start = performance.now();
    const baseRow = {
      id: 'EIDSCA.' + def.checkId,
      title: makeTitle(def),
      description: def.description,
      detailMd: def.detailMd,
      severity: def.severity || 'Info',
      tag: 'EIDSCA',
      docUrl: def.docUrl,
      links: def.links,
      relativeUri: def.relativeUri,
      apiVersion: def.apiVersion,
      propertyPath: def.propertyPath,
      operator: def.operator,
      recommended: def.recommended,
      expectation: describeExpectation(def.operator, def.recommended),
    };
    try {
      const data = await Graph.graph(def.relativeUri, { apiVersion: def.apiVersion });
      // Skip-if guards (e.g. older tenants where the property doesn't exist).
      if (def.skipIf && def.skipIf.requireMatch) {
        const flat = JSON.stringify(data);
        if (!flat.includes(def.skipIf.requireMatch)) {
          return { ...baseRow, status: 'Skipped', actual: null, reason: 'Setting not available in this tenant', durationMs: Math.round(performance.now() - start) };
        }
      }
      const actual = Graph.pickPath(data, def.propertyPath);
      const ok = compare(actual, def.operator, def.recommended);
      return {
        ...baseRow,
        status: ok ? 'Passed' : 'Failed',
        actual: actual === undefined ? null : actual,
        reason: ok
          ? `Configured value ${JSON.stringify(actual)} matches the EIDSCA recommendation.`
          : `Configured value ${JSON.stringify(actual)} does not match the EIDSCA recommendation (${describeExpectation(def.operator, def.recommended)}).`,
        durationMs: Math.round(performance.now() - start),
      };
    } catch (e) {
      // 403 most likely means we lack the scope for that endpoint - report as Skipped.
      if (e.status === 403 || e.status === 401) {
        return { ...baseRow, status: 'Skipped', actual: null, reason: `Insufficient permissions to read ${def.relativeUri}: ${e.message}`, durationMs: Math.round(performance.now() - start) };
      }
      // 404 sometimes means a tenant-specific resource isn't provisioned (e.g. a new auth method not enabled).
      if (e.status === 404) {
        return { ...baseRow, status: 'Skipped', actual: null, reason: `Endpoint not available in this tenant (${def.relativeUri}).`, durationMs: Math.round(performance.now() - start) };
      }
      return { ...baseRow, status: 'Error', actual: null, reason: e.message, durationMs: Math.round(performance.now() - start) };
    }
  }

  // Map a checkId prefix to a human-friendly run-category.
  // Maester's tests overview tags every EIDSCA test as "General", so we follow that.
  // See https://maester.dev/docs/tests/ for the canonical taxonomy.
  function runCategoryFor(_checkId) {
    return 'General';
  }

  // Returns the full catalog as test descriptors the registry can consume.
  async function buildCatalog() {
    const defs = await loadDefs();
    return defs.map(def => ({
      id: 'EIDSCA.' + def.checkId,
      title: makeTitle(def),
      severity: def.severity || 'Info',
      tag: 'EIDSCA',
      category: 'EIDSCA',
      runCategory: runCategoryFor(def.checkId),
      docUrl: def.docUrl,
      implemented: true,
      run: (ctx) => runOne(def, ctx),
    }));
  }

  window.TestsEIDSCA = { buildCatalog, loadDefs };
})();
