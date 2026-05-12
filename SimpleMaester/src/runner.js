// Test runner. Executes a list of tests with concurrency, progress, cancellation.
// Emits events: progress, log, done. The Runner is also responsible for stamping
// each scan with tenant context and persisting to History.
(() => {
  let _abort = null;

  function emit(target, type, detail) { target.dispatchEvent(new CustomEvent(type, { detail })); }

  async function discoverTenant() {
    try {
      const org = await Graph.graph(`organization?$select=id,displayName,verifiedDomains`, { apiVersion: 'v1.0' });
      const o = org.value?.[0] || {};
      const dom = (o.verifiedDomains || []).find(d => d.isInitial)?.name || (o.verifiedDomains || [])[0]?.name;
      return { tenantId: o.id, tenantDomain: dom, tenantName: o.displayName };
    } catch (e) {
      Log.warn('Tenant discovery failed:', e.message);
      return { tenantId: 'unknown', tenantDomain: 'unknown', tenantName: 'unknown' };
    }
  }

  // The Maester module caches identical Graph responses across tests so 20 EIDSCA
  // checks against /policies/authorizationPolicy result in 1 HTTP call. We do the same:
  // we clear the cache once at the start of the scan; everything inside the scan
  // benefits from caching.
  // Tests run sequentially so we can attach the Graph calls each one made to its
  // own result row (per-test capture would otherwise interleave under concurrency).
  // Most calls are cached after the first hit so the wall-clock impact is small.
  async function run(tests, { target } = {}) {
    if (!target) target = new EventTarget();
    Graph.clearCache();
    if (window.Exo) Exo.clearCache();
    _abort = { cancelled: false };
    const scanId = 'scan-' + Date.now();
    const startedAt = new Date().toISOString();
    const tenant = await discoverTenant();
    emit(target, 'log', { level: 'info', text: `Starting scan against ${tenant.tenantDomain || tenant.tenantId} (${tests.length} tests)` });
    const t0 = performance.now();
    const results = [];
    let done = 0;
    for (const test of tests) {
      if (_abort.cancelled) {
        results.push({ id: test.id, title: test.title, severity: test.severity, tag: test.tag, category: test.category, status: 'Skipped', reason: 'Scan cancelled', durationMs: 0, graphCalls: [] });
        done++; emit(target, 'progress', { done, total: tests.length, current: test });
        continue;
      }
      emit(target, 'log', { level: 'debug', text: `> ${test.id}` });
      // Reset any per-test sub-progress when moving to the next test.
      emit(target, 'progress', { done, total: tests.length, current: test, sub: null });
      let row;
      let calls = [];
      const testScopes = Array.isArray(test.requiredScopes) && test.requiredScopes.length
        ? test.requiredScopes
        : Auth.SCOPES.graphFull;
      let consentOk = true;
      try {
        consentOk = await Auth.ensureScopes(testScopes);
      } catch (e) {
        consentOk = false;
        row = {
          id: test.id,
          status: 'Skipped',
          reason: `Unable to verify consent for required scopes: ${e.message}`,
          durationMs: 0,
        };
      }
      if (!consentOk) {
        if (!row) {
          row = {
            id: test.id,
            status: 'Skipped',
            reason: 'Consent flow was started for this test. Complete consent and run the scan again.',
            durationMs: 0,
          };
        }
        results.push({
          id: test.id, title: test.title, severity: test.severity, tag: test.tag, category: test.category,
          docUrl: test.docUrl, description: test.description, detailMd: test.detailMd,
          ...row,
          graphCalls: [],
        });
        done++;
        emit(target, 'progress', { done, total: tests.length, current: test, row });
        emit(target, 'log', { level: 'warn', text: `[SKIP] ${test.id} - ${row.reason}` });
        continue;
      }
      Graph.beginCapture();
      try {
        const reportSubProgress = ({ done: subDone = 0, total: subTotal = 0, label = '' } = {}) => {
          emit(target, 'progress', {
            done,
            total: tests.length,
            current: test,
            sub: { done: subDone, total: subTotal, label },
          });
        };
        row = await test.run({ tenant, reportSubProgress });
      } catch (e) {
        row = { id: test.id, status: 'Error', reason: e.message, durationMs: 0 };
      } finally {
        calls = Graph.endCapture();
      }
      // Backfill metadata from the catalog entry.
      row = {
        id: test.id, title: test.title, severity: test.severity, tag: test.tag, category: test.category,
        docUrl: test.docUrl, description: test.description, detailMd: test.detailMd,
        ...row,
        graphCalls: calls,
      };
      results.push(row);
      done++;
      emit(target, 'progress', { done, total: tests.length, current: test, row });
      const icon = row.status === 'Passed' ? 'OK ' : row.status === 'Failed' ? 'FAIL' : row.status === 'Skipped' ? 'SKIP' : 'ERR ';
      emit(target, 'log', { level: row.status === 'Failed' || row.status === 'Error' ? 'warn' : 'info', text: `[${icon}] ${test.id} - ${test.title}` });
    }
    const summary = summarise(results);
    const scan = {
      id: scanId, tenantId: tenant.tenantId, tenantDomain: tenant.tenantDomain, tenantName: tenant.tenantName,
      startedAt, durationMs: Math.round(performance.now() - t0),
      results, summary,
      cancelled: _abort.cancelled,
    };
    await History.saveScan(scan);
    emit(target, 'done', { scan });
    _abort = null;
    return scan;
  }

  function cancel() { if (_abort) _abort.cancelled = true; }

  function summarise(results) {
    const c = { total: results.length, passed: 0, failed: 0, skipped: 0, error: 0 };
    for (const r of results) {
      if (r.status === 'Passed') c.passed++;
      else if (r.status === 'Failed') c.failed++;
      else if (r.status === 'Skipped') c.skipped++;
      else c.error++;
    }
    return c;
  }

  window.Runner = { run, cancel, summarise };
})();
