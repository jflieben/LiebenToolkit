(() => {
  let cancelled = false;

  function estimateMsForTest(test, historicalAvg, fallbackMs) {
    const ms = Number(historicalAvg?.[test.id] || 0);
    if (ms > 0) return ms;
    return fallbackMs;
  }

  async function run(tests, options) {
    cancelled = false;
    const opts = options || {};
    const progress = opts.progress || (() => {});
    const log = opts.log || (() => {});
    const account = Auth.getAccount();

    const tenantId = account?.tenantId || '';
    let tenantDomain = '';
    let tenantName = '';
    try {
      const org = await Api.graph('organization', { query: { '$select': 'verifiedDomains,id,displayName' } });
      const first = org?.value?.[0];
      const initial = (first?.verifiedDomains || []).find(d => d.isInitial);
      const preferred = (first?.verifiedDomains || []).find(d => d.isDefault) || initial;
      tenantDomain = preferred?.name || first?.displayName || '';
      tenantName = first?.displayName || '';
    } catch (e) {
      log(`Tenant lookup failed: ${e.message}`);
    }

    const results = [];
    const total = tests.length;
    let done = 0;
    const runStartedAt = Date.now();
    const historicalAvg = Store.getAverageDurationByTest();
    const completedDurations = [];

    for (const test of tests) {
      if (cancelled) {
        log('Run cancelled by user.');
        break;
      }

      const start = performance.now();
      const averageCompletedMs = completedDurations.length
        ? Math.round(completedDurations.reduce((a, b) => a + b, 0) / completedDurations.length)
        : 3500;
      const currentEstimateMs = estimateMsForTest(test, historicalAvg, averageCompletedMs);
      const remainingTests = tests.slice(done + 1);
      const remainingEstimateMs = remainingTests.reduce((acc, t) => acc + estimateMsForTest(t, historicalAvg, averageCompletedMs), 0);

      let lastLongLogSecond = 0;
      let subProgress = null;
      progress({
        done,
        total,
        current: test,
        elapsedTotalMs: Date.now() - runStartedAt,
        elapsedCurrentMs: 0,
        etaMs: currentEstimateMs + remainingEstimateMs,
        currentEstimateMs,
        subProgress,
      });
      log(`Running ${test.id} ${test.title}`);

      const tick = setInterval(() => {
        const elapsedCurrentMs = Math.max(0, Math.round(performance.now() - start));
        const elapsedTotalMs = Date.now() - runStartedAt;
        const currentRemainingMs = Math.max(0, currentEstimateMs - elapsedCurrentMs);
        const etaMs = currentRemainingMs + remainingEstimateMs;

        progress({
          done,
          total,
          current: test,
          elapsedTotalMs,
          elapsedCurrentMs,
          etaMs,
          currentEstimateMs,
          subProgress,
        });

        const elapsedSeconds = Math.floor(elapsedCurrentMs / 1000);
        if (elapsedSeconds >= 10 && elapsedSeconds % 10 === 0 && elapsedSeconds !== lastLongLogSecond) {
          lastLongLogSecond = elapsedSeconds;
          log(`Still running ${test.id} (${elapsedSeconds}s elapsed)...`);
        }
      }, 1000);

      const trace = [];
      Api.setTraceCollector(trace);

      // Port of the module's TestTimeout: a hung control marks itself Error and the run
      // continues. Best-effort — in-flight fetches are abandoned, not aborted.
      const testTimeoutMs = Number(opts.testTimeoutMs) > 0 ? Number(opts.testTimeoutMs) : 180000;

      try {
        let out;
        if (!test.implemented || !test.run) {
          out = {
            id: test.id,
            title: test.title,
            status: 'Skipped',
            message: 'Not yet implemented in browser engine.',
          };
        } else {
          const execution = test.run(test, {
            tenantId,
            tenantDomain,
            report: (message, percent) => {
              subProgress = {
                message: `${message || ''}`,
                percent: Number.isFinite(Number(percent)) ? Number(percent) : null,
              };
            },
          });
          let timeoutHandle;
          const timeout = new Promise((_, reject) => {
            timeoutHandle = setTimeout(() => reject(new Error(`Test timed out after ${Math.round(testTimeoutMs / 1000)}s`)), testTimeoutMs);
          });
          try {
            out = await Promise.race([execution, timeout]);
          } finally {
            clearTimeout(timeoutHandle);
          }
        }
        out.durationMs = Math.round(performance.now() - start);
        out.pillar = test.pillar;
        out.risk = test.risk;
        out.category = test.category;
        out.implementationLevel = test.implementationLevel || 'none';
        out.implementationNotes = test.implementationNotes || null;
        out.documentationUrl = test.documentationUrl || '';
        out.documentationSearchUrl = test.documentationSearchUrl || '';
        out.apiTrace = trace;
        results.push(out);
      } catch (e) {
        results.push({
          id: test.id,
          title: test.title,
          status: 'Error',
          message: e.message || `${e}`,
          durationMs: Math.round(performance.now() - start),
          pillar: test.pillar,
          risk: test.risk,
          category: test.category,
          implementationLevel: test.implementationLevel || 'none',
          implementationNotes: test.implementationNotes || null,
          documentationUrl: test.documentationUrl || '',
          documentationSearchUrl: test.documentationSearchUrl || '',
          apiTrace: trace,
        });
      } finally {
        clearInterval(tick);
        Api.setTraceCollector(null);
      }

      done += 1;
      const duration = Number(results[results.length - 1]?.durationMs || 0);
      if (duration > 0) completedDurations.push(duration);
      progress({
        done,
        total,
        current: test,
        elapsedTotalMs: Date.now() - runStartedAt,
        elapsedCurrentMs: duration,
        etaMs: tests.slice(done).reduce((acc, t) => {
          const avg = completedDurations.length
            ? Math.round(completedDurations.reduce((a, b) => a + b, 0) / completedDurations.length)
            : 3500;
          return acc + estimateMsForTest(t, historicalAvg, avg);
        }, 0),
        currentEstimateMs,
        subProgress,
      });
    }

    // Tenant insight collection (sankey/overview graphics data) runs after the tests so a
    // collector failure can never invalidate test results.
    let tenantInfo = null;
    let tenantInfoNotes = null;
    if (opts.collectTenantInfo !== false && !cancelled && typeof TenantInfo !== 'undefined') {
      try {
        log('Collecting tenant insight data (overview graphics)...');
        const collected = await TenantInfo.collect({
          log,
          days: opts.tenantInfoDays,
          progress: (message) => progress({
            done, total, current: null,
            elapsedTotalMs: Date.now() - runStartedAt,
            elapsedCurrentMs: 0, etaMs: 0, currentEstimateMs: 0,
            subProgress: { message, percent: null },
          }),
        });
        tenantInfo = collected.tenantInfo;
        tenantInfoNotes = collected.notes;
      } catch (e) {
        log(`Tenant insight collection failed: ${e.message || e}`);
      }
    }

    const scan = {
      id: `scan-${Date.now()}`,
      startedAt: new Date().toISOString(),
      tenantId,
      tenantDomain,
      tenantName,
      results,
      summary: Store.computeSummary(results),
      tenantInfo,
      tenantInfoNotes,
      implementationCoverage: {
        totalCatalog: tests.length,
        implemented: tests.filter(t => t.implemented).length,
      },
    };

    const persistence = Store.saveScan(scan);
    if (persistence?.warning) {
      log(`Storage warning: ${persistence.warning}`);
    }
    scan.persistence = persistence || { saved: true };
    return scan;
  }

  window.Runner = {
    run,
    cancel: () => { cancelled = true; },
  };
})();
