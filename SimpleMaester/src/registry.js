// Test registry: combines all sources (EIDSCA, native, stubs) into one indexed catalog.
(() => {
  let _all = null;

  async function buildAll() {
    const [eidsca, native, cis, cisa, intune, gov, stubs] = await Promise.all([
      TestsEIDSCA.buildCatalog(),
      Promise.resolve(TestsNative.buildCatalog()),
      Promise.resolve(window.TestsCIS ? TestsCIS.buildCatalog() : []),
      Promise.resolve(window.TestsCISA ? TestsCISA.buildCatalog() : []),
      Promise.resolve(window.TestsIntune ? TestsIntune.buildCatalog() : []),
      Promise.resolve(window.TestsGovernance ? TestsGovernance.buildCatalog() : []),
      Promise.resolve(TestsStubs.buildCatalog()),
    ]);
    return [...native, ...cis, ...cisa, ...intune, ...gov, ...eidsca, ...stubs];
  }

  async function getAll() {
    if (!_all) _all = await buildAll();
    return _all;
  }
  async function getById(id) {
    const all = await getAll();
    return all.find(t => t.id === id);
  }
  async function categories() {
    const all = await getAll();
    return Array.from(new Set(all.map(t => t.category))).sort();
  }
  async function runCategories() {
    const all = await getAll();
    const groups = new Map();
    for (const t of all) {
      if (!t.implemented) continue;
      const key = t.runCategory || t.category || 'Other';
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(t);
    }
    return Array.from(groups.entries())
      .map(([name, tests]) => ({ name, tests, count: tests.length }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }
  async function stats() {
    const all = await getAll();
    return {
      total: all.length,
      implemented: all.filter(t => t.implemented).length,
      stub: all.filter(t => !t.implemented).length,
    };
  }

  window.Registry = { getAll, getById, categories, runCategories, stats };
})();
