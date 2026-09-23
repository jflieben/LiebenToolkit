// Promise-based concurrency limiter with cancellation
(() => {
  class CancelToken {
    constructor() { this.cancelled = false; this._listeners = []; }
    cancel() { this.cancelled = true; this._listeners.forEach(fn => { try { fn(); } catch {} }); }
    onCancel(fn) { this._listeners.push(fn); }
    throwIfCancelled() { if (this.cancelled) throw new Error('Cancelled'); }
  }

  async function pmap(items, mapper, { concurrency = 8, onProgress, cancelToken } = {}) {
    const results = new Array(items.length);
    let next = 0;
    let done = 0;
    const errors = [];

    async function worker() {
      while (true) {
        if (cancelToken && cancelToken.cancelled) return;
        const i = next++;
        if (i >= items.length) return;
        try {
          results[i] = await mapper(items[i], i);
        } catch (e) {
          errors.push({ item: items[i], error: e });
          results[i] = { __error: e };
        } finally {
          done++;
          if (onProgress) onProgress(done, items.length);
        }
      }
    }

    const workers = Array.from({ length: Math.min(concurrency, items.length) }, () => worker());
    await Promise.all(workers);
    return { results, errors };
  }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  window.Concurrency = { pmap, sleep, CancelToken };
})();
