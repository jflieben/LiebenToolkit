// Promise-based concurrency limiter with cancellation & pause support
(() => {
  class CancelToken {
    constructor() {
      this.cancelled = false;
      this.paused = false;
      this._listeners = [];
      this._ac = (typeof AbortController !== 'undefined') ? new AbortController() : null;
    }
    get signal() { return this._ac ? this._ac.signal : undefined; }
    cancel() {
      if (this.cancelled) return;
      this.cancelled = true;
      if (this._ac) { try { this._ac.abort(); } catch {} }
      this._listeners.forEach(fn => { try { fn(); } catch {} });
    }
    onCancel(fn) { this._listeners.push(fn); }
    throwIfCancelled() { if (this.cancelled) throw new Error('Cancelled'); }
    pause() { this.paused = true; }
    resume() { this.paused = false; }
    async waitIfPaused() {
      while (this.paused && !this.cancelled) { await sleep(200); }
      this.throwIfCancelled();
    }
  }

  // Race a promise against a cancel signal. Rejects promptly with 'Cancelled' when the
  // signal aborts; cleans up its listener when the promise settles first (no leak).
  function raceCancel(promise, signal) {
    if (!signal) return promise;
    if (signal.aborted) return Promise.reject(new Error('Cancelled'));
    return new Promise((resolve, reject) => {
      const onAbort = () => reject(new Error('Cancelled'));
      signal.addEventListener('abort', onAbort, { once: true });
      promise.then(
        (v) => { signal.removeEventListener('abort', onAbort); resolve(v); },
        (e) => { signal.removeEventListener('abort', onAbort); reject(e); }
      );
    });
  }

  async function pmap(items, mapper, { concurrency = 4, onProgress, cancelToken } = {}) {
    const results = new Array(items.length);
    let next = 0;
    let done = 0;
    const errors = [];

    async function worker() {
      while (true) {
        if (cancelToken && cancelToken.cancelled) return;
        if (cancelToken && cancelToken.paused) { await cancelToken.waitIfPaused(); }
        const i = next++;
        if (i >= items.length) return;
        try {
          results[i] = await mapper(items[i], i);
        } catch (e) {
          errors.push({ item: items[i], error: e });
          results[i] = { __error: e };
        } finally {
          done++;
          if (onProgress) onProgress(done, items.length, i);
        }
      }
    }

    const workers = Array.from({ length: Math.min(concurrency, items.length) }, () => worker());
    await Promise.all(workers);
    return { results, errors };
  }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  window.Concurrency = { pmap, sleep, CancelToken, raceCancel };
})();
