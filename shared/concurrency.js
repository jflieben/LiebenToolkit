// Promise pool with cancel and pause.
(() => {
  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

  class CancelToken {
    constructor() {
      this.cancelled = false;
      this.paused = false;
      this._listeners = [];
      this._ac = typeof AbortController !== 'undefined' ? new AbortController() : null;
    }
    get signal() { return this._ac ? this._ac.signal : undefined; }
    cancel() {
      if (this.cancelled) return;
      this.cancelled = true;
      if (this._ac) { try { this._ac.abort(); } catch { /* already aborted */ } }
      this._listeners.forEach((fn) => { try { fn(); } catch { /* listener failed */ } });
    }
    onCancel(fn) { this._listeners.push(fn); }
    throwIfCancelled() { if (this.cancelled) throw new Error('Cancelled'); }
    pause() { this.paused = true; }
    resume() { this.paused = false; }
    async waitIfPaused() {
      while (this.paused && !this.cancelled) await sleep(200);
      this.throwIfCancelled();
    }
  }

  // Rejects with 'Cancelled' as soon as the signal aborts.
  function raceCancel(promise, signal) {
    if (!signal) return promise;
    if (signal.aborted) return Promise.reject(new Error('Cancelled'));
    return new Promise((resolve, reject) => {
      const onAbort = () => reject(new Error('Cancelled'));
      signal.addEventListener('abort', onAbort, { once: true });
      promise.then(
        (v) => { signal.removeEventListener('abort', onAbort); resolve(v); },
        (e) => { signal.removeEventListener('abort', onAbort); reject(e); },
      );
    });
  }

  // Map items with at most `concurrency` in flight. Errors are collected, not thrown.
  async function pmap(items, mapper, { concurrency = 8, onProgress, cancelToken } = {}) {
    const results = new Array(items.length);
    const errors = [];
    let next = 0;
    let done = 0;
    async function worker() {
      for (;;) {
        if (cancelToken && cancelToken.cancelled) return;
        if (cancelToken && cancelToken.paused) await cancelToken.waitIfPaused().catch(() => {});
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
          if (onProgress) onProgress(done, items.length, i);
        }
      }
    }
    await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, worker));
    return { results, errors };
  }

  // Time-left estimate from progress so far. Returns '' until there is enough signal.
  function eta(startMs, done, total) {
    if (!done || done >= total) return '';
    const left = ((Date.now() - startMs) / done) * (total - done) / 1000;
    if (!isFinite(left) || left < 1) return '';
    if (left < 90) return `about ${Math.round(left)}s left`;
    if (left < 5400) return `about ${Math.round(left / 60)} min left`;
    return `about ${(left / 3600).toFixed(1)} h left`;
  }

  window.Concurrency = { pmap, sleep, CancelToken, raceCancel, eta };
})();
