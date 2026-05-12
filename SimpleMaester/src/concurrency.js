// A tiny semaphore so we can fan out Graph calls without flooding the throttle limit.
(() => {
  class Semaphore {
    constructor(max) { this.max = max; this.cur = 0; this.q = []; }
    async acquire() {
      if (this.cur < this.max) { this.cur++; return; }
      await new Promise(r => this.q.push(r));
      this.cur++;
    }
    release() { this.cur--; const n = this.q.shift(); if (n) n(); }
    async run(fn) { await this.acquire(); try { return await fn(); } finally { this.release(); } }
  }
  async function pmap(items, max, fn, onProgress) {
    const sem = new Semaphore(max);
    let done = 0;
    return Promise.all(items.map((it, idx) => sem.run(async () => {
      const r = await fn(it, idx);
      done++;
      if (onProgress) onProgress(done, items.length, it);
      return r;
    })));
  }
  window.Concurrency = { Semaphore, pmap };
})();
