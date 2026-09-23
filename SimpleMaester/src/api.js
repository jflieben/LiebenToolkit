// Microsoft Graph wrapper with throttle, retry and per-scan caching.
// Maester's PowerShell code calls Invoke-MtGraphRequest with a built-in cache
// keyed on URL+ApiVersion+Method. We do the same here so that running e.g.
// 20 EIDSCA tests that all hit /policies/authorizationPolicy issues exactly one
// HTTPS call. The cache is reset by Runner before every scan.
(() => {
  const cache = new Map();
  function _key(url, method) { return method + '|' + url; }
  function clearCache() { cache.clear(); }

  // Per-test Graph call recorder. Runner calls beginCapture() before test.run()
  // and endCapture() after, getting back a small array of {method,url,status,
  // durationMs,fromCache,response,truncated} entries to attach to the result row.
  // Limits keep IndexedDB / memory usage sane:
  //   MAX_CALLS_PER_TEST: only the first N calls are kept (older overflow is summarised)
  //   MAX_RESPONSE_CHARS: each response body is truncated to this many chars
  const MAX_CALLS_PER_TEST = 8;
  const MAX_RESPONSE_CHARS = 1500;
  let _capture = null; // { calls: [], overflow: 0 }

  function beginCapture() { _capture = { calls: [], overflow: 0 }; }
  function endCapture() {
    const c = _capture;
    _capture = null;
    if (!c) return [];
    if (c.overflow > 0) {
      c.calls.push({ method: 'INFO', url: `... and ${c.overflow} more call(s) not shown`, status: 0, durationMs: 0, fromCache: false, response: '', truncated: false });
    }
    return c.calls;
  }
  function _record(entry) {
    if (!_capture) return;
    if (_capture.calls.length >= MAX_CALLS_PER_TEST) { _capture.overflow++; return; }
    _capture.calls.push(entry);
  }
  function _shorten(data) {
    if (data == null) return { response: '', truncated: false };
    let str;
    try { str = typeof data === 'string' ? data : JSON.stringify(data); }
    catch { str = String(data); }
    if (str.length <= MAX_RESPONSE_CHARS) return { response: str, truncated: false };
    return { response: str.substring(0, MAX_RESPONSE_CHARS), truncated: true };
  }

  // GET helper that follows @odata.nextLink and concatenates value arrays.
  async function graph(relativeUri, opts = {}) {
    const apiVersion = opts.apiVersion || 'v1.0';
    const method = (opts.method || 'GET').toUpperCase();
    const all = !!opts.all;
    const url = relativeUri.startsWith('http')
      ? relativeUri
      : `https://graph.microsoft.com/${apiVersion}/${relativeUri.replace(/^\//,'')}`;

    const ck = _key(url, method);
    if (method === 'GET' && cache.has(ck) && !opts.bypassCache) {
      const cached = cache.get(ck);
      const { response, truncated } = _shorten(cached);
      _record({ method, url, status: 200, durationMs: 0, fromCache: true, response, truncated });
      return cached;
    }

    const tok = await Auth.getToken(opts.tokenScopes || Auth.SCOPES.graphFull);
    const doFetch = async (u) => {
      let attempt = 0;
      while (true) {
        const t0 = performance.now();
        const res = await fetch(u, {
          method,
          headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json', 'ConsistencyLevel': opts.consistencyLevel || 'eventual' },
          body: opts.body ? JSON.stringify(opts.body) : undefined,
        });
        if (res.status === 429 || res.status === 503 || res.status === 504) {
          attempt++;
          if (attempt > 5) throw new Error(`Graph throttled and gave up after 5 retries: ${res.status}`);
          const wait = parseInt(res.headers.get('Retry-After') || '0', 10) * 1000 || (Math.pow(2, attempt) * 1000);
          Log.warn(`Graph ${res.status}, sleeping ${wait}ms`);
          await new Promise(r => setTimeout(r, wait));
          continue;
        }
        if (res.status === 404) {
          _record({ method, url: u, status: 404, durationMs: Math.round(performance.now() - t0), fromCache: false, response: '', truncated: false });
          const err = new Error(`Graph 404 for ${u}`);
          err.status = 404;
          throw err;
        }
        if (!res.ok) {
          const t = await res.text().catch(() => '');
          const sh = _shorten(t);
          _record({ method, url: u, status: res.status, durationMs: Math.round(performance.now() - t0), fromCache: false, response: sh.response, truncated: sh.truncated });
          const err = new Error(`Graph ${res.status} for ${u}: ${t.substring(0, 400)}`);
          err.status = res.status;
          err.body = t;
          throw err;
        }
        if (res.status === 204) {
          _record({ method, url: u, status: 204, durationMs: Math.round(performance.now() - t0), fromCache: false, response: '', truncated: false });
          return null;
        }
        const json = await res.json();
        const sh = _shorten(json);
        _record({ method, url: u, status: res.status, durationMs: Math.round(performance.now() - t0), fromCache: false, response: sh.response, truncated: sh.truncated });
        return json;
      }
    };

    if (!all) {
      const data = await doFetch(url);
      if (method === 'GET') cache.set(ck, data);
      return data;
    }
    // pagination
    let next = url;
    const items = [];
    while (next) {
      const page = await doFetch(next);
      if (Array.isArray(page?.value)) items.push(...page.value);
      next = page?.['@odata.nextLink'] || null;
    }
    if (method === 'GET') cache.set(ck, items);
    return items;
  }

  async function graphAll(relativeUri, opts = {}) {
    return graph(relativeUri, { ...opts, all: true });
  }

  // Resolve nested property paths like "defaultUserRolePermissions.allowedToCreateApps".
  function pickPath(obj, path) {
    if (obj == null || !path) return undefined;
    const parts = path.split('.');
    let cur = obj;
    for (const p of parts) {
      if (cur == null) return undefined;
      // primitive cast: arrays without index get returned as-is
      cur = cur[p];
    }
    return cur;
  }

  window.Graph = { graph, graphAll, clearCache, pickPath, beginCapture, endCapture };
  // Shared per-test API call recorder so other clients (e.g. Exchange InvokeCommand)
  // can attach their calls to the same drawer entry as Graph calls.
  window.ApiRecorder = {
    record: _record,
    shorten: _shorten,
  };
})();
