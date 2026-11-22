// Graph and ARM client with a per-scan cache, like Maester's Invoke-MtGraphRequest:
// 20 EIDSCA tests that read /policies/authorizationPolicy cause one HTTP call. Runner
// clears the cache before every scan. Retry and backoff come from Http.request.
(() => {
  const cache = new Map();
  function clearCache() { cache.clear(); }

  // Per-test call recorder. Runner calls beginCapture() before test.run() and
  // endCapture() after; the calls are shown in the result drawer.
  const MAX_CALLS_PER_TEST = 8;
  const MAX_RESPONSE_CHARS = 1500;
  let capture = null;

  function beginCapture() { capture = { calls: [], overflow: 0 }; }
  function endCapture() {
    const c = capture;
    capture = null;
    if (!c) return [];
    if (c.overflow > 0) {
      c.calls.push({ method: 'INFO', url: `... and ${c.overflow} more call(s) not shown`, status: 0, durationMs: 0, fromCache: false, response: '', truncated: false });
    }
    return c.calls;
  }
  function record(entry) {
    if (!capture) return;
    if (capture.calls.length >= MAX_CALLS_PER_TEST) { capture.overflow++; return; }
    capture.calls.push(entry);
  }
  function shorten(data, max = MAX_RESPONSE_CHARS) {
    if (data == null) return { response: '', truncated: false };
    let str;
    try { str = typeof data === 'string' ? data : JSON.stringify(data); }
    catch { str = String(data); }
    if (str.length <= max) return { response: str, truncated: false };
    return { response: str.substring(0, max), truncated: true };
  }

  const ALLOWED_HOSTS = ['graph.microsoft.com', 'management.azure.com'];
  function assertHost(url) {
    const host = new URL(url).hostname.toLowerCase();
    if (!ALLOWED_HOSTS.includes(host)) throw new Error(`Refusing to send a token to ${host}`);
  }

  // GET (or other method) on Graph, or on ARM with { tokenScopes: 'arm' }.
  // opts.all follows @odata.nextLink / nextLink and returns the concatenated value arrays.
  async function graph(relativeUri, opts = {}) {
    const apiVersion = opts.apiVersion || 'v1.0';
    const method = (opts.method || 'GET').toUpperCase();
    const url = /^https:\/\//i.test(relativeUri)
      ? relativeUri
      : `https://graph.microsoft.com/${apiVersion}/${relativeUri.replace(/^\//, '')}`;
    assertHost(url);
    const label = /management\.azure\.com/i.test(url) ? 'ARM' : 'Graph';
    const tokenSet = opts.tokenScopes || 'read';

    const ck = `${method}|${url}|${opts.all ? 'all' : 'one'}`;
    if (method === 'GET' && cache.has(ck) && !opts.bypassCache) {
      const cached = cache.get(ck);
      record({ method, url, status: 200, durationMs: 0, fromCache: true, ...shorten(cached) });
      return cached;
    }

    const fetchOne = async (u) => {
      const t0 = performance.now();
      const res = await Http.request(u, {
        method,
        raw: true,
        label,
        token: () => Auth.getToken(tokenSet),
        headers: label === 'Graph' ? { ConsistencyLevel: opts.consistencyLevel || 'eventual' } : {},
        body: opts.body,
      });
      const durationMs = Math.round(performance.now() - t0);
      if (res.status === 204) { record({ method, url: u, status: 204, durationMs, fromCache: false, response: '', truncated: false }); return null; }
      const text = await res.text().catch(() => '');
      record({ method, url: u, status: res.status, durationMs, fromCache: false, ...shorten(text) });
      if (!res.ok) {
        const err = new Error(`${label} ${Http.errorMessage(res.status, res.statusText, text)} (${u})`);
        err.status = res.status;
        err.body = text;
        throw err;
      }
      try { return text ? JSON.parse(text) : null; } catch { return text; }
    };

    let data;
    if (!opts.all) {
      data = await fetchOne(url);
    } else {
      data = [];
      let next = url;
      while (next) {
        const page = await fetchOne(next);
        if (Array.isArray(page && page.value)) data.push(...page.value);
        next = (page && (page['@odata.nextLink'] || page.nextLink)) || null;
        if (next) assertHost(next);
      }
    }
    if (method === 'GET') cache.set(ck, data);
    return data;
  }

  function graphAll(relativeUri, opts = {}) { return graph(relativeUri, { ...opts, all: true }); }
  function arm(url, opts = {}) { return graph(url, { ...opts, tokenScopes: 'arm' }); }

  // Resolve nested property paths like "defaultUserRolePermissions.allowedToCreateApps".
  function pickPath(obj, path) {
    if (obj == null || !path) return undefined;
    let cur = obj;
    for (const p of path.split('.')) {
      if (cur == null) return undefined;
      cur = cur[p];
    }
    return cur;
  }

  // DNS lookup over Cloudflare DNS-over-HTTPS (no token). Used by the SPF, DKIM and DMARC tests.
  async function dns(name, type = 'TXT') {
    if (!/^[A-Z]{1,10}$/.test(type)) throw new Error(`Unsupported DNS record type ${type}`);
    const url = `https://cloudflare-dns.com/dns-query?name=${encodeURIComponent(name)}&type=${type}`;
    const t0 = performance.now();
    const data = await Http.json(url, { headers: { Accept: 'application/dns-json' }, label: 'DNS', retries: 2 });
    record({ method: 'DNS', url, status: 200, durationMs: Math.round(performance.now() - t0), fromCache: false, ...shorten(data) });
    return ((data && data.Answer) || []).map((a) => String(a.data || '').replace(/"/g, ''));
  }

  window.Graph = { graph, graphAll, arm, dns, clearCache, pickPath, beginCapture, endCapture, shorten };
})();
