(() => {
  let traceCollector = null;
  const GRAPH_REQUEST_TIMEOUT_MS = 30000;

  function truncatePreview(value, limit) {
    const max = limit || 2400;
    if (value == null) return '';
    let text = '';
    try {
      text = typeof value === 'string' ? value : JSON.stringify(value, null, 2);
    } catch {
      text = `${value}`;
    }
    if (text.length <= max) return text;
    return `${text.slice(0, max)}\n...truncated...`;
  }

  function recordTrace(entry) {
    if (!traceCollector) return;
    traceCollector.push(entry);
  }

  function buildUrl(path, options) {
    const beta = !!options.beta;
    let url = `https://graph.microsoft.com/${beta ? 'beta' : 'v1.0'}/${path.replace(/^\//, '')}`;
    const query = options.query || null;
    if (query && typeof query === 'object') {
      const qs = new URLSearchParams();
      for (const [k, v] of Object.entries(query)) {
        if (v !== undefined && v !== null && `${v}` !== '') qs.set(k, `${v}`);
      }
      const serialized = qs.toString();
      if (serialized) url += `?${serialized}`;
    }
    return url;
  }

  function parseGraphErrorMessage(rawText) {
    if (!rawText) return '';
    try {
      const parsed = JSON.parse(rawText);
      return `${parsed?.error?.message || ''}`;
    } catch {
      return `${rawText}`;
    }
  }

  function dedupeUrls(urls) {
    const seen = new Set();
    const out = [];
    for (const url of urls || []) {
      if (!url || seen.has(url)) continue;
      seen.add(url);
      out.push(url);
    }
    return out;
  }

  function removeQueryParam(url, key) {
    try {
      const parsed = new URL(url);
      parsed.searchParams.delete(key);
      return parsed.toString();
    } catch {
      return url;
    }
  }

  function build400FallbackUrls(url, rawText) {
    const message = parseGraphErrorMessage(rawText).toLowerCase();
    const candidates = [];

    const supportsTopRemoval =
      message.includes('top is not allowed') ||
      message.includes('custom page sizes');
    const supportsSelectRemoval =
      message.includes("query option 'select' is not allowed") ||
      message.includes('could not find a property named') ||
      message.includes('select and expand failed');
    const supportsFilterRemoval =
      message.includes('filtered searches against this resource are not supported') ||
      message.includes("query option 'filter' is not allowed");
    const supportsExpandRemoval =
      message.includes("query option 'expand' is not allowed") ||
      message.includes('select and expand failed');

    let current = url;
    if (supportsTopRemoval) {
      current = removeQueryParam(current, '$top');
      candidates.push(current);
    }
    if (supportsSelectRemoval) {
      current = removeQueryParam(current, '$select');
      candidates.push(current);
    }
    if (supportsExpandRemoval) {
      current = removeQueryParam(current, '$expand');
      candidates.push(current);
    }
    if (supportsFilterRemoval) {
      current = removeQueryParam(current, '$filter');
      candidates.push(current);
    }

    // Generic safety fallbacks when Graph doesn't give a precise reason.
    candidates.push(removeQueryParam(url, '$top'));
    candidates.push(removeQueryParam(url, '$select'));
    candidates.push(removeQueryParam(removeQueryParam(url, '$top'), '$select'));
    candidates.push(removeQueryParam(removeQueryParam(removeQueryParam(url, '$top'), '$select'), '$filter'));
    candidates.push(removeQueryParam(removeQueryParam(removeQueryParam(removeQueryParam(url, '$top'), '$select'), '$filter'), '$expand'));

    return dedupeUrls(candidates).filter(candidate => candidate && candidate !== url);
  }

  // Graph and ARM calls go through the shared Http client (retry, backoff, Retry-After).
  // A deadline covers the call including its retries; the runner's per-control timeout
  // is the outer bound.
  function deadline(timeoutMs) {
    const controller = new AbortController();
    let fired = false;
    const timer = setTimeout(() => { fired = true; controller.abort(); }, timeoutMs * 3);
    return { signal: controller.signal, timedOut: () => fired, done: () => clearTimeout(timer) };
  }

  async function send(label, url, { method, body, headers, token, timeoutMs }) {
    const startedAt = performance.now();
    const limit = deadline(timeoutMs);
    const traceBase = () => ({ method, url, durationMs: Math.round(performance.now() - startedAt), requestPreview: truncatePreview(body, 1200) });
    let response;
    try {
      response = await Http.request(url, { method, body, headers, token, retries: 3, signal: limit.signal, raw: true, label });
    } catch (error) {
      if (limit.timedOut()) {
        const timeoutError = new Error(`${label} request timed out after ${Math.round((timeoutMs * 3) / 1000)}s: ${url}`);
        timeoutError.status = 408;
        recordTrace({ ...traceBase(), status: 408, ok: false, responsePreview: timeoutError.message, nextLink: null });
        throw timeoutError;
      }
      recordTrace({ ...traceBase(), status: 0, ok: false, responsePreview: truncatePreview(error.message || `${error}`, 3000), nextLink: null });
      throw error;
    } finally {
      limit.done();
    }
    const status = response.status;
    const rawText = status === 204 ? '' : await response.text();
    let parsed = null;
    if (rawText) { try { parsed = JSON.parse(rawText); } catch { parsed = rawText; } }
    recordTrace({ ...traceBase(), status, ok: response.ok, responsePreview: truncatePreview(parsed ?? rawText, 3000), nextLink: parsed?.['@odata.nextLink'] || parsed?.nextLink || null });
    return { response, status, rawText, parsed };
  }

  // Tokens are silent during a run: consent is settled before the run starts, and a
  // popup or redirect halfway through a scan would lose it.
  const graphToken = () => Auth.getToken('read', { interactive: false });

  async function fetchJson(url, options) {
    const method = options.method || 'GET';
    const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : GRAPH_REQUEST_TIMEOUT_MS;
    const { response, status, rawText, parsed } = await send('Graph', url, { method, body: options.body, headers: options.headers, token: graphToken, timeoutMs });

    if (!response.ok) {
      if (method === 'GET' && status === 400 && (options._fallbackDepth || 0) < 3) {
        for (const candidate of build400FallbackUrls(url, rawText)) {
          try {
            return await fetchJson(candidate, { ...options, _fallbackDepth: (options._fallbackDepth || 0) + 1 });
          } catch (retryErr) {
            if (retryErr?.status !== 400) throw retryErr;
          }
        }
      }
      const err = new Error(`Graph ${status}: ${rawText}`);
      err.status = status;
      throw err;
    }
    return status === 204 ? null : parsed;
  }

  const ARM_ROOT = 'https://management.azure.com';
  const ARM_SCOPE = 'https://management.azure.com/user_impersonation';

  // Azure Resource Manager fetch with the 'arm' scope set. Throws a tagged error
  // (armConsent) when no ARM token can be had silently, so callers can skip cleanly.
  async function armFetch(url, options) {
    const opts = options || {};
    const method = opts.method || 'GET';
    const timeoutMs = Number(opts.timeoutMs) > 0 ? Number(opts.timeoutMs) : GRAPH_REQUEST_TIMEOUT_MS;
    let token;
    try {
      token = await Auth.getToken('arm', { interactive: false });
    } catch {
      const e = new Error('Azure Resource Manager consent is required.');
      e.status = 401;
      e.armConsent = true;
      throw e;
    }
    const { response, status, rawText, parsed } = await send('ARM', url, { method, body: opts.body, token, timeoutMs });
    if (!response.ok) { const err = new Error(`ARM ${status}: ${rawText}`); err.status = status; throw err; }
    return parsed;
  }

  async function arm(path, opts) {
    const url = /^https?:/i.test(path) ? path : `${ARM_ROOT}/${path.replace(/^\//, '')}`;
    return armFetch(url, opts || {});
  }

  // Runs an Azure Resource Graph (KQL) query across all accessible subscriptions,
  // paging via $skipToken. Returns an array of row objects (objectArray format).
  async function armResourceGraph(query, opts) {
    const options = opts || {};
    const subsResp = await arm('subscriptions?api-version=2020-01-01');
    const subscriptions = (Array.isArray(subsResp?.value) ? subsResp.value : []).map(s => s.subscriptionId).filter(Boolean);
    if (!subscriptions.length) return [];
    const rows = [];
    let skipToken = null;
    do {
      const requestOptions = { $top: 1000 };
      if (skipToken) requestOptions.$skipToken = skipToken;
      const resp = await arm('providers/Microsoft.ResourceGraph/resources?api-version=2021-03-01', {
        method: 'POST',
        body: { subscriptions, query, options: requestOptions },
        timeoutMs: options.timeoutMs,
      });
      if (Array.isArray(resp?.data)) rows.push(...resp.data);
      skipToken = resp?.$skipToken || null;
    } while (skipToken);
    return rows;
  }

  async function graph(path, opts) {
    const options = opts || {};
    return fetchJson(buildUrl(path, options), options);
  }

  async function graphAll(path, opts) {
    const options = opts || {};
    const items = [];
    let nextUrl = buildUrl(path, options);

    while (nextUrl) {
      const payload = await fetchJson(nextUrl, {
        method: options.method || 'GET',
        timeoutMs: options.timeoutMs,
        headers: options.headers,
      });
      if (Array.isArray(payload?.value)) items.push(...payload.value);
      nextUrl = payload?.['@odata.nextLink'] || null;
    }

    return items;
  }

  // Pages a Graph collection with limits: stops after maxItems or maxMs and reports progress.
  // Returns { items, truncated }. Used by long-running collectors (sign-in logs, devices).
  async function graphPaged(path, opts) {
    const options = opts || {};
    const items = [];
    const maxItems = Number(options.maxItems) > 0 ? Number(options.maxItems) : Infinity;
    const maxMs = Number(options.maxMs) > 0 ? Number(options.maxMs) : Infinity;
    const startedAt = Date.now();
    let truncated = false;
    let nextUrl = buildUrl(path, options);

    while (nextUrl) {
      const payload = await fetchJson(nextUrl, {
        method: 'GET',
        timeoutMs: options.timeoutMs,
        headers: options.headers,
      });
      if (Array.isArray(payload?.value)) items.push(...payload.value);
      if (typeof options.onPage === 'function') options.onPage(items.length);
      nextUrl = payload?.['@odata.nextLink'] || null;
      if (nextUrl && (items.length >= maxItems || (Date.now() - startedAt) >= maxMs)) {
        truncated = true;
        break;
      }
    }

    return { items, truncated };
  }

  window.Api = {
    graph,
    graphAll,
    graphPaged,
    arm,
    armResourceGraph,
    ARM_SCOPE,
    setTraceCollector: (collector) => {
      traceCollector = Array.isArray(collector) ? collector : null;
    },
  };
})();
