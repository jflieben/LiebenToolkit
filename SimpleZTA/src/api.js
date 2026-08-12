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

  const THROTTLE_STATUS = new Set([429, 503, 504]);
  const MAX_THROTTLE_RETRIES = 3;

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function throttleWaitMs(response, attempt) {
    const retryAfter = Number(response?.headers?.get?.('Retry-After'));
    if (Number.isFinite(retryAfter) && retryAfter > 0) return Math.min(retryAfter, 120) * 1000;
    return Math.min(2 ** attempt, 30) * 1000;
  }

  async function fetchJson(url, options) {
    const method = options.method || 'GET';
    const scopes = options.scopes || null;
    const body = options.body;
    const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : GRAPH_REQUEST_TIMEOUT_MS;
    const token = await Auth.getToken(scopes);
    if (!token) throw new Error('Interactive consent started. Retry after sign-in redirect.');

    const startedAt = performance.now();
    let status = 0;
    let rawText = '';

    try {
      let response;
      // All calls in this app are reads, so retrying throttled responses is always safe.
      for (let attempt = 1; ; attempt += 1) {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), timeoutMs);
        try {
          response = await fetch(url, {
            method,
            headers: {
              Authorization: `Bearer ${token}`,
              'Content-Type': 'application/json',
              ...(options.headers || {}),
            },
            body: body ? JSON.stringify(body) : undefined,
            signal: controller.signal,
          });
        } finally {
          clearTimeout(timeout);
        }
        if (!THROTTLE_STATUS.has(response.status) || attempt > MAX_THROTTLE_RETRIES) break;
        const waitMs = throttleWaitMs(response, attempt);
        recordTrace({
          method,
          url,
          status: response.status,
          ok: false,
          durationMs: Math.round(performance.now() - startedAt),
          requestPreview: truncatePreview(body, 1200),
          responsePreview: `Throttled (attempt ${attempt}/${MAX_THROTTLE_RETRIES}); retrying in ${Math.round(waitMs / 1000)}s`,
          nextLink: null,
        });
        await sleep(waitMs);
      }

      status = response.status;
      rawText = response.status === 204 ? '' : await response.text();

      let parsed = null;
      if (rawText) {
        try {
          parsed = JSON.parse(rawText);
        } catch {
          parsed = rawText;
        }
      }

      recordTrace({
        method,
        url,
        status,
        ok: response.ok,
        durationMs: Math.round(performance.now() - startedAt),
        requestPreview: truncatePreview(body, 1200),
        responsePreview: truncatePreview(parsed ?? rawText, 3000),
        nextLink: parsed?.['@odata.nextLink'] || null,
      });

      if (!response.ok) {
        if (
          method === 'GET' &&
          status === 400 &&
          (options._fallbackDepth || 0) < 3
        ) {
          const candidates = build400FallbackUrls(url, rawText);
          for (const candidate of candidates) {
            try {
              return await fetchJson(candidate, {
                ...options,
                _fallbackDepth: (options._fallbackDepth || 0) + 1,
              });
            } catch (retryErr) {
              if (retryErr?.status !== 400) throw retryErr;
            }
          }
        }

        const err = new Error(`Graph ${status}: ${rawText}`);
        err.status = status;
        throw err;
      }

      if (response.status === 204) return null;
      return parsed;
    } catch (error) {
      if (error && (error.name === 'AbortError' || /aborted|timeout/i.test(`${error.message || ''}`))) {
        const timeoutError = new Error(`Graph request timed out after ${timeoutMs}ms: ${url}`);
        timeoutError.status = 408;
        recordTrace({
          method,
          url,
          status: 408,
          ok: false,
          durationMs: Math.round(performance.now() - startedAt),
          requestPreview: truncatePreview(body, 1200),
          responsePreview: truncatePreview(timeoutError.message, 3000),
          nextLink: null,
        });
        throw timeoutError;
      }
      recordTrace({
        method,
        url,
        status,
        ok: false,
        durationMs: Math.round(performance.now() - startedAt),
        requestPreview: truncatePreview(body, 1200),
        responsePreview: truncatePreview(rawText || error.message || `${error}`, 3000),
        nextLink: null,
      });
      throw error;
    }
  }

  const ARM_ROOT = 'https://management.azure.com';
  const ARM_SCOPE = 'https://management.azure.com/user_impersonation';

  // Azure Resource Manager fetch. Uses a delegated ARM token; throws a tagged error
  // (armConsent) when the ARM scope has not been consented so callers can skip cleanly.
  async function armFetch(url, options) {
    const opts = options || {};
    const method = opts.method || 'GET';
    const body = opts.body;
    const timeoutMs = Number(opts.timeoutMs) > 0 ? Number(opts.timeoutMs) : GRAPH_REQUEST_TIMEOUT_MS;
    const token = await Auth.getTokenSilent([ARM_SCOPE]);
    if (!token) { const e = new Error('Azure Resource Manager consent is required.'); e.status = 401; e.armConsent = true; throw e; }

    const startedAt = performance.now();
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    let response, rawText = '', status = 0;
    try {
      response = await fetch(url, {
        method,
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: body ? JSON.stringify(body) : undefined,
        signal: controller.signal,
      });
      status = response.status;
      rawText = response.status === 204 ? '' : await response.text();
    } catch (error) {
      recordTrace({ method, url, status: 0, ok: false, durationMs: Math.round(performance.now() - startedAt), requestPreview: truncatePreview(body, 1200), responsePreview: truncatePreview(error.message || `${error}`, 3000), nextLink: null });
      if (error && (error.name === 'AbortError' || /aborted|timeout/i.test(`${error.message || ''}`))) { const e = new Error(`ARM request timed out after ${timeoutMs}ms`); e.status = 408; throw e; }
      throw error;
    } finally {
      clearTimeout(timeout);
    }

    let parsed = null;
    if (rawText) { try { parsed = JSON.parse(rawText); } catch { parsed = rawText; } }
    recordTrace({ method, url, status, ok: response.ok, durationMs: Math.round(performance.now() - startedAt), requestPreview: truncatePreview(body, 1200), responsePreview: truncatePreview(parsed ?? rawText, 3000), nextLink: parsed?.['@odata.nextLink'] || parsed?.nextLink || null });
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
        scopes: options.scopes || null,
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
        scopes: options.scopes || null,
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
