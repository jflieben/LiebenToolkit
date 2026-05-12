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
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), timeoutMs);
      let response;
      try {
        response = await fetch(url, {
          method,
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: body ? JSON.stringify(body) : undefined,
          signal: controller.signal,
        });
      } finally {
        clearTimeout(timeout);
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
      });
      if (Array.isArray(payload?.value)) items.push(...payload.value);
      nextUrl = payload?.['@odata.nextLink'] || null;
    }

    return items;
  }

  window.Api = {
    graph,
    graphAll,
    setTraceCollector: (collector) => {
      traceCollector = Array.isArray(collector) ? collector : null;
    },
  };
})();
