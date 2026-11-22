// Shared HTTP client for Microsoft APIs: bearer token, retry with backoff on 429/5xx
// and network errors, Retry-After, readable errors, paging and URL encoding helpers.
//
//   Http.json(url, { token: () => Auth.getReadToken(), method, body, headers, signal })
//   Http.pageAll(url, opts, (items, page) => {})     follows @odata.nextLink
//   Http.url`${base}/users/${id}`                    encodes every interpolation
//   Http.odataString(name)                           'O''Brien' for $filter
(() => {
  const RETRY_STATUS = new Set([429, 500, 502, 503, 504]);

  function sleep(ms, signal) {
    return new Promise((resolve, reject) => {
      const t = setTimeout(resolve, ms);
      if (signal) signal.addEventListener('abort', () => { clearTimeout(t); reject(new Error('Cancelled')); }, { once: true });
    });
  }

  function retryDelay(res, attempt) {
    const ra = res && res.headers ? res.headers.get('Retry-After') : null;
    if (ra) {
      const secs = Number(ra);
      if (Number.isFinite(secs)) return Math.min(120, Math.max(1, secs)) * 1000;
      const when = Date.parse(ra);
      if (!Number.isNaN(when)) return Math.min(120000, Math.max(1000, when - Date.now()));
    }
    return Math.min(60000, 1000 * 2 ** attempt) + Math.floor(Math.random() * 400);
  }

  // Graph, ARM and SharePoint all have their own error shapes; get one readable line out.
  function errorMessage(status, statusText, bodyText) {
    let msg = '';
    try {
      const j = JSON.parse(bodyText);
      const e = j.error || j['odata.error'] || j;
      msg = (e && (e.message && (e.message.value || e.message))) || j.error_description || j.Message || '';
      const code = e && e.code ? ` (${e.code})` : '';
      if (msg) msg = `${msg}${code}`;
    } catch { /* not JSON */ }
    if (!msg) msg = (bodyText || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().substring(0, 300);
    return `HTTP ${status}${statusText ? ' ' + statusText : ''}${msg ? ': ' + msg : ''}`;
  }

  class HttpError extends Error {
    constructor(status, message, body, url) {
      super(message);
      this.name = 'HttpError';
      this.status = status;
      this.body = body;
      this.url = url;
      try { const j = JSON.parse(body); this.code = (j.error && j.error.code) || (j['odata.error'] && j['odata.error'].code) || undefined; } catch { /* not JSON */ }
    }
  }

  // timeout: milliseconds per attempt; a timed-out attempt is retried like a network error.
  async function request(url, { method = 'GET', body, headers = {}, token, retries = 5, signal, raw = false, label, timeout } = {}) {
    const name = label || 'HTTP';
    for (let attempt = 0; ; attempt++) {
      const h = { Accept: 'application/json', ...headers };
      if (token) h.Authorization = `Bearer ${typeof token === 'function' ? await token() : token}`;
      let payload = body;
      if (body != null && typeof body === 'object' && !(body instanceof Blob) && !(body instanceof ArrayBuffer) && !(body instanceof FormData) && !ArrayBuffer.isView(body)) {
        payload = JSON.stringify(body);
        if (!Object.keys(h).some((k) => k.toLowerCase() === 'content-type')) h['Content-Type'] = 'application/json';
      }
      Log.dbg(`${name} ${method} ${url}${attempt ? ` (retry ${attempt})` : ''}`);
      let res;
      const attemptAc = timeout ? new AbortController() : null;
      const timer = attemptAc ? setTimeout(() => attemptAc.abort(), timeout) : null;
      const onOuterAbort = () => attemptAc && attemptAc.abort();
      if (attemptAc && signal) signal.addEventListener('abort', onOuterAbort, { once: true });
      try {
        res = await fetch(url, { method, headers: h, body: payload, signal: attemptAc ? attemptAc.signal : signal });
      } catch (e) {
        if (signal && signal.aborted) throw new Error('Cancelled');
        if (e.name === 'AbortError' && !(attemptAc && attemptAc.signal.aborted)) throw new Error('Cancelled');
        if (attemptAc && attemptAc.signal.aborted) e = new Error(`timed out after ${Math.round(timeout / 1000)}s`);
        if (attempt >= retries) throw new Error(`${name} network error on ${method} ${url}: ${e.message}`);
        const wait = retryDelay(null, attempt);
        Log.warn(`${name} network error, retrying in ${Math.round(wait / 1000)}s: ${e.message}`);
        await sleep(wait, signal);
        continue;
      } finally {
        if (timer) clearTimeout(timer);
        if (attemptAc && signal) signal.removeEventListener('abort', onOuterAbort);
      }
      if (RETRY_STATUS.has(res.status) && attempt < retries) {
        const wait = retryDelay(res, attempt);
        Log.warn(`${name} ${res.status} on ${method}, backing off ${Math.round(wait / 1000)}s (attempt ${attempt + 1}/${retries})`);
        await sleep(wait, signal);
        continue;
      }
      if (raw) return res;
      if (!res.ok) {
        const txt = await res.text().catch(() => '');
        throw new HttpError(res.status, `${name} ${method}: ${errorMessage(res.status, res.statusText, txt)}`, txt, url);
      }
      return res;
    }
  }

  async function json(url, opts = {}) {
    const res = await request(url, opts);
    if (opts.raw) return res;
    if (res.status === 204 || res.status === 202) return null;
    const ct = res.headers.get('Content-Type') || '';
    if (ct.includes('json')) return res.json();
    const txt = await res.text();
    try { return JSON.parse(txt); } catch { return txt; }
  }

  // Follows @odata.nextLink (Graph, Dataverse) or nextLink (ARM). A next link on another
  // host is refused: the token for this host must not be sent anywhere else.
  async function pageAll(url, opts = {}, onPage) {
    const { nextKey, ...reqOpts } = opts;
    const host = new URL(url).host;
    const out = [];
    let next = url;
    let page = 0;
    while (next) {
      if (new URL(next, url).host !== host) throw new Error(`Refusing to follow a next page link to another host: ${new URL(next, url).host}`);
      const data = await json(next, reqOpts);
      const items = data && Array.isArray(data.value) ? data.value : [];
      out.push(...items);
      page++;
      if (onPage) onPage(items, page, out.length);
      next = data && (nextKey ? data[nextKey] : (data['@odata.nextLink'] || data.nextLink));
      if (opts.signal && opts.signal.aborted) throw new Error('Cancelled');
    }
    return out;
  }

  // Tagged template: every interpolated value is URI-component encoded, except
  // values wrapped with Http.raw() (for a trusted base URL).
  class Raw { constructor(v) { this.v = v; } }
  function rawPart(v) { return new Raw(String(v)); }
  function url(strings, ...values) {
    let out = strings[0];
    values.forEach((v, i) => {
      out += v instanceof Raw ? v.v : encodeURIComponent(String(v));
      out += strings[i + 1];
    });
    return out;
  }

  // Quote a value for an OData $filter string literal.
  function odataString(value) { return `'${String(value).replace(/'/g, "''")}'`; }

  // A GUID check before putting an id from a file or the URL into a path.
  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  function isGuid(v) { return GUID.test(String(v || '')); }

  window.Http = { request, json, pageAll, url, raw: rawPart, odataString, isGuid, HttpError, errorMessage, sleep };
})();
