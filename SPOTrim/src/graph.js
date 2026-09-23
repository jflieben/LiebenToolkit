// Microsoft Graph client with retry/backoff and pagination
(() => {
  const BASE = 'https://graph.microsoft.com/v1.0';

  async function callRaw(url, { method = 'GET', body, headers = {}, attempt = 0, accept = 'application/json' } = {}) {
    const token = await Auth.getGraphToken();
    const opts = {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: accept,
        ...(body ? { 'Content-Type': 'application/json' } : {}),
        ...headers,
      },
    };
    if (body) opts.body = typeof body === 'string' ? body : JSON.stringify(body);
    Log.dbg(`GRAPH ${method} ${url}`);
    const res = await fetch(url, opts);
    if (res.status === 429 || res.status === 503 || res.status === 504) {
      const ra = parseInt(res.headers.get('Retry-After') || '0', 10);
      const wait = (ra > 0 ? ra : Math.min(60, Math.pow(2, attempt))) * 1000;
      if (attempt >= 5) throw new Error(`Graph ${res.status} after retries: ${url}`);
      Log.warn(`Graph ${res.status} → backing off ${wait}ms (attempt ${attempt + 1})`);
      await Concurrency.sleep(wait);
      return callRaw(url, { method, body, headers, attempt: attempt + 1, accept });
    }
    return res;
  }

  async function call(path, opts = {}) {
    const url = path.startsWith('http') ? path : BASE + path;
    const res = await callRaw(url, opts);
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      const e = new Error(`Graph ${res.status} ${res.statusText}: ${txt.substring(0, 300)}`);
      e.status = res.status;
      throw e;
    }
    if (res.status === 204) return null;
    const ct = res.headers.get('Content-Type') || '';
    if (ct.includes('json')) return res.json();
    return res.text();
  }

  async function* pageAll(path, opts = {}) {
    let next = path.startsWith('http') ? path : BASE + path;
    while (next) {
      const data = await call(next, opts);
      if (data && Array.isArray(data.value)) {
        for (const v of data.value) yield v;
      }
      next = data && data['@odata.nextLink'];
    }
  }

  async function listAllSites() {
    // Use search=* to enumerate; works in delegated context with Sites.Read.All
    const out = [];
    const select = 'id,displayName,name,webUrl,createdDateTime,lastModifiedDateTime,siteCollection,root,isPersonalSite';
    for await (const s of pageAll(`/sites?search=*&$select=${select}&$top=200`)) {
      out.push(s);
    }
    return out;
  }

  /** Lowercase, strip trailing slash, decode percent-encoding for matching. */
  function normalizeUrl(u) {
    if (!u) return '';
    try { u = decodeURIComponent(u); } catch {}
    return u.toLowerCase().replace(/\/+$/, '').trim();
  }

  function isoDate(d) { return d.toISOString().substring(0, 10); }

  /**
   * Per-site engagement signals, all fetched in a single Graph $batch.
   * All endpoints live on graph.microsoft.com which has proper CORS, so this
   * works directly from a static-hosted SPA - no proxy required.
   *
   * Sub-requests (5):
   *   a - /sites/{id}/analytics/allTime
   *   r - /sites/{id}/analytics/lastSevenDays
   *   i - /sites/{id}/drive/items/root/getActivitiesByInterval (last 30 days, daily)
   *   p - /sites/{id}/pages/microsoft.graph.sitePage?$top=1&$orderby=lastModifiedDateTime desc
   *   q - /sites/{id}/drive?$select=quota
   *
   * Sub-responses that 404 are treated as "no data" (returns nulls/zeros).
   * Other 4xx/5xx on individual sub-responses are recorded in `_errors` but
   * don't fail the whole call.
   */
  async function getSiteEngagement(siteId) {
    if (!siteId) return null;
    const enc = encodeURIComponent(siteId);

    const now = new Date();
    const start30 = new Date(now.getTime() - 30 * 86400000);
    const startStr = isoDate(start30);
    const endStr = isoDate(now);

    const pagesUrl = `/sites/${enc}/pages/microsoft.graph.sitePage?$top=1&$orderby=lastModifiedDateTime%20desc&$select=name,title,webUrl,lastModifiedDateTime,lastModifiedBy`;

    const batchBody = {
      requests: [
        { id: 'a', method: 'GET', url: `/sites/${enc}/analytics/allTime` },
        { id: 'r', method: 'GET', url: `/sites/${enc}/analytics/lastSevenDays` },
        { id: 'i', method: 'GET', url: `/sites/${enc}/drive/items/root/getActivitiesByInterval(startDateTime='${startStr}',endDateTime='${endStr}',interval='day')` },
        { id: 'p', method: 'GET', url: pagesUrl },
        { id: 'q', method: 'GET', url: `/sites/${enc}/drive?$select=quota` },
      ],
    };
    const res = await callRaw(`${BASE}/$batch`, { method: 'POST', body: batchBody, accept: 'application/json' });
    if (!res.ok) {
      const t = await res.text().catch(() => '');
      const e = new Error(`engagement $batch HTTP ${res.status}: ${t.substring(0, 200)}`);
      e.status = res.status;
      throw e;
    }
    const json = await res.json();
    const subs = (json && json.responses) || [];
    const byId = {};
    for (const s of subs) byId[s.id] = s;

    const errs = [];
    function pickAnalytics(id) {
      const s = byId[id];
      if (!s) return null;
      if (s.status === 404) return { access: { actionCount: 0, actorCount: 0 } };
      if (s.status >= 400) {
        const msg = s.body?.error?.message || JSON.stringify(s.body || '');
        errs.push(`${id}: HTTP ${s.status} ${String(msg).substring(0, 120)}`);
        return null;
      }
      return s.body || null;
    }
    function pickGeneric(id) {
      const s = byId[id];
      if (!s) return null;
      if (s.status === 404) return null;
      if (s.status >= 400) {
        const msg = s.body?.error?.message || JSON.stringify(s.body || '');
        errs.push(`${id}: HTTP ${s.status} ${String(msg).substring(0, 120)}`);
        return null;
      }
      return s.body || null;
    }

    const all = pickAnalytics('a');
    const rec = pickAnalytics('r');
    const interval = pickGeneric('i');
    const pages = pickGeneric('p');
    const drive = pickGeneric('q');

    // Sum the daily ItemActivityStat windows for the last 30 days.
    let last30Views = 0, last30Actors = 0, last30LastDay = null;
    if (interval && Array.isArray(interval.value)) {
      for (const w of interval.value) {
        const ac = w?.access?.actionCount || 0;
        const at = w?.access?.actorCount  || 0;
        last30Views  += ac;
        last30Actors  = Math.max(last30Actors, at); // upper bound (per-window distinct)
        if (ac > 0) {
          const end = w?.endDateTime ? new Date(w.endDateTime) : null;
          if (end && (!last30LastDay || end > last30LastDay)) last30LastDay = end;
        }
      }
    }

    let lastPageEdit = null;
    if (pages && Array.isArray(pages.value) && pages.value[0]) {
      const p = pages.value[0];
      lastPageEdit = {
        date: p.lastModifiedDateTime || null,
        name: p.title || p.name || '',
        webUrl: p.webUrl || '',
        modifiedBy: p.lastModifiedBy?.user?.displayName || p.lastModifiedBy?.user?.email || null,
      };
    }

    const quotaUsedBytes = drive?.quota?.used ?? null;

    const result = {
      allTimeViews:  all?.access?.actionCount ?? null,
      allTimeActors: all?.access?.actorCount ?? null,
      allTimeStart:  all?.startDateTime || null,
      allTimeEnd:    all?.endDateTime || null,
      recentViews:   rec?.access?.actionCount ?? null,
      recentActors:  rec?.access?.actorCount ?? null,
      recentStart:   rec?.startDateTime || null,
      recentEnd:     rec?.endDateTime || null,
      last30Views,
      last30Actors,
      last30LastDay: last30LastDay ? last30LastDay.toISOString() : null,
      lastPageEdit,
      quotaUsedBytes,
    };
    if (errs.length) result._errors = errs;
    return result;
  }

  window.Graph = {
    call, pageAll, listAllSites, normalizeUrl,
    getSiteEngagement,
  };
})();
