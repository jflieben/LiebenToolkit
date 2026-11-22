// Microsoft Graph calls for SPOTrim. Read only. Tokens are taken silently: consent is
// asked before a run starts, so nothing opens a sign-in window halfway through.
(() => {
  const BASE = 'https://graph.microsoft.com/v1.0';
  const token = () => Auth.getToken('read', { interactive: false, noRedirect: true });

  const get = (path, opts = {}) => Http.json(path.startsWith('http') ? path : BASE + path, { ...opts, label: 'Graph', token });

  // SharePoint hosts of the signed-in tenant. /sites/root follows a tenant rename;
  // the initial onmicrosoft.com domain is the fallback.
  async function detectHosts() {
    try {
      const root = await get('/sites/root?$select=siteCollection,webUrl');
      const hosts = Sites.hostsFromRoot(root && root.siteCollection && root.siteCollection.hostname);
      if (hosts) return hosts;
      Log.warn(`Unexpected root site host: ${root && root.siteCollection && root.siteCollection.hostname}`);
    } catch (e) {
      Log.warn(`Root site lookup failed, falling back to the initial domain: ${e.message}`);
    }
    const org = await get('/organization?$select=verifiedDomains');
    const domains = (org && org.value && org.value[0] && org.value[0].verifiedDomains) || [];
    const initial = (domains.find((d) => d.isInitial) || {}).name || '';
    const hosts = Sites.hostsFromRoot(`${initial.split('.')[0]}.sharepoint.com`);
    if (!hosts) throw new Error(`Could not work out the SharePoint host from initial domain "${initial}"`);
    return hosts;
  }

  // Lowercase, no trailing slash, percent-decoded. For matching URLs.
  function normalizeUrl(u) {
    if (!u) return '';
    try { u = decodeURIComponent(u); } catch { /* keep as is */ }
    return u.toLowerCase().replace(/\/+$/, '').trim();
  }

  const isoDate = (d) => d.toISOString().substring(0, 10);

  // Engagement signals for one site in a single $batch:
  //   a  analytics/allTime, r  analytics/lastSevenDays,
  //   i  drive root getActivitiesByInterval (30 days, daily),
  //   p  most recently edited site page, q  drive quota.
  // 404 on a sub-request means "no data"; other errors land in _errors.
  async function getSiteEngagement(siteId) {
    if (!siteId) return null;
    const enc = encodeURIComponent(siteId);
    const now = new Date();
    const startStr = isoDate(new Date(now.getTime() - 30 * 86400000));
    const endStr = isoDate(now);
    const pagesUrl = `/sites/${enc}/pages/microsoft.graph.sitePage?$top=1&$orderby=lastModifiedDateTime%20desc&$select=name,title,webUrl,lastModifiedDateTime,lastModifiedBy`;

    const json = await Http.json(`${BASE}/$batch`, {
      method: 'POST', label: 'Graph', token,
      body: {
        requests: [
          { id: 'a', method: 'GET', url: `/sites/${enc}/analytics/allTime` },
          { id: 'r', method: 'GET', url: `/sites/${enc}/analytics/lastSevenDays` },
          { id: 'i', method: 'GET', url: `/sites/${enc}/drive/items/root/getActivitiesByInterval(startDateTime='${startStr}',endDateTime='${endStr}',interval='day')` },
          { id: 'p', method: 'GET', url: pagesUrl },
          { id: 'q', method: 'GET', url: `/sites/${enc}/drive?$select=quota` },
        ],
      },
    });
    return parseEngagement(json);
  }

  function parseEngagement(json) {
    const byId = {};
    for (const s of (json && json.responses) || []) byId[s.id] = s;
    const errs = [];
    function pick(id, notFound) {
      const s = byId[id];
      if (!s) return null;
      if (s.status === 404) return notFound;
      if (s.status >= 400) {
        const msg = (s.body && s.body.error && s.body.error.message) || JSON.stringify(s.body || '');
        errs.push(`${id}: HTTP ${s.status} ${String(msg).substring(0, 120)}`);
        return null;
      }
      return s.body || null;
    }
    const all = pick('a', { access: { actionCount: 0, actorCount: 0 } });
    const rec = pick('r', { access: { actionCount: 0, actorCount: 0 } });
    const interval = pick('i', null);
    const pages = pick('p', null);
    const drive = pick('q', null);

    let last30Views = 0, last30Actors = 0, last30LastDay = null;
    if (interval && Array.isArray(interval.value)) {
      for (const w of interval.value) {
        const ac = (w && w.access && w.access.actionCount) || 0;
        last30Views += ac;
        // Per window distinct, so this is a lower bound for the 30 days.
        last30Actors = Math.max(last30Actors, (w && w.access && w.access.actorCount) || 0);
        if (ac > 0 && w.endDateTime) {
          const end = new Date(w.endDateTime);
          if (!last30LastDay || end > last30LastDay) last30LastDay = end;
        }
      }
    }

    let lastPageEdit = null;
    const p = pages && Array.isArray(pages.value) ? pages.value[0] : null;
    if (p) {
      const by = (p.lastModifiedBy && p.lastModifiedBy.user) || {};
      lastPageEdit = { date: p.lastModifiedDateTime || null, name: p.title || p.name || '', webUrl: p.webUrl || '', modifiedBy: by.displayName || by.email || null };
    }

    const access = (x) => (x && x.access) || {};
    const result = {
      allTimeViews: access(all).actionCount ?? null,
      allTimeActors: access(all).actorCount ?? null,
      allTimeStart: (all && all.startDateTime) || null,
      allTimeEnd: (all && all.endDateTime) || null,
      recentViews: access(rec).actionCount ?? null,
      recentActors: access(rec).actorCount ?? null,
      recentStart: (rec && rec.startDateTime) || null,
      recentEnd: (rec && rec.endDateTime) || null,
      last30Views,
      last30Actors,
      last30LastDay: last30LastDay ? last30LastDay.toISOString() : null,
      lastPageEdit,
      quotaUsedBytes: (drive && drive.quota && drive.quota.used) ?? null,
    };
    if (errs.length) result._errors = errs;
    return result;
  }

  window.Graph = { detectHosts, normalizeUrl, getSiteEngagement, parseEngagement };
})();
