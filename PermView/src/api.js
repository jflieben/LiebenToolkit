// Multi-API HTTP client for PermView. Each backend (Graph, ARM, PowerBI,
// Azure DevOps, Power Platform) gets its own helper that knows how to fetch
// its token, deal with throttling and follow that API's pagination scheme.
(() => {
  async function _fetch(url, { method = 'GET', body, headers = {}, getToken, attempt = 0 } = {}) {
    const token = await getToken();
    const opts = {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        ...(body ? { 'Content-Type': 'application/json' } : {}),
        ...headers,
      },
    };
    if (body) opts.body = typeof body === 'string' ? body : JSON.stringify(body);
    Log.dbg(`${method} ${url}`);
    const res = await fetch(url, opts);
    if (res.status === 429 || res.status === 503 || res.status === 504) {
      const ra = parseInt(res.headers.get('Retry-After') || '0', 10);
      const wait = (ra > 0 ? ra : Math.min(60, Math.pow(2, attempt))) * 1000;
      if (attempt >= 5) throw _err(res, url, await res.text().catch(() => ''));
      Log.warn(`${res.status}, backing off ${wait}ms (attempt ${attempt + 1})`);
      await Concurrency.sleep(wait);
      return _fetch(url, { method, body, headers, getToken, attempt: attempt + 1 });
    }
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      throw _err(res, url, txt);
    }
    if (res.status === 204) return null;
    const ct = res.headers.get('Content-Type') || '';
    if (ct.includes('json')) return res.json();
    return res.text();
  }

  function _err(res, url, txt) {
    const e = new Error(`${res.status} ${res.statusText}: ${(txt || '').substring(0, 400)}`);
    e.status = res.status;
    e.url = url;
    return e;
  }

  // ---------- Microsoft Graph ----------
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const Graph = {
    call(path, opts = {}) {
      const url = path.startsWith('http') ? path : GRAPH + path;
      return _fetch(url, { ...opts, getToken: Auth.getGraphToken });
    },
    async pageAll(path, opts = {}) {
      const out = [];
      let next = path.startsWith('http') ? path : GRAPH + path;
      while (next) {
        const data = await _fetch(next, { ...opts, getToken: Auth.getGraphToken });
        if (data && Array.isArray(data.value)) out.push(...data.value);
        next = data && data['@odata.nextLink'];
      }
      return out;
    },
  };

  // ---------- Azure Resource Manager ----------
  const ARM = 'https://management.azure.com';
  const Arm = {
    call(path, opts = {}) {
      const url = path.startsWith('http') ? path : ARM + path;
      return _fetch(url, { ...opts, getToken: Auth.getArmToken });
    },
    async pageAll(path, opts = {}) {
      const out = [];
      let next = path.startsWith('http') ? path : ARM + path;
      while (next) {
        const data = await _fetch(next, { ...opts, getToken: Auth.getArmToken });
        if (data && Array.isArray(data.value)) out.push(...data.value);
        next = data && data.nextLink;
      }
      return out;
    },
  };

  // ---------- Power BI REST ----------
  const PBI = 'https://api.powerbi.com/v1.0/myorg';
  const PowerBI = {
    call(path, opts = {}) {
      const url = path.startsWith('http') ? path : PBI + path;
      return _fetch(url, { ...opts, getToken: Auth.getPowerBIToken });
    },
  };

  // ---------- Azure DevOps ----------
  // VSSPS profile (find user) + accounts (orgs) live on app.vssps.visualstudio.com.
  // Project / group queries live on dev.azure.com/{org} or vssps.dev.azure.com/{org}.
  const DevOps = {
    call(url, opts = {}) {
      return _fetch(url, { ...opts, getToken: Auth.getDevOpsToken });
    },
  };

  // ---------- Power Platform (Flow / PowerApps / Environments) ----------
  // BAP API uses the Flow service token.
  const BAP = 'https://api.bap.microsoft.com';
  const PowerPlatform = {
    call(path, opts = {}) {
      const url = path.startsWith('http') ? path : BAP + path;
      return _fetch(url, { ...opts, getToken: Auth.getFlowToken });
    },
  };

  // ---------- Dataverse Web API ----------
  // Each environment has its own instance URL (e.g. https://orgname.crm4.dynamics.com).
  // Tokens are per-instance, so the caller passes the instance URL.
  const Dataverse = {
    call(instanceUrl, path, opts = {}) {
      const base = instanceUrl.replace(/\/$/, '');
      const url = path.startsWith('http') ? path : `${base}/api/data/v9.2${path}`;
      return _fetch(url, {
        ...opts,
        getToken: () => Auth.getDataverseToken(instanceUrl),
        headers: { 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', ...(opts.headers || {}) },
      });
    },
  };

  window.Api = { Graph, Arm, PowerBI, DevOps, PowerPlatform, Dataverse };
})();
