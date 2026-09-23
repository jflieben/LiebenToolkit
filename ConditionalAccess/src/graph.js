(() => {
  const V1 = 'https://graph.microsoft.com/v1.0';

  async function callRaw(url, { method = 'GET', body, headers = {}, write = false, attempt = 0 } = {}) {
    const token = write ? await Auth.getWriteToken() : await Auth.getReadToken();
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

    const res = await fetch(url, opts);
    if (res.status === 429 || res.status === 503 || res.status === 504) {
      if (attempt >= 5) throw new Error(`Graph ${res.status} after retries: ${url}`);
      const retryAfter = parseInt(res.headers.get('Retry-After') || '0', 10);
      const delayMs = (retryAfter > 0 ? retryAfter : Math.pow(2, attempt + 1)) * 1000;
      Log.warn(`Graph ${res.status} retry in ${delayMs}ms: ${url}`);
      await new Promise((r) => setTimeout(r, delayMs));
      return callRaw(url, { method, body, headers, write, attempt: attempt + 1 });
    }
    return res;
  }

  async function call(pathOrUrl, opts = {}) {
    const url = pathOrUrl.startsWith('http') ? pathOrUrl : `${V1}${pathOrUrl}`;
    const res = await callRaw(url, opts);
    if (!res.ok) {
      const text = await res.text().catch(() => '');
      const e = new Error(`Graph ${res.status} ${res.statusText}: ${text.substring(0, 500)}`);
      e.status = res.status;
      throw e;
    }
    if (res.status === 204) return null;
    const ct = (res.headers.get('Content-Type') || '').toLowerCase();
    return ct.includes('json') ? res.json() : res.text();
  }

  async function pageAll(path) {
    let next = `${V1}${path}`;
    const out = [];
    while (next) {
      const data = await call(next);
      if (data && Array.isArray(data.value)) out.push(...data.value);
      next = data && data['@odata.nextLink'];
    }
    return out;
  }

  async function listConditionalAccessPolicies() {
    return pageAll('/identity/conditionalAccess/policies?$top=200');
  }

  async function createConditionalAccessPolicy(policyBody) {
    return call('/identity/conditionalAccess/policies', {
      method: 'POST',
      body: policyBody,
      write: true,
    });
  }

  async function updateConditionalAccessPolicy(policyId, policyBody) {
    return call(`/identity/conditionalAccess/policies/${encodeURIComponent(policyId)}`, {
      method: 'PATCH',
      body: policyBody,
      write: true,
    });
  }

  async function getDirectoryObjectsByIds(ids) {
    const unique = Array.from(new Set((ids || []).filter(Boolean)));
    if (!unique.length) return [];
    const chunks = [];
    for (let i = 0; i < unique.length; i += 1000) chunks.push(unique.slice(i, i + 1000));
    const all = [];
    for (const chunk of chunks) {
      const data = await call('/directoryObjects/getByIds', {
        method: 'POST',
        body: {
          ids: chunk,
          types: ['user', 'group'],
        },
      });
      if (data && Array.isArray(data.value)) all.push(...data.value);
    }
    return all;
  }

  async function searchUsers(term) {
    const q = (term || '').trim().replace(/'/g, "''");
    if (!q) return [];
    const path = `/users?$select=id,displayName,userPrincipalName&$top=25&$filter=startswith(displayName,'${q}') or startswith(userPrincipalName,'${q}')`;
    const data = await call(path, { headers: { ConsistencyLevel: 'eventual' } });
    return (data && data.value) || [];
  }

  async function searchGroups(term) {
    const q = (term || '').trim().replace(/'/g, "''");
    if (!q) return [];
    const path = `/groups?$select=id,displayName&$top=25&$filter=startswith(displayName,'${q}')`;
    const data = await call(path, { headers: { ConsistencyLevel: 'eventual' } });
    return (data && data.value) || [];
  }

  window.Graph = {
    listConditionalAccessPolicies,
    createConditionalAccessPolicy,
    updateConditionalAccessPolicy,
    getDirectoryObjectsByIds,
    searchUsers,
    searchGroups,
  };
})();