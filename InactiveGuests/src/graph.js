// Microsoft Graph client for InactiveGuests.
// Uses /beta/users for guest enumeration because that's where signInActivity
// lives reliably, and v1.0 for everything else.
(() => {
  const BASE_V1   = 'https://graph.microsoft.com/v1.0';
  const BASE_BETA = 'https://graph.microsoft.com/beta';

  async function callRaw(url, { method = 'GET', body, headers = {}, attempt = 0, write = false } = {}) {
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
    Log.dbg(`GRAPH ${method} ${url}`);
    const res = await fetch(url, opts);
    if (res.status === 429 || res.status === 503 || res.status === 504) {
      const ra = parseInt(res.headers.get('Retry-After') || '0', 10);
      const wait = (ra > 0 ? ra : Math.min(60, Math.pow(2, attempt))) * 1000;
      if (attempt >= 5) throw new Error(`Graph ${res.status} after retries: ${url}`);
      Log.warn(`Graph ${res.status}, backing off ${wait}ms (attempt ${attempt + 1})`);
      await Concurrency.sleep(wait);
      return callRaw(url, { method, body, headers, attempt: attempt + 1, write });
    }
    return res;
  }

  async function call(url, opts = {}) {
    const res = await callRaw(url, opts);
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      const e = new Error(`Graph ${res.status} ${res.statusText}: ${txt.substring(0, 400)}`);
      e.status = res.status;
      throw e;
    }
    if (res.status === 204) return null;
    const ct = res.headers.get('Content-Type') || '';
    if (ct.includes('json')) return res.json();
    return res.text();
  }

  /**
   * Enumerate every guest user in the tenant including signInActivity.
   * onProgress(loadedCount, pageNumber).
   */
  async function listAllGuests(onProgress) {
    const select = [
      'id','userPrincipalName','displayName','mail','userType',
      'externalUserState','externalUserStateChangeDateTime',
      'createdDateTime','creationType','accountEnabled','companyName',
      'signInActivity',
    ].join(',');
    const url0 = `${BASE_BETA}/users?$filter=userType eq 'Guest'&$select=${select}&$top=999`;
    const out = [];
    let pageNo = 0;
    let next = url0;
    while (next) {
      const data = await call(next, { headers: { ConsistencyLevel: 'eventual' } });
      if (data && Array.isArray(data.value)) out.push(...data.value);
      pageNo++;
      if (onProgress) onProgress(out.length, pageNo);
      next = data && data['@odata.nextLink'];
    }
    return out;
  }

  /**
   * List recently deleted users (Recycle Bin). Entra keeps deleted users for
   * 30 days; this is the directory equivalent of /users with userType filter.
   */
  async function listDeletedGuests(onProgress) {
    // /directory/deletedItems/microsoft.graph.user lets us $filter on userType.
    // deletedDateTime isn't part of the default projection, must be selected explicitly.
    const select = 'id,displayName,userPrincipalName,mail,userType,accountEnabled,deletedDateTime';
    const url0 = `${BASE_V1}/directory/deletedItems/microsoft.graph.user?$filter=userType eq 'Guest'&$select=${select}&$top=999`;
    const out = [];
    let pageNo = 0;
    let next = url0;
    while (next) {
      const data = await call(next, { headers: { ConsistencyLevel: 'eventual' } });
      if (data && Array.isArray(data.value)) out.push(...data.value);
      pageNo++;
      if (onProgress) onProgress(out.length, pageNo);
      next = data && data['@odata.nextLink'];
    }
    return out;
  }

  function disableUser(id) {
    return call(`${BASE_V1}/users/${id}`, { method: 'PATCH', body: { accountEnabled: false }, write: true });
  }
  function enableUser(id) {
    return call(`${BASE_V1}/users/${id}`, { method: 'PATCH', body: { accountEnabled: true }, write: true });
  }
  function deleteUser(id) {
    return call(`${BASE_V1}/users/${id}`, { method: 'DELETE', write: true });
  }
  /** Restore a deleted directory object (user) within the 30-day soft-delete window. */
  function restoreUser(id) {
    return call(`${BASE_V1}/directory/deletedItems/${id}/restore`, { method: 'POST', write: true });
  }

  window.Graph = {
    call,
    listAllGuests, listDeletedGuests,
    disableUser, enableUser, deleteUser, restoreUser,
  };
})();
