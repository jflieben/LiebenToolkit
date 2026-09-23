// Microsoft Graph client for InactiveDevices.
(() => {
  const BASE_V1 = 'https://graph.microsoft.com/v1.0';
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

  async function listAllDevices(onProgress) {
    const select = [
      'id', 'deviceId', 'displayName', 'accountEnabled', 'createdDateTime',
      'registrationDateTime', 'approximateLastSignInDateTime', 'operatingSystem',
      'operatingSystemVersion', 'trustType', 'profileType', 'isManaged',
      'isCompliant', 'onPremisesSyncEnabled', 'devicePhysicalIds', 'extensionAttributes',
    ].join(',');

    let next = `${BASE_BETA}/devices?$select=${select}&$top=999`;
    let pageNo = 0;
    const out = [];

    while (next) {
      const data = await call(next, { headers: { ConsistencyLevel: 'eventual' } });
      if (data && Array.isArray(data.value)) out.push(...data.value);
      pageNo++;
      if (onProgress) onProgress(out.length, pageNo);
      next = data && data['@odata.nextLink'];
    }

    return out;
  }

  async function listManagedDevices(onProgress) {
    const select = [
      'id', 'azureADDeviceId', 'deviceName', 'userPrincipalName',
      'lastSyncDateTime', 'managementAgent', 'managedDeviceOwnerType',
      'operatingSystem', 'enrolledDateTime', 'complianceState', 'autopilotEnrolled',
    ].join(',');

    let next = `${BASE_BETA}/deviceManagement/managedDevices?$select=${select}&$top=999`;
    let pageNo = 0;
    const out = [];

    while (next) {
      const data = await call(next, { headers: { ConsistencyLevel: 'eventual' } });
      if (data && Array.isArray(data.value)) out.push(...data.value);
      pageNo++;
      if (onProgress) onProgress(out.length, pageNo);
      next = data && data['@odata.nextLink'];
    }

    return out;
  }

  async function tryListManagedDevices(onProgress) {
    try {
      const list = await listManagedDevices(onProgress);
      return { list, unavailable: false, message: '' };
    } catch (e) {
      const known = e && (e.status === 401 || e.status === 403 || e.status === 404);
      if (known) {
        return {
          list: [],
          unavailable: true,
          message: 'Intune enrichment is unavailable (missing permission, missing Intune license, or no Intune setup). Device listing still works.',
        };
      }
      throw e;
    }
  }

  function disableDevice(id, extProps = null) {
    const body = { accountEnabled: false };
    if (extProps && Object.keys(extProps).length) body.extensionAttributes = extProps;
    return call(`${BASE_V1}/devices/${id}`, { method: 'PATCH', body, write: true });
  }

  function enableDevice(id) {
    return call(`${BASE_V1}/devices/${id}`, { method: 'PATCH', body: { accountEnabled: true }, write: true });
  }

  function deleteDevice(id) {
    return call(`${BASE_V1}/devices/${id}`, { method: 'DELETE', write: true });
  }

  window.Graph = {
    call,
    listAllDevices,
    tryListManagedDevices,
    disableDevice,
    enableDevice,
    deleteDevice,
  };
})();
