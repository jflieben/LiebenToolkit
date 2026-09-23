// Microsoft Graph client for DupedDevices.
// Handles retry/backoff, pagination and PATCH/DELETE on /devices.
(() => {
  const BASE = 'https://graph.microsoft.com/v1.0';
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

  async function call(path, opts = {}) {
    const url = path.startsWith('http') ? path : BASE + path;
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

  async function* pageAll(path, opts = {}) {
    let next = path.startsWith('http') ? path : BASE + path;
    while (next) {
      const data = await call(next, opts);
      if (data && Array.isArray(data.value)) for (const v of data.value) yield v;
      next = data && data['@odata.nextLink'];
    }
  }

  /**
   * List all devices in the tenant. Uses ConsistencyLevel=eventual + $count
   * so we can $top=999 (the maximum the directory endpoint supports) and avoid
   * server-side throttling on small page sizes.
   */
  async function listAllDevices(onProgress) {
    const select = [
      'id','deviceId','displayName','operatingSystem','operatingSystemVersion',
      'trustType','accountEnabled','approximateLastSignInDateTime','createdDateTime',
      'physicalIds','extensionAttributes','isManaged','isCompliant',
    ].join(',');
    const path = `/devices?$top=999&$select=${select}`;
    const out = [];
    let pageNo = 0;
    let next = BASE + path;
    while (next) {
      const data = await call(next, { headers: { ConsistencyLevel: 'eventual' } });
      if (data && Array.isArray(data.value)) {
        for (const v of data.value) out.push(v);
      }
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
          message: 'Intune enrichment is unavailable (missing permission/license or no Intune setup). Duplicate detection still works, but Autopilot recognition may be partial.',
        };
      }
      throw e;
    }
  }

  /** Disable a device by setting accountEnabled=false. Optionally PATCH an extensionAttribute. */
  async function disableDevice(deviceId, extProps = null) {
    const body = { accountEnabled: false };
    if (extProps && Object.keys(extProps).length) body.extensionAttributes = extProps;
    return call(`/devices/${deviceId}`, { method: 'PATCH', body, write: true });
  }

  /** Re-enable a device. */
  async function enableDevice(deviceId) {
    return call(`/devices/${deviceId}`, { method: 'PATCH', body: { accountEnabled: true }, write: true });
  }

  /** Just write extensionAttributes (no enable/disable change). */
  async function patchExtensionAttributes(deviceId, extProps) {
    return call(`/devices/${deviceId}`, { method: 'PATCH', body: { extensionAttributes: extProps }, write: true });
  }

  /** Delete a device permanently. */
  async function deleteDevice(deviceId) {
    return call(`/devices/${deviceId}`, { method: 'DELETE', write: true });
  }

  window.Graph = {
    listAllDevices,
    tryListManagedDevices,
    disableDevice, enableDevice, patchExtensionAttributes, deleteDevice,
    pageAll, call,
  };
})();
