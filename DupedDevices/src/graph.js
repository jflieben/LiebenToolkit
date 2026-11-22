// Microsoft Graph calls for DupedDevices.
(() => {
  const V1 = 'https://graph.microsoft.com/v1.0';
  const BETA = 'https://graph.microsoft.com/beta';

  const DEVICE_FIELDS = [
    'id', 'deviceId', 'displayName', 'operatingSystem', 'operatingSystemVersion',
    'trustType', 'accountEnabled', 'approximateLastSignInDateTime', 'createdDateTime',
    'physicalIds', 'extensionAttributes', 'isManaged', 'isCompliant',
  ].join(',');

  const MANAGED_FIELDS = [
    'id', 'azureADDeviceId', 'deviceName', 'userPrincipalName', 'lastSyncDateTime',
    'managementAgent', 'managedDeviceOwnerType', 'operatingSystem', 'enrolledDateTime',
    'complianceState', 'autopilotEnrolled',
  ].join(',');

  const read = { label: 'Graph', token: () => Auth.getReadToken() };
  const write = { label: 'Graph', token: () => Auth.getWriteToken() };

  function deviceUrl(id) {
    if (!Http.isGuid(id)) throw new Error(`Not a device object id: ${id}`);
    return Http.url`${Http.raw(V1)}/devices/${id}`;
  }

  // onProgress(loadedCount, pageNumber)
  function listAllDevices(onProgress, signal) {
    return Http.pageAll(`${V1}/devices?$top=999&$select=${DEVICE_FIELDS}`, { ...read, signal },
      (items, page, total) => onProgress && onProgress(total, page));
  }

  function listDisabledDevices(onProgress, signal) {
    return Http.pageAll(`${V1}/devices?$filter=accountEnabled eq false&$top=999&$select=${DEVICE_FIELDS}`, { ...read, signal },
      (items, page, total) => onProgress && onProgress(total, page));
  }

  // Intune enrichment is optional: no licence, no Intune or no permission still lets the scan run.
  async function tryListManagedDevices(onProgress, signal) {
    try {
      const list = await Http.pageAll(`${BETA}/deviceManagement/managedDevices?$select=${MANAGED_FIELDS}&$top=999`, { ...read, signal },
        (items, page, total) => onProgress && onProgress(total, page));
      return { list, unavailable: false, message: '' };
    } catch (e) {
      if (e && [401, 403, 404].includes(e.status)) {
        Log.warn('Intune managedDevices not readable:', e.message);
        return { list: [], unavailable: true, message: 'Intune data is not available (no permission, no licence or no Intune). Duplicate detection still works, Autopilot detection falls back to the Entra [ZTDId] marker.' };
      }
      throw e;
    }
  }

  // Fresh state of one device, or null when it no longer exists.
  async function getDevice(id) {
    try {
      return await Http.json(`${deviceUrl(id)}?$select=id,displayName,accountEnabled,extensionAttributes`, read);
    } catch (e) {
      if (e && e.status === 404) return null;
      throw e;
    }
  }

  function disableDevice(id, extProps) {
    const body = { accountEnabled: false };
    if (extProps && Object.keys(extProps).length) body.extensionAttributes = extProps;
    return Http.json(deviceUrl(id), { ...write, method: 'PATCH', body });
  }

  const deleteDevice = (id) => Http.json(deviceUrl(id), { ...write, method: 'DELETE' });

  window.Graph = { listAllDevices, listDisabledDevices, tryListManagedDevices, getDevice, disableDevice, deleteDevice };
})();
