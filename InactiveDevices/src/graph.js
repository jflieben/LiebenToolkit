// Microsoft Graph calls for InactiveDevices. Devices come from beta for registrationDateTime
// and profileType; writes use v1.0.
(() => {
  const V1 = 'https://graph.microsoft.com/v1.0';
  const BETA = 'https://graph.microsoft.com/beta';

  const DEVICE_FIELDS = [
    'id', 'deviceId', 'displayName', 'accountEnabled', 'createdDateTime',
    'registrationDateTime', 'approximateLastSignInDateTime', 'operatingSystem',
    'operatingSystemVersion', 'trustType', 'profileType', 'isManaged',
    'isCompliant', 'onPremisesSyncEnabled', 'physicalIds', 'extensionAttributes',
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
    return Http.pageAll(`${BETA}/devices?$select=${DEVICE_FIELDS}&$top=999`, { ...read, signal },
      (items, page, total) => onProgress && onProgress(total, page));
  }

  function listDisabledDevices(onProgress, signal) {
    return Http.pageAll(`${BETA}/devices?$filter=accountEnabled eq false&$select=${DEVICE_FIELDS}&$top=999`, { ...read, signal },
      (items, page, total) => onProgress && onProgress(total, page));
  }

  // Intune enrichment is optional: no licence, no Intune or no permission still lists devices.
  async function tryListManagedDevices(onProgress, signal) {
    try {
      const list = await Http.pageAll(`${BETA}/deviceManagement/managedDevices?$select=${MANAGED_FIELDS}&$top=999`, { ...read, signal },
        (items, page, total) => onProgress && onProgress(total, page));
      return { list, unavailable: false, message: '' };
    } catch (e) {
      if (e && [401, 403, 404].includes(e.status)) {
        Log.warn('Intune managedDevices not readable:', e.message);
        return { list: [], unavailable: true, message: 'Intune data is not available (no permission, no licence or no Intune). Devices are still listed, without primary user and Intune sync date.' };
      }
      throw e;
    }
  }

  // Fresh state of one device, or null when it no longer exists.
  async function getDevice(id) {
    try {
      return await Http.json(`${deviceUrl(id)}?$select=id,displayName,accountEnabled,extensionAttributes,physicalIds`, read);
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
  const enableDevice = (id) => Http.json(deviceUrl(id), { ...write, method: 'PATCH', body: { accountEnabled: true } });
  const deleteDevice = (id) => Http.json(deviceUrl(id), { ...write, method: 'DELETE' });

  window.Graph = { listAllDevices, listDisabledDevices, tryListManagedDevices, getDevice, disableDevice, enableDevice, deleteDevice };
})();
