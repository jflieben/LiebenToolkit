// Pure helpers for InactiveDevices - safe to unit-test.
(() => {
  function parseDate(value) {
    if (!value) return null;
    const d = value instanceof Date ? value : new Date(value);
    return isNaN(d.getTime()) ? null : d;
  }

  function daysBetween(a, b) {
    if (!a || !b) return null;
    return Math.floor((b.getTime() - a.getTime()) / 86400000);
  }

  function normalizeJoinType(trustType) {
    switch ((trustType || '').toLowerCase()) {
      case 'azuread': return 'Azure AD joined';
      case 'serverad': return 'Hybrid Azure AD joined';
      case 'workplace': return 'Azure AD registered';
      default: return 'Unknown';
    }
  }

  /**
   * Returns true if the device is (or was) registered via Windows Autopilot.
   * Uses two complementary signals:
   *   1. managed.autopilotEnrolled  — accurate for Intune-managed devices
   *   2. devicePhysicalIds [ZTDId]  — fallback for devices not in Intune
   */
  function isAutopilot(device, managed) {
    if (managed && managed.autopilotEnrolled === true) return true;
    const ids = device.devicePhysicalIds;
    if (!Array.isArray(ids)) return false;
    return ids.some((v) => typeof v === 'string' && v.toLowerCase().startsWith('[ztdid]'));
  }

  function osBucket(os) {
    const v = String(os || '').toLowerCase();
    if (!v) return 'other';
    if (v.includes('windows')) return 'windows';
    if (v.includes('ios')) return 'ios';
    if (v.includes('android')) return 'android';
    if (v.includes('mac')) return 'macos';
    if (v.includes('linux')) return 'linux';
    return 'other';
  }

  function keyForDevice(device) {
    return (device.deviceId || device.id || '').toLowerCase();
  }

  function buildManagedMap(managedDevices) {
    const map = new Map();
    for (const md of managedDevices || []) {
      const key = String(md.azureADDeviceId || '').toLowerCase();
      if (!key) continue;
      map.set(key, md);
    }
    return map;
  }

  function enrich(device, managedMap, now = new Date()) {
    const created = parseDate(device.createdDateTime || device.registrationDateTime);
    const lastSignIn = parseDate(device.approximateLastSignInDateTime);
    const managed = managedMap ? managedMap.get(keyForDevice(device)) : null;
    const lastSync = parseDate(managed && managed.lastSyncDateTime);

    let lastActive = lastSignIn;
    if (lastSync && (!lastActive || lastSync.getTime() > lastActive.getTime())) {
      lastActive = lastSync;
    }

    const deviceAgeDays = created ? daysBetween(created, now) : null;
    const inactiveDays = lastActive ? daysBetween(lastActive, now) : deviceAgeDays;

    const primaryUser = (managed && managed.userPrincipalName) || '';
    const operatingSystem = device.operatingSystem || (managed && managed.operatingSystem) || '';

    const deviceIdKey = (device.deviceId || '').toLowerCase();
    return {
      ...device,
      primaryUser,
      joinType: normalizeJoinType(device.trustType),
      isAutopilot: isAutopilot(device, managed),
      isManagedEffective: !!(device.isManaged || managed),
      lastSignIn,
      lastSync,
      lastActive,
      inactiveDays,
      deviceAgeDays,
      operatingSystem,
      osBucket: osBucket(operatingSystem),
      managedDevice: managed || null,
      neverActive: !lastActive,
    };
  }

  function passesFilter(d, f) {
    if (f.search) {
      const q = f.search.toLowerCase();
      const hay = [
        d.displayName, d.deviceId, d.id, d.primaryUser,
        d.operatingSystem, d.operatingSystemVersion,
      ].join('\n').toLowerCase();
      if (!hay.includes(q)) return false;
    }

    if (typeof f.minInactive === 'number' && (d.inactiveDays == null || d.inactiveDays < f.minInactive)) return false;
    if (typeof f.minAge === 'number' && (d.deviceAgeDays == null || d.deviceAgeDays < f.minAge)) return false;

    if (f.osFilter && f.osFilter !== 'any' && d.osBucket !== f.osFilter) return false;
    if (f.joinType && f.joinType !== 'any' && d.joinType !== f.joinType) return false;

    if (f.autopilot === 'yes' && !d.isAutopilot) return false;
    if (f.autopilot === 'no' && d.isAutopilot) return false;

    if (f.managed === 'managed' && !d.isManagedEffective) return false;
    if (f.managed === 'unmanaged' && d.isManagedEffective) return false;

    const hasPrimary = !!d.primaryUser;
    if (f.hasPrimaryUser === 'yes' && !hasPrimary) return false;
    if (f.hasPrimaryUser === 'no' && hasPrimary) return false;

    if (f.enabled === 'enabled' && !d.accountEnabled) return false;
    if (f.enabled === 'disabled' && d.accountEnabled) return false;

    return true;
  }

  function fmtDate(d) {
    if (!d) return '';
    const dt = d instanceof Date ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    return dt.toISOString().substring(0, 10);
  }

  window.Devices = {
    parseDate,
    daysBetween,
    normalizeJoinType,
    isAutopilot,
    buildManagedMap,
    enrich,
    passesFilter,
    fmtDate,
  };
})();
