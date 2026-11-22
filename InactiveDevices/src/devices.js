// Pure helpers for InactiveDevices.
(() => {
  const DAY = 86400000;
  const STAMP_PREFIX = 'InactiveDevices:disabled:';

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
      case 'azuread': return 'Entra joined';
      case 'serverad': return 'Hybrid joined';
      case 'workplace': return 'Entra registered';
      default: return 'Unknown';
    }
  }

  // Intune autopilotEnrolled first, the Entra [ZTDId] marker in physicalIds as fallback.
  // Graph calls the property physicalIds; devicePhysicalIds is the dynamic-group name.
  function isAutopilot(device, managed) {
    if (managed && managed.autopilotEnrolled === true) return true;
    const ids = [].concat(device.physicalIds || [], device.devicePhysicalIds || []);
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

  function osLabel(d) {
    return [d.operatingSystem, d.operatingSystemVersion].filter(Boolean).join(' ');
  }

  // One line per device for confirm dialogs.
  function describe(d) {
    const last = fmtDate(d.lastActive || d.approximateLastSignInDateTime);
    return [d.displayName || d.deviceId || '(no name)', osLabel(d) || 'unknown OS', last ? `last active ${last}` : 'never active'].join('  |  ');
  }

  // Autopilot devices are never disabled or deleted by this tool.
  function splitActionable(rows, kind) {
    if (kind === 'enable') return { actionable: rows.slice(), skipped: [] };
    return { actionable: rows.filter((d) => !d.isAutopilot), skipped: rows.filter((d) => d.isAutopilot) };
  }

  function trimSelection(selected, allowed) {
    let dropped = 0;
    for (const id of [...selected]) if (!allowed.has(id)) { selected.delete(id); dropped++; }
    return dropped;
  }

  function buildDisableStampValue(now = new Date()) {
    return `${STAMP_PREFIX}${now.toISOString()}`;
  }

  function parseDisableStamp(value) {
    if (typeof value !== 'string' || !value.startsWith(STAMP_PREFIX)) return null;
    const d = new Date(value.substring(STAMP_PREFIX.length));
    return isNaN(d.getTime()) ? null : d;
  }

  // Whole days, 1 to 3650. Anything else is refused (null) rather than corrected.
  function parseGraceDays(raw) {
    const s = String(raw == null ? '' : raw).trim();
    if (!/^\d+$/.test(s)) return null;
    const n = parseInt(s, 10);
    return n >= 1 && n <= 3650 ? n : null;
  }

  function findExpiredDisabled(devices, attrName, graceDays, now = new Date()) {
    if (!attrName || parseGraceDays(graceDays) == null) return [];
    const cutoff = now.getTime() - graceDays * DAY;
    const out = [];
    for (const d of devices) {
      if (d.accountEnabled !== false) continue;
      const value = (d.extensionAttributes || {})[attrName];
      const stamp = parseDisableStamp(value);
      if (!stamp || stamp.getTime() > cutoff) continue;
      out.push({ device: d, stampValue: value, disabledAt: stamp, daysDisabled: Math.floor((now.getTime() - stamp.getTime()) / DAY) });
    }
    out.sort((a, b) => a.disabledAt - b.disabledAt);
    return out;
  }

  // Check a device freshly read from Graph right before a grace-period delete.
  function recheckExpired(fresh, attrName, expectedValue, cutoffMs) {
    if (!fresh) return { ok: false, reason: 'no longer exists' };
    if (fresh.accountEnabled !== false) return { ok: false, reason: 'enabled again since the search' };
    const value = (fresh.extensionAttributes || {})[attrName];
    if (value !== expectedValue) return { ok: false, reason: `${attrName} changed since the search` };
    const stamp = parseDisableStamp(value);
    if (!stamp || stamp.getTime() > cutoffMs) return { ok: false, reason: 'not past the grace period' };
    if (isAutopilot(fresh, null)) return { ok: false, reason: 'Autopilot device' };
    return { ok: true };
  }

  window.Devices = {
    parseDate, daysBetween, normalizeJoinType, isAutopilot, osBucket, buildManagedMap, enrich, passesFilter, fmtDate,
    osLabel, describe, splitActionable, trimSelection,
    buildDisableStampValue, parseDisableStamp, parseGraceDays, findExpiredDisabled, recheckExpired, STAMP_PREFIX,
  };
})();
