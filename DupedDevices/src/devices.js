// Duplicate-device detection and grouping logic.
// Groups Entra devices that share a hardware ID and decides which one is the
// "current" registration (= keep) and which are stale (= candidates for cleanup).
(() => {
  // physicalIds entries look like:
  //   "[HWID]:abc123..."          (classic single-user / group-joined device)
  //   "[USER-HWID]:userId:hwid"   (newer per-user-per-device variant)
  //   "[USER-GID]:..."            (group ID per user, NOT a hardware identifier)
  //   "[GID]:..."                 (group ID, NOT a hardware identifier)
  // We treat a device as having a hardware ID only when [HWID] or [USER-HWID]
  // is present and we use the trailing token as the canonical key.
  function extractHwid(device) {
    const ids = Array.isArray(device.physicalIds) ? device.physicalIds : [];
    let hwid = null;
    for (const raw of ids) {
      if (typeof raw !== 'string') continue;
      if (raw.startsWith('[HWID]:')) {
        // "[HWID]:value"  -> value
        hwid = raw.substring('[HWID]:'.length).trim();
        if (hwid) return hwid;
      } else if (raw.startsWith('[USER-HWID]:')) {
        // "[USER-HWID]:userId:value" -> last part
        const parts = raw.split(':');
        const v = parts[parts.length - 1];
        if (v) hwid = hwid || v.trim();
      }
    }
    return hwid;
  }

  /** Pick the most recent activity timestamp for a device. */
  function lastActivity(d) {
    const a = d.approximateLastSignInDateTime ? new Date(d.approximateLastSignInDateTime).getTime() : 0;
    const c = d.createdDateTime ? new Date(d.createdDateTime).getTime() : 0;
    return Math.max(a, c);
  }

  function keyForDevice(device) {
    return String((device && device.deviceId) || '').toLowerCase();
  }

  function buildManagedMap(list) {
    const map = new Map();
    for (const m of list || []) {
      const key = String(m.azureADDeviceId || '').toLowerCase();
      if (!key) continue;
      map.set(key, m);
    }
    return map;
  }

  /**
   * Autopilot recognition mirrors InactiveDevices:
   * 1) Prefer Intune's managedDevice.autopilotEnrolled when available.
   * 2) Fallback to Entra physicalIds marker [ZTDId].
   */
  function isAutopilot(device, managed) {
    if (managed && managed.autopilotEnrolled === true) return true;
    const ids = Array.isArray(device && device.physicalIds) ? device.physicalIds : [];
    return ids.some((v) => typeof v === 'string' && v.toLowerCase().startsWith('[ztdid]'));
  }

  /**
   * Group devices by hardware ID and return only the groups with > 1 entries.
   * Each group sorts devices "newest first" and flags the most recently active
   * one as `current`. The rest are candidates for disable/delete.
   */
  function findDuplicates(devices, opts = {}) {
    const includeDisabledInComparison = !!opts.includeDisabledInComparison;
    // Following the original PowerShell logic: only enabled devices count
    // when looking for duplicates, otherwise an old "still disabled" record
    // would shadow a fresh re-enrollment.
    const pool = devices.filter(d => includeDisabledInComparison || d.accountEnabled !== false);
    const byHwid = new Map();
    for (const d of pool) {
      const h = extractHwid(d);
      if (!h) continue;
      if (!byHwid.has(h)) byHwid.set(h, []);
      byHwid.get(h).push(d);
    }
    const groups = [];
    for (const [hwid, list] of byHwid) {
      if (list.length < 2) continue;
      list.sort((a, b) => lastActivity(b) - lastActivity(a));
      const current = list[0];
      const stale   = list.slice(1);
      groups.push({
        hwid,
        deviceCount: list.length,
        devices: list,
        current,
        stale,
      });
    }
    // Stable, useful order: most-stale-devices first.
    groups.sort((a, b) => b.deviceCount - a.deviceCount || (a.devices[0].displayName || '').localeCompare(b.devices[0].displayName || ''));
    return groups;
  }

  /** Read the disable timestamp the tool may have stored earlier on a device. */
  function getDisableStamp(device, attrName) {
    if (!attrName) return null;
    const ext = device.extensionAttributes || {};
    const v = ext[attrName];
    if (!v) return null;
    const parsed = new Date(v);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  /** Identify devices that were previously disabled by this tool and have aged past the grace period. */
  function findExpiredDisabled(devices, attrName, gracePeriodDays) {
    if (!attrName || !(gracePeriodDays >= 0)) return [];
    const cutoff = Date.now() - gracePeriodDays * 86400000;
    const out = [];
    for (const d of devices) {
      if (d.accountEnabled !== false) continue;
      const stamp = getDisableStamp(d, attrName);
      if (!stamp) continue;
      if (stamp.getTime() <= cutoff) {
        out.push({ device: d, disabledAt: stamp, daysDisabled: Math.floor((Date.now() - stamp.getTime()) / 86400000) });
      }
    }
    out.sort((a, b) => a.disabledAt - b.disabledAt);
    return out;
  }

  /** Build the timestamp value we'll write into the chosen extension attribute. */
  function buildDisableStampValue() {
    return `DupedDevices:disabled:${new Date().toISOString()}`;
  }

  /** Parse a "DupedDevices:disabled:<iso>" stamp back into a Date. */
  function parseDupedDevicesStamp(value) {
    if (typeof value !== 'string') return null;
    const m = value.match(/^DupedDevices:disabled:(.+)$/);
    if (!m) return null;
    const d = new Date(m[1]);
    return isNaN(d.getTime()) ? null : d;
  }

  window.Devices = {
    extractHwid,
    lastActivity,
    keyForDevice,
    buildManagedMap,
    isAutopilot,
    findDuplicates,
    findExpiredDisabled,
    getDisableStamp,
    buildDisableStampValue,
    parseDupedDevicesStamp,
  };
})();
