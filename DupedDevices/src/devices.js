// Duplicate detection, grace-period logic and other pure helpers for DupedDevices.
(() => {
  const DAY = 86400000;
  const STAMP_PREFIX = 'DupedDevices:disabled:';
  const GRACE_MIN = 1;
  const GRACE_MAX = 3650;

  // physicalIds entries:
  //   "[HWID]:value"             hardware id
  //   "[USER-HWID]:userId:value" per-user variant, last token is the hardware id
  //   "[GID]:..." / "[USER-GID]:..." group ids, not hardware
  function extractHwid(device) {
    const ids = Array.isArray(device.physicalIds) ? device.physicalIds : [];
    let hwid = null;
    for (const raw of ids) {
      if (typeof raw !== 'string') continue;
      if (raw.startsWith('[HWID]:')) {
        const v = raw.substring('[HWID]:'.length).trim();
        if (v) return v;
      } else if (raw.startsWith('[USER-HWID]:')) {
        const parts = raw.split(':');
        const v = parts[parts.length - 1];
        if (v) hwid = hwid || v.trim();
      }
    }
    return hwid;
  }

  function lastActivity(d) {
    const a = d.approximateLastSignInDateTime ? new Date(d.approximateLastSignInDateTime).getTime() : 0;
    const c = d.createdDateTime ? new Date(d.createdDateTime).getTime() : 0;
    return Math.max(a || 0, c || 0);
  }

  function keyForDevice(device) {
    return String((device && device.deviceId) || '').toLowerCase();
  }

  function buildManagedMap(list) {
    const map = new Map();
    for (const m of list || []) {
      const key = String(m.azureADDeviceId || '').toLowerCase();
      if (key) map.set(key, m);
    }
    return map;
  }

  // Intune autopilotEnrolled first, the Entra [ZTDId] marker as fallback.
  function isAutopilot(device, managed) {
    if (managed && managed.autopilotEnrolled === true) return true;
    const ids = Array.isArray(device && device.physicalIds) ? device.physicalIds : [];
    return ids.some((v) => typeof v === 'string' && v.toLowerCase().startsWith('[ztdid]'));
  }

  // Groups of more than one device per hardware id. The most recently active one is
  // `current` (keep), the rest are `stale`. Disabled devices only count when asked,
  // so an old disabled record cannot shadow a fresh re-enrolment.
  function findDuplicates(devices, opts = {}) {
    const pool = devices.filter((d) => opts.includeDisabledInComparison || d.accountEnabled !== false);
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
      groups.push({ hwid, deviceCount: list.length, devices: list, current: list[0], stale: list.slice(1) });
    }
    groups.sort((a, b) => b.deviceCount - a.deviceCount || (a.devices[0].displayName || '').localeCompare(b.devices[0].displayName || ''));
    return groups;
  }

  function isMixedAutopilotGroup(group) {
    let ap = false, nonAp = false;
    for (const d of group.devices || []) {
      if (d.isAutopilot) ap = true; else nonAp = true;
      if (ap && nonAp) return true;
    }
    return false;
  }

  function filterGroups(groups, { search = '', hideMixedAutopilot = false } = {}) {
    const q = String(search).toLowerCase().trim();
    return groups.filter((g) => {
      if (hideMixedAutopilot && isMixedAutopilotGroup(g)) return false;
      if (!q) return true;
      const hay = `${g.hwid} ${g.devices.map((d) => `${d.displayName || ''} ${d.operatingSystem || ''}`).join(' ')}`.toLowerCase();
      return hay.includes(q);
    });
  }

  function staleIds(groups) {
    const out = new Set();
    for (const g of groups) for (const d of g.stale) out.add(d.id);
    return out;
  }

  // Drops ids that are not in `allowed` (hidden by a filter). Returns how many went.
  function trimSelection(selected, allowed) {
    let dropped = 0;
    for (const id of [...selected]) if (!allowed.has(id)) { selected.delete(id); dropped++; }
    return dropped;
  }

  function buildDisableStampValue(now = new Date()) {
    return `${STAMP_PREFIX}${now.toISOString()}`;
  }

  function parseDupedDevicesStamp(value) {
    if (typeof value !== 'string' || !value.startsWith(STAMP_PREFIX)) return null;
    const d = new Date(value.substring(STAMP_PREFIX.length));
    return isNaN(d.getTime()) ? null : d;
  }

  // Whole days, 1 to 3650. Anything else is refused (null) rather than corrected.
  function parseGraceDays(raw) {
    const s = String(raw == null ? '' : raw).trim();
    if (!/^\d+$/.test(s)) return null;
    const n = parseInt(s, 10);
    return n >= GRACE_MIN && n <= GRACE_MAX ? n : null;
  }

  // Disabled devices whose DupedDevices stamp in `attrName` is at least `graceDays` old.
  function findExpiredDisabled(devices, attrName, graceDays, now = new Date()) {
    if (!attrName || parseGraceDays(graceDays) == null) return [];
    const cutoff = now.getTime() - graceDays * DAY;
    const out = [];
    for (const d of devices) {
      if (d.accountEnabled !== false) continue;
      const value = (d.extensionAttributes || {})[attrName];
      const stamp = parseDupedDevicesStamp(value);
      if (!stamp || stamp.getTime() > cutoff) continue;
      out.push({ device: d, stampValue: value, disabledAt: stamp, daysDisabled: Math.floor((now.getTime() - stamp.getTime()) / DAY) });
    }
    out.sort((a, b) => a.disabledAt - b.disabledAt);
    return out;
  }

  // Check a device freshly read from Graph right before a grace-period delete.
  function recheckExpired(fresh, attrName, expectedValue, cutoffMs) {
    if (!fresh) return { ok: false, reason: 'no longer exists' };
    if (fresh.accountEnabled !== false) return { ok: false, reason: 'enabled again since the scan' };
    const value = (fresh.extensionAttributes || {})[attrName];
    if (value !== expectedValue) return { ok: false, reason: `${attrName} changed since the scan` };
    const stamp = parseDupedDevicesStamp(value);
    if (!stamp || stamp.getTime() > cutoffMs) return { ok: false, reason: 'not past the grace period' };
    return { ok: true };
  }

  function fmtDate(d) {
    if (!d) return '';
    const dt = d instanceof Date ? d : new Date(d);
    return isNaN(dt.getTime()) ? '' : dt.toISOString().substring(0, 10);
  }

  function osLabel(d) {
    return [d.operatingSystem, d.operatingSystemVersion].filter(Boolean).join(' ');
  }

  // One line per device for confirm dialogs.
  function describe(d) {
    const last = fmtDate(d.approximateLastSignInDateTime);
    return [d.displayName || '(no name)', osLabel(d) || 'unknown OS', last ? `last active ${last}` : 'never active'].join('  |  ');
  }

  function rowsFromGroups(groups, attrName) {
    const rows = [];
    for (const g of groups) {
      for (const d of g.devices) {
        const ext = d.extensionAttributes || {};
        rows.push({
          hwid: g.hwid,
          count: g.deviceCount,
          role: d === g.current ? 'Keep' : 'Stale',
          displayName: d.displayName || '',
          deviceId: d.deviceId || '',
          id: d.id || '',
          os: osLabel(d),
          trustType: d.trustType || '',
          autopilot: !!d.isAutopilot,
          primaryUser: d.primaryUser || '',
          enabled: d.accountEnabled !== false,
          lastSignIn: fmtDate(d.approximateLastSignInDateTime),
          created: fmtDate(d.createdDateTime),
          managed: !!d.isManaged,
          compliant: d.isCompliant == null ? '' : d.isCompliant ? 'Yes' : 'No',
          stamp: attrName ? (ext[attrName] || '') : '',
        });
      }
    }
    return rows;
  }

  window.Devices = {
    extractHwid, lastActivity, keyForDevice, buildManagedMap, isAutopilot,
    findDuplicates, isMixedAutopilotGroup, filterGroups, staleIds, trimSelection,
    buildDisableStampValue, parseDupedDevicesStamp, parseGraceDays, findExpiredDisabled, recheckExpired,
    fmtDate, osLabel, describe, rowsFromGroups, STAMP_PREFIX,
  };
})();
