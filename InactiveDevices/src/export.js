// Excel + CSV export for InactiveDevices.
(() => {
  const COLS = [
    { key: 'displayName', label: 'Display name' },
    { key: 'deviceId', label: 'Device id' },
    { key: 'primaryUser', label: 'Primary user' },
    { key: 'operatingSystem', label: 'OS' },
    { key: 'operatingSystemVersion', label: 'OS version' },
    { key: 'joinType', label: 'Join type' },
    { key: 'isAutopilot', label: 'Autopilot' },
    { key: 'isManagedEffective', label: 'Managed' },
    { key: 'accountEnabled', label: 'Enabled' },
    { key: 'createdDateTime', label: 'Created' },
    { key: 'lastSignIn', label: 'Last sign-in (Entra)' },
    { key: 'lastSync', label: 'Last sync (Intune)' },
    { key: 'lastActive', label: 'Last active (effective)' },
    { key: 'inactiveDays', label: 'Inactive (days)' },
    { key: 'deviceAgeDays', label: 'Age (days)' },
    { key: 'id', label: 'Object id' },
  ];

  function rowsFor(devices) {
    return devices.map((d) => {
      const row = {};
      for (const c of COLS) {
        let v = d[c.key];
        if (c.key === 'createdDateTime' || c.key === 'lastSignIn' || c.key === 'lastSync' || c.key === 'lastActive') {
          v = Devices.fmtDate(v);
        }
        if (typeof v === 'boolean') v = v ? 'Yes' : 'No';
        row[c.label] = v == null ? '' : v;
      }
      return row;
    });
  }

  async function ensureXlsx() {
    if (window.XLSX) return window.XLSX;
    await new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
      s.onload = resolve;
      s.onerror = () => reject(new Error('Failed to load xlsx from CDN'));
      document.head.appendChild(s);
    });
    return window.XLSX;
  }

  async function exportXlsx(devices, filename = 'inactive-devices.xlsx') {
    const XLSX = await ensureXlsx();
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rowsFor(devices));
    XLSX.utils.book_append_sheet(wb, ws, 'Devices');
    XLSX.writeFile(wb, filename);
  }

  function exportCsv(devices, filename = 'inactive-devices.csv') {
    const rows = rowsFor(devices);
    if (!rows.length) return;
    const headers = Object.keys(rows[0]);
    const esc = (v) => {
      const s = String(v == null ? '' : v);
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const csv = [headers.join(','), ...rows.map((r) => headers.map((h) => esc(r[h])).join(','))].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 0);
  }

  window.Export = { exportXlsx, exportCsv };
})();
