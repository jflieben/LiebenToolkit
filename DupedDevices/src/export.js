// CSV / XLSX export of duplicate device groups.
(() => {
  function fmt(d) {
    if (!d) return '';
    if (typeof d === 'string') d = new Date(d);
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    return d.toISOString().substring(0, 19).replace('T', ' ');
  }

  function rowsFromGroups(groups, attrName) {
    const rows = [];
    for (const g of groups) {
      for (const d of g.devices) {
        const ext = d.extensionAttributes || {};
        rows.push({
          HardwareId: g.hwid,
          DuplicateCount: g.deviceCount,
          Role: d === g.current ? 'KEEP (current)' : 'STALE (candidate)',
          DisplayName: d.displayName || '',
          DeviceId: d.deviceId || '',
          ObjectId: d.id || '',
          OperatingSystem: [d.operatingSystem, d.operatingSystemVersion].filter(Boolean).join(' '),
          TrustType: d.trustType || '',
          Autopilot: d.isAutopilot ? 'Yes' : 'No',
          PrimaryUser: d.primaryUser || '',
          Enabled: d.accountEnabled === false ? 'No' : 'Yes',
          LastSignIn: fmt(d.approximateLastSignInDateTime),
          Created: fmt(d.createdDateTime),
          Managed: d.isManaged ? 'Yes' : 'No',
          Compliant: d.isCompliant ? 'Yes' : (d.isCompliant === false ? 'No' : ''),
          DisableStamp: attrName ? (ext[attrName] || '') : '',
        });
      }
    }
    return rows;
  }

  function csvEscape(v) {
    const s = (v == null ? '' : String(v));
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function exportCsv(groups, attrName) {
    const rows = rowsFromGroups(groups, attrName);
    if (!rows.length) { Log.warn('Nothing to export'); return; }
    const headers = Object.keys(rows[0]);
    const lines = [headers.join(',')];
    for (const r of rows) lines.push(headers.map(h => csvEscape(r[h])).join(','));
    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    download(blob, `duped-devices-${stamp()}.csv`);
  }

  function exportXlsx(groups, attrName) {
    const rows = rowsFromGroups(groups, attrName);
    if (!rows.length) { Log.warn('Nothing to export'); return; }
    if (!window.XLSX) { Log.err('xlsx library missing, falling back to CSV'); return exportCsv(groups, attrName); }
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Duplicates');
    XLSX.writeFile(wb, `duped-devices-${stamp()}.xlsx`);
  }

  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 0);
  }
  function stamp() { return new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19); }

  window.Exporter = { exportCsv, exportXlsx, rowsFromGroups };
})();
