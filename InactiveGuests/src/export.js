// Excel + CSV export for InactiveGuests.
(() => {
  const COLS = [
    { key: 'displayName',     label: 'Display name' },
    { key: 'userPrincipalName', label: 'UPN' },
    { key: 'mail',            label: 'Mail' },
    { key: 'companyName',     label: 'Company' },
    { key: 'accountEnabled',  label: 'Enabled' },
    { key: 'redemptionState', label: 'Redemption' },
    { key: 'createdDateTime', label: 'Created' },
    { key: 'accountAgeDays',  label: 'Age (days)' },
    { key: 'lastSignIn',      label: 'Last sign-in' },
    { key: 'inactiveDays',    label: 'Inactive (days)' },
    { key: 'id',              label: 'Object id' },
  ];

  function rowsFor(guests) {
    return guests.map(g => {
      const r = {};
      for (const c of COLS) {
        let v = g[c.key];
        if (c.key === 'createdDateTime' || c.key === 'lastSignIn') v = Guests.fmtDate(v);
        if (typeof v === 'boolean') v = v ? 'Yes' : 'No';
        r[c.label] = v == null ? '' : v;
      }
      return r;
    });
  }

  async function ensureXlsx() {
    if (window.XLSX) return window.XLSX;
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = '../vendor/xlsx.0.18.5.full.min.js';
      s.onload = res; s.onerror = () => rej(new Error('Failed to load xlsx from CDN'));
      document.head.appendChild(s);
    });
    return window.XLSX;
  }

  async function exportXlsx(guests, filename = 'inactive-guests.xlsx') {
    const XLSX = await ensureXlsx();
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rowsFor(guests));
    XLSX.utils.book_append_sheet(wb, ws, 'Guests');
    XLSX.writeFile(wb, filename);
  }

  function exportCsv(guests, filename = 'inactive-guests.csv') {
    const rows = rowsFor(guests);
    if (!rows.length) return;
    const headers = Object.keys(rows[0]);
    const esc = v => {
      const s = String(v == null ? '' : v);
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const csv = [headers.join(','), ...rows.map(r => headers.map(h => esc(r[h])).join(','))].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 0);
  }

  window.Export = { exportXlsx, exportCsv };
})();
