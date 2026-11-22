// CSV and XLSX export. Text that starts with = + - @ or a tab/CR is prefixed with
// a quote so Excel shows it instead of running it as a formula.
//
//   Export.csv(rows, { columns: [{ key, label, format }], filename })
//   Export.xlsx([{ name, rows, columns }], filename)
(() => {
  const XLSX_SRC = 'xlsx.0.20.3.full.min.js';

  function neutralise(v) {
    if (typeof v !== 'string') return v;
    return /^[=+\-@\t\r]/.test(v) ? `'${v}` : v;
  }

  function toCell(v) {
    if (v == null) return '';
    if (v instanceof Date) return isNaN(v) ? '' : v.toISOString().replace('T', ' ').substring(0, 19);
    if (typeof v === 'boolean') return v ? 'Yes' : 'No';
    if (Array.isArray(v)) return v.join('; ');
    if (typeof v === 'object') { try { return JSON.stringify(v); } catch { return String(v); } }
    return v;
  }

  // rows: array of objects. columns optional; default is every key of the first row.
  function table(rows, columns) {
    const cols = columns && columns.length
      ? columns
      : Object.keys(rows[0] || {}).map((k) => ({ key: k, label: k }));
    const header = cols.map((c) => c.label || c.key);
    const body = rows.map((r) => cols.map((c) => neutralise(toCell(c.format ? c.format(r[c.key], r) : r[c.key]))));
    return { header, body };
  }

  function save(blob, filename) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  function csv(rows, { columns, filename = 'export.csv', delimiter = ',' } = {}) {
    if (!rows || !rows.length) return false;
    const { header, body } = table(rows, columns);
    const q = (v) => {
      const s = String(v == null ? '' : v);
      return /[",\n\r;]/.test(s) || s.includes(delimiter) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const text = [header, ...body].map((r) => r.map(q).join(delimiter)).join('\r\n');
    // BOM so Excel opens UTF-8 names correctly.
    save(new Blob(['﻿', text], { type: 'text/csv;charset=utf-8' }), filename);
    return true;
  }

  let xlsxLoading = null;
  function loadXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxLoading) {
      xlsxLoading = new Promise((resolve, reject) => {
        const base = (window.Toolkit && Toolkit.paths) ? Toolkit.paths.root : '../';
        const s = document.createElement('script');
        s.src = `${base}vendor/${XLSX_SRC}`;
        s.onload = () => resolve(window.XLSX);
        s.onerror = () => { xlsxLoading = null; reject(new Error('Could not load the Excel library')); };
        document.head.appendChild(s);
      });
    }
    return xlsxLoading;
  }

  // sheets: [{ name, rows, columns }]
  async function xlsx(sheets, filename = 'export.xlsx') {
    const X = await loadXlsx();
    const wb = X.utils.book_new();
    for (const sh of sheets) {
      const { header, body } = table(sh.rows || [], sh.columns);
      const ws = X.utils.aoa_to_sheet([header, ...body]);
      ws['!cols'] = header.map((hd, i) => ({ wch: Math.min(60, Math.max(String(hd).length, ...body.slice(0, 200).map((r) => String(r[i] ?? '').length)) + 2) }));
      X.utils.book_append_sheet(wb, ws, String(sh.name || 'Sheet').replace(/[\\/?*[\]:]/g, ' ').substring(0, 31));
    }
    X.writeFile(wb, filename);
    return true;
  }

  function json(data, filename = 'export.json') {
    save(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), filename);
    return true;
  }

  function stamp() { return new Date().toISOString().substring(0, 10); }

  window.Export = { csv, xlsx, json, table, neutralise, loadXlsx, save, stamp };
})();
