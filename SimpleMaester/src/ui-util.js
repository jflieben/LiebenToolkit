// Small UI helpers shared by the SimpleMaester views.
(() => {
  const STATUS = { Passed: 'ok', Failed: 'danger', Error: 'warn', Skipped: 'neutral' };
  const SEVERITY = { Critical: 'danger', High: 'danger', Medium: 'warn', Low: 'info', Info: 'neutral' };
  const STATUS_KEYS = new Set(['passed', 'failed', 'error', 'skipped']);

  const esc = (v) => Toolkit.esc(v);
  const statusClass = (s) => STATUS[s] || 'neutral';
  const severityClass = (s) => SEVERITY[s] || 'neutral';
  // Class for sparkline bars; anything unknown counts as skipped.
  const statusKey = (s) => (STATUS_KEYS.has(String(s || '').toLowerCase()) ? String(s).toLowerCase() : 'skipped');

  function badge(cls, text) { return `<span class="badge ${cls}">${esc(text)}</span>`; }
  function statusBadge(s) { return badge(statusClass(s), s || '-'); }
  function severityBadge(s) { return badge(severityClass(s), s || '-'); }
  function tagBadge(t) { return badge('info', t || '-'); }

  function sparkHtml(points) {
    return `<div class="spark">${points.map((p) => `<span class="${statusKey(p.status)}" title="${esc(`${p.at}: ${p.status || '-'}`)}"></span>`).join('')}</div>`;
  }

  window.SMUi = { esc, statusClass, severityClass, statusKey, badge, statusBadge, severityBadge, tagBadge, sparkHtml };
})();
