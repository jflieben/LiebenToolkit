// XLSX / CSV export
(() => {
  function fmt(d) {
    if (!d) return '';
    if (d instanceof Date) return d.toISOString().substring(0, 10);
    return String(d);
  }

  function rowsFromResults(results) {
    return results.map(r => ({
      Title: r.site.displayName || r.site.name || '',
      URL: r.site.webUrl || '',
      Type: classifyType(r.site),
      Owners: (r.owners || []).map(o => o.title || o.email || o.loginName || '').filter(Boolean).join('; '),
      OwnerEmails: (r.owners || []).map(o => o.email || '').filter(Boolean).join('; '),
      Template: r.site.root?.['@odata.type'] || r.site.siteCollection?.root?.['@odata.type'] || '',
      Created: fmt(new Date(r.site.createdDateTime || r.signals.spoWebCreated || '')),
      LastModified: fmt(r.lastModified),
      LastViewed: fmt(r.lastViewed),
      LastViewedApprox: r.lastViewedApprox ? 'Yes' : '',
      DaysIdle: r.daysSinceModified ?? '',
      PageViewsAllTime: r.pageViews ?? '',
      ViewsLast30Days: r.last30Views ?? '',
      ViewsLast7Days: r.recentViews ?? '',
      StorageMB: r.storageMB ?? '',
      Recommendation: r.advice,
      Why: r.reason,
      GraphLastModified: fmt(r.signals.graphSiteLastModified),
      WebLastUserModified: fmt(r.signals.spoWebLastUserModified),
      WebLastModified: fmt(r.signals.spoWebLastModified),
      LibMaxUserModified: fmt(r.signals.libMaxUserModified),
      LibCount: r.signals.libCount ?? '',
      LibTotalItems: r.signals.libTotalItems ?? '',
      LastPageEditDate: fmt(r.signals.lastPageEditDate),
      LastPageEditName: r.signals.lastPageEditName ?? '',
      LastPageEditBy: r.signals.lastPageEditBy ?? '',
      AnalyticsAllTimeViews: r.signals.analyticsAllTimeViews ?? '',
      AnalyticsAllTimeActors: r.signals.analyticsAllTimeActors ?? '',
      AnalyticsRecentViews: r.signals.analyticsRecentViews ?? '',
      AnalyticsRecentActors: r.signals.analyticsRecentActors ?? '',
      Last30DayLastAccess: fmt(r.signals.last30LastDay),
      Errors: (r.errors || []).join(' | '),
    }));
  }

  /**
   * Detailed site-type classification based on the SharePoint WebTemplate.
   * The template is captured during discovery in `site.root['@odata.type']`.
   * Examples:
   *   GROUP#0              → Microsoft 365 group site (often Teams-connected)
   *   TEAMCHANNEL#0        → Teams private channel
   *   TEAMCHANNEL#1        → Teams shared channel
   *   SITEPAGEPUBLISHING#0 → Communication site
   *   STS#3 / STS#0        → Classic team site
   *   SPSPERS#10           → OneDrive personal site
   *   APPCATALOG#0         → App catalog
   *   SRCHCEN#0            → Search center
   *   EHS#1                → Modern team site (no group)
   *   POINTPUBLISHINGHUB#0 → Hub site (publishing)
   *   POINTPUBLISHINGTOPIC#0 → Topic site (Viva)
   */
  function classifyType(site) {
    const tpl = ((site.root && site.root['@odata.type']) || '').toUpperCase();
    const url = (site.webUrl || '').toLowerCase();
    if (site.isPersonalSite || tpl.startsWith('SPSPERS') || /-my\.sharepoint\.com\/personal\//i.test(url)) return 'OneDrive';
    if (tpl.startsWith('TEAMCHANNEL#0')) return 'Teams private channel';
    if (tpl.startsWith('TEAMCHANNEL#1')) return 'Teams shared channel';
    if (tpl.startsWith('TEAMCHANNEL'))   return 'Teams channel';
    if (tpl.startsWith('GROUP'))         return 'Group / Teams site';
    if (tpl.startsWith('SITEPAGEPUBLISHING')) return 'Communication site';
    if (tpl.startsWith('POINTPUBLISHINGHUB'))   return 'Hub site';
    if (tpl.startsWith('POINTPUBLISHINGTOPIC')) return 'Topic site (Viva)';
    if (tpl.startsWith('STS#3'))         return 'Modern team site';
    if (tpl.startsWith('STS'))           return 'Classic team site';
    if (tpl.startsWith('EHS'))           return 'Modern team site (no group)';
    if (tpl.startsWith('APPCATALOG'))    return 'App catalog';
    if (tpl.startsWith('SRCHCEN'))       return 'Search center';
    if (tpl.startsWith('ENTERWIKI'))     return 'Enterprise wiki';
    if (tpl.startsWith('BICENTERSITE'))  return 'BI center';
    if (tpl.startsWith('BLANKINTERNET')) return 'Publishing site';
    if (tpl.startsWith('TENANTADMIN'))   return 'Tenant admin';
    if (tpl) return tpl;
    return 'SharePoint';
  }

  /** Coarse two-bucket category for badge styling. */
  function categoryOf(site) {
    return classifyType(site) === 'OneDrive' ? 'OneDrive' : 'SharePoint';
  }

  // Backward-compat alias
  function classify(site) { return categoryOf(site); }

  function exportXlsx(results) {
    const rows = rowsFromResults(results);
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'StaleSites');
    const stamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    XLSX.writeFile(wb, `SPOTrim-StaleSites-${stamp}.xlsx`);
  }

  function exportCsv(results) {
    const rows = rowsFromResults(results);
    if (!rows.length) return;
    const headers = Object.keys(rows[0]);
    const csv = [headers.join(',')]
      .concat(rows.map(r => headers.map(h => csvEscape(r[h])).join(',')))
      .join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const stamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    triggerDownload(blob, `SPOTrim-StaleSites-${stamp}.csv`);
  }

  function csvEscape(v) {
    if (v === null || v === undefined) return '';
    const s = String(v);
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function triggerDownload(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 500);
  }

  window.Exporter = { exportXlsx, exportCsv, classify, classifyType, categoryOf };
})();
