// Site classification, tenant host checks and export rows. Pure, no network.
(() => {
  // Site type from the SharePoint WebTemplate captured during discovery (site.root['@odata.type']).
  function classifyType(site) {
    const tpl = ((site.root && site.root['@odata.type']) || '').toUpperCase();
    const url = (site.webUrl || '').toLowerCase();
    if (site.isPersonalSite || tpl.startsWith('SPSPERS') || /-my\.sharepoint\.com\/personal\//i.test(url)) return 'OneDrive';
    if (tpl.startsWith('TEAMCHANNEL#0')) return 'Teams private channel';
    if (tpl.startsWith('TEAMCHANNEL#1')) return 'Teams shared channel';
    if (tpl.startsWith('TEAMCHANNEL')) return 'Teams channel';
    if (tpl.startsWith('GROUP')) return 'Group / Teams site';
    if (tpl.startsWith('SITEPAGEPUBLISHING')) return 'Communication site';
    if (tpl.startsWith('POINTPUBLISHINGHUB')) return 'Hub site';
    if (tpl.startsWith('POINTPUBLISHINGTOPIC')) return 'Topic site (Viva)';
    if (tpl.startsWith('STS#3')) return 'Modern team site';
    if (tpl.startsWith('STS')) return 'Classic team site';
    if (tpl.startsWith('EHS')) return 'Modern team site (no group)';
    if (tpl.startsWith('APPCATALOG')) return 'App catalog';
    if (tpl.startsWith('SRCHCEN')) return 'Search center';
    if (tpl.startsWith('ENTERWIKI')) return 'Enterprise wiki';
    if (tpl.startsWith('BICENTERSITE')) return 'BI center';
    if (tpl.startsWith('BLANKINTERNET')) return 'Publishing site';
    if (tpl.startsWith('TENANTADMIN')) return 'Tenant admin';
    if (tpl) return tpl;
    return 'SharePoint';
  }

  function categoryOf(site) { return classifyType(site) === 'OneDrive' ? 'OneDrive' : 'SharePoint'; }

  // contoso.sharepoint.com -> the three hosts of that tenant. Anything else is refused,
  // so a token is never requested for a host that came back odd from an API.
  function hostsFromRoot(hostname) {
    const h = String(hostname || '').trim().toLowerCase();
    const m = /^([a-z0-9][a-z0-9-]*)\.sharepoint\.com$/.exec(h);
    if (!m || /-(my|admin)$/.test(m[1])) return null;
    return { tenant: h, my: `${m[1]}-my.sharepoint.com`, admin: `${m[1]}-admin.sharepoint.com` };
  }

  // Only https URLs on the tenant's own SharePoint or OneDrive host get a SharePoint token.
  function isTenantSiteUrl(url, hosts) {
    if (!hosts) return false;
    try {
      const u = new URL(url);
      const host = u.hostname.toLowerCase();
      return u.protocol === 'https:' && !u.username && !u.password && !u.port && (host === hosts.tenant || host === hosts.my);
    } catch { return false; }
  }

  function fmtDate(d) {
    if (!d) return '';
    const x = d instanceof Date ? d : new Date(d);
    return isNaN(x.getTime()) ? '' : x.toISOString().substring(0, 10);
  }

  const ownerName = (o) => o.title || o.email || o.loginName || '';

  // One flat row per result with every raw signal, so it can be re-scored in Excel.
  function exportRow(r) {
    const s = r.signals || {};
    return {
      Title: r.site.displayName || r.site.name || '',
      URL: r.site.webUrl || '',
      Type: classifyType(r.site),
      Owners: (r.owners || []).map(ownerName).filter(Boolean).join('; '),
      OwnerEmails: (r.owners || []).map((o) => o.email || '').filter(Boolean).join('; '),
      Template: (r.site.root && r.site.root['@odata.type']) || '',
      Created: fmtDate(r.site.createdDateTime || s.spoWebCreated),
      LastModified: fmtDate(r.lastModified),
      LastViewed: fmtDate(r.lastViewed),
      LastViewedApprox: r.lastViewedApprox ? 'Yes' : '',
      DaysIdle: r.daysSinceModified ?? '',
      PageViewsAllTime: r.pageViews ?? '',
      ViewsLast30Days: r.last30Views ?? '',
      ViewsLast7Days: r.recentViews ?? '',
      StorageMB: r.storageMB ?? '',
      Recommendation: r.advice,
      Why: r.reason,
      GraphLastModified: fmtDate(s.graphSiteLastModified),
      WebLastUserModified: fmtDate(s.spoWebLastUserModified),
      WebLastModified: fmtDate(s.spoWebLastModified),
      LibMaxUserModified: fmtDate(s.libMaxUserModified),
      LibCount: s.libCount ?? '',
      LibTotalItems: s.libTotalItems ?? '',
      LastPageEditDate: fmtDate(s.lastPageEditDate),
      LastPageEditName: s.lastPageEditName ?? '',
      LastPageEditBy: s.lastPageEditBy ?? '',
      AnalyticsAllTimeViews: s.analyticsAllTimeViews ?? '',
      AnalyticsAllTimeActors: s.analyticsAllTimeActors ?? '',
      AnalyticsRecentViews: s.analyticsRecentViews ?? '',
      AnalyticsRecentActors: s.analyticsRecentActors ?? '',
      Last30DayLastAccess: fmtDate(s.last30LastDay),
      Errors: (r.errors || []).join(' | '),
    };
  }

  const EXPORT_COLUMNS = [
    ['Title', 'Title'], ['URL', 'URL'], ['Type', 'Type'], ['Owners', 'Owners'], ['OwnerEmails', 'Owner emails'],
    ['Template', 'Template'], ['Created', 'Created'], ['LastModified', 'Last modified'], ['LastViewed', 'Last viewed'],
    ['LastViewedApprox', 'Last viewed approx'], ['DaysIdle', 'Days idle'], ['PageViewsAllTime', 'Page views all time'],
    ['ViewsLast30Days', 'Views last 30 days'], ['ViewsLast7Days', 'Views last 7 days'], ['StorageMB', 'Storage MB'],
    ['Recommendation', 'Recommendation'], ['Why', 'Why'], ['GraphLastModified', 'Graph last modified'],
    ['WebLastUserModified', 'Web last user modified'], ['WebLastModified', 'Web last modified'],
    ['LibMaxUserModified', 'Libraries last user modified'], ['LibCount', 'Library count'], ['LibTotalItems', 'Library items'],
    ['LastPageEditDate', 'Last page edit'], ['LastPageEditName', 'Last page edited'], ['LastPageEditBy', 'Last page edit by'],
    ['AnalyticsAllTimeViews', 'Analytics all time views'], ['AnalyticsAllTimeActors', 'Analytics all time actors'],
    ['AnalyticsRecentViews', 'Analytics 7 day views'], ['AnalyticsRecentActors', 'Analytics 7 day actors'],
    ['Last30DayLastAccess', 'Last access in 30 days'], ['Errors', 'Errors'],
  ].map(([key, label]) => ({ key, label }));

  window.Sites = { classifyType, categoryOf, hostsFromRoot, isTenantSiteUrl, fmtDate, ownerName, exportRow, EXPORT_COLUMNS };
})();
