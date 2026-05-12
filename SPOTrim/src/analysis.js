// Per-site staleness analysis
(() => {
  function parseDate(s) {
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  function maxDate(...dates) {
    let m = null;
    for (const d of dates) if (d && (!m || d > m)) m = d;
    return m;
  }
  function daysBetween(a, b) {
    if (!a || !b) return null;
    return Math.floor((b - a) / 86400000);
  }
  function isAccessDenied(err) {
    if (!err) return false;
    if (err.status === 401 || err.status === 403) return true;
    const m = String(err.message || err);
    return /\b(401|403)\b/.test(m)
        || /unauthorized/i.test(m)
        || /access\s*denied/i.test(m)
        || /forbidden/i.test(m);
  }

  /**
   * Analyze a single site. Combines Graph site data (already on the site object),
   * SharePoint REST web/lists data, and per-site Graph engagement signals
   * (analytics, drive activities, last page edit, drive quota).
   */
  async function analyzeSite(site, opts = {}) {
    const { staleDays = 365 } = opts;
    const signals = {};
    const now = new Date();

    // 1. Graph site lastModifiedDateTime
    signals.graphSiteLastModified = parseDate(site.lastModifiedDateTime);

    // 2. SharePoint REST web (LastItemUserModifiedDate excludes system updates)
    let webInfo = null;
    let webErr = null;
    let webDenied = false;
    try {
      webInfo = await SharePoint.getWebInfo(site.webUrl);
      signals.spoWebLastUserModified = parseDate(webInfo.LastItemUserModifiedDate);
      signals.spoWebLastModified     = parseDate(webInfo.LastItemModifiedDate);
      signals.spoWebCreated          = parseDate(webInfo.Created);
    } catch (e) {
      webErr = e.message;
      webDenied = isAccessDenied(e);
      Log.dbg(`Web info failed for ${site.webUrl}: ${e.message}`);
    }

    // 3. SharePoint document libraries
    let libs = null;
    let libErr = null;
    let libDenied = false;
    try {
      libs = await SharePoint.getDocLibsLastModified(site.webUrl);
      const userModified = libs.map(l => parseDate(l.LastItemUserModifiedDate)).filter(Boolean);
      signals.libMaxUserModified = userModified.length ? maxDate(...userModified) : null;
      signals.libCount = libs.length;
      signals.libTotalItems = libs.reduce((s, l) => s + (l.ItemCount || 0), 0);
    } catch (e) {
      libErr = e.message;
      libDenied = isAccessDenied(e);
      Log.dbg(`Lib info failed for ${site.webUrl}: ${e.message}`);
    }

    // 3b. Site owners (best-effort; never fails the analysis)
    let owners = [];
    let ownersErr = null;
    try {
      owners = await SharePoint.getSiteOwners(site.webUrl);
      signals.ownerCount = owners.length;
    } catch (e) {
      ownersErr = e.message;
      Log.dbg(`Owner lookup failed for ${site.webUrl}: ${e.message}`);
    }

    // 4. Per-site Graph engagement: analytics (allTime + lastSevenDays),
    //    drive root activity (last 30 days), most recently edited site page,
    //    and current drive storage quota usage. One $batch, CORS-friendly.
    let engagementErr = null;
    if (site.id) {
      try {
        const e = await Graph.getSiteEngagement(site.id);
        if (e) {
          signals.analyticsAllTimeViews  = e.allTimeViews;
          signals.analyticsAllTimeActors = e.allTimeActors;
          signals.analyticsRecentViews   = e.recentViews;
          signals.analyticsRecentActors  = e.recentActors;
          signals.analyticsRecentEnd     = parseDate(e.recentEnd);
          signals.analyticsAllTimeEnd    = parseDate(e.allTimeEnd);
          signals.last30Views            = e.last30Views;
          signals.last30Actors           = e.last30Actors;
          signals.last30LastDay          = parseDate(e.last30LastDay);
          if (e.lastPageEdit) {
            signals.lastPageEditDate     = parseDate(e.lastPageEdit.date);
            signals.lastPageEditName     = e.lastPageEdit.name;
            signals.lastPageEditBy       = e.lastPageEdit.modifiedBy;
          }
          signals.quotaUsedBytes         = e.quotaUsedBytes;
          if (e._errors && e._errors.length) {
            engagementErr = e._errors.join('; ');
          }
        }
      } catch (e) {
        engagementErr = e.message;
        Log.dbg(`Site engagement failed for ${site.webUrl}: ${e.message}`);
      }
    }

    // Aggregate "last modified" = best (most recent) of user-modification signals
    const lastModified = maxDate(
      signals.graphSiteLastModified,
      signals.spoWebLastUserModified,
      signals.libMaxUserModified,
      signals.lastPageEditDate
    );
    // "Last viewed". Pick the freshest signal we have (more granular wins).
    let lastViewed = null;
    let lastViewedApprox = false;
    if (signals.last30LastDay) {
      lastViewed = signals.last30LastDay;
      lastViewedApprox = false; // bucketed daily, but accurate to within 1 day
    } else if (signals.analyticsRecentViews && signals.analyticsRecentEnd) {
      lastViewed = signals.analyticsRecentEnd;
      lastViewedApprox = true; // "at some point within the last 7 days"
    }

    // Combined page-view signal: prefer the broadest count we have (all-time)
    const pageViews = signals.analyticsAllTimeViews ?? null;
    const pageViewsSource = pageViews != null ? 'siteAnalyticsAllTime' : null;
    const recentViews = signals.analyticsRecentViews ?? null;
    const last30Views = signals.last30Views ?? null;

    const daysSinceModified = lastModified ? daysBetween(lastModified, now) : null;
    const daysSinceViewed   = lastViewed   ? daysBetween(lastViewed,   now) : null;

    // Recommendation
    let advice = 'Unknown';
    let reason = [];
    const haveModSignal = lastModified !== null;
    const haveViewSignal = lastViewed !== null || pageViews != null;

    if (haveModSignal) {
      if (daysSinceModified >= staleDays) {
        if (last30Views && last30Views > 0) {
          advice = 'Review';
          reason.push(`Not modified in ${daysSinceModified}d, but ${last30Views} access${last30Views === 1 ? '' : 'es'} in the last 30 days (still consulted/archive).`);
        } else if (recentViews && recentViews > 0) {
          advice = 'Review';
          reason.push(`Not modified in ${daysSinceModified}d, but ${recentViews} access${recentViews === 1 ? '' : 'es'} in the last 7 days (still consulted/archive).`);
        } else if (pageViews != null && pageViews > 0) {
          advice = 'Review';
          reason.push(`Not modified in ${daysSinceModified}d, but ${pageViews} all-time page view${pageViews === 1 ? '' : 's'}.`);
        } else if (haveViewSignal && daysSinceViewed !== null && daysSinceViewed >= staleDays) {
          advice = 'Clean up';
          reason.push(`Not modified in ${daysSinceModified}d and not viewed in ${daysSinceViewed}d.`);
        } else if (pageViews === 0) {
          advice = 'Clean up';
          reason.push(`Not modified in ${daysSinceModified}d and 0 all-time page views.`);
        } else {
          advice = 'Review';
          reason.push(`Not modified in ${daysSinceModified}d. View data unavailable.`);
        }
      } else {
        advice = 'Keep';
        reason.push(`Modified ${daysSinceModified}d ago (within threshold of ${staleDays}d).`);
      }
    } else {
      // No modification signal. Distinguish "permission denied" from other errors
      // so the operator can clearly see WHY a site wasn't analyzed.
      const allErrsAreDenied = (webDenied || !webErr) && (libDenied || !libErr) && (webDenied || libDenied);
      if (allErrsAreDenied) {
        advice = 'No access';
        reason.push('You do not have permission to read this site (HTTP 401/403). Add yourself as a site collection administrator (or grant Sites.FullControl.All app permission) to analyze it.');
        if (webErr) reason.push(`Web REST: ${webErr}`);
        if (libErr) reason.push(`Lists REST: ${libErr}`);
      } else {
        advice = 'Unknown';
        reason.push('No modification signal could be retrieved.');
        if (webErr) reason.push(`Web REST error: ${webErr}`);
        if (libErr) reason.push(`Lists REST error: ${libErr}`);
      }
    }

    return {
      site,
      signals,
      lastModified,
      lastViewed,
      lastViewedApprox,
      daysSinceModified,
      daysSinceViewed,
      pageViews,
      pageViewsSource,
      recentViews,
      last30Views,
      owners,
      storageMB: signals.quotaUsedBytes ? Math.round(signals.quotaUsedBytes / 1048576) : null,
      ...applySystemSiteOverride(site, advice, reason),
      errors: [webErr, libErr, engagementErr, ownersErr].filter(Boolean),
      accessDenied: (webDenied || libDenied),
    };
  }

  /**
   * Some sites are part of the M365/SharePoint platform and should never be
   * cleaned up - even if our heuristics say "stale". Removing them breaks
   * tenant-wide functionality (admin center, app catalog, search, MySite host,
   * Viva topic store, etc.).
   *
   * Detection uses two reliable signals from the discovery payload:
   *   1. The SharePoint web template captured in `site.root['@odata.type']`
   *   2. Well-known URL paths (root site, /search, /sites/appcatalog, ...)
   *      and the *-admin / *-my host names.
   *
   * Returns the (possibly overridden) advice + reason. We always KEEP these.
   */
  function applySystemSiteOverride(site, advice, reasonArr) {
    const why = isSystemSite(site);
    if (!why) return { advice, reason: reasonArr.join(' ') };
    return {
      advice: 'Keep',
      reason: `System site (${why}). SPOTrim never recommends cleaning these up.`,
    };
  }

  function isSystemSite(site) {
    const tpl = ((site.root && site.root['@odata.type']) || '').toUpperCase();
    const url = (site.webUrl || '').toLowerCase();

    // Template-based detection (most reliable).
    if (tpl.startsWith('TENANTADMIN'))         return 'tenant admin center';
    if (tpl.startsWith('SPSMSITEHOST'))        return 'OneDrive (MySite) host';
    if (tpl.startsWith('APPCATALOG'))          return 'app catalog';
    if (tpl.startsWith('SRCHCEN'))             return 'search center';
    if (tpl.startsWith('POINTPUBLISHINGHUB'))  return 'SharePoint home / start page';
    if (tpl.startsWith('POINTPUBLISHINGTOPIC'))return 'Viva Topics topic store';
    if (tpl.startsWith('SPSCOMMU'))            return 'community portal';

    // URL-based detection (covers cases where template is missing/odd).
    let host = '', path = '';
    try {
      const u = new URL(site.webUrl);
      host = u.hostname.toLowerCase();
      path = u.pathname.toLowerCase().replace(/\/+$/, '');
    } catch { return null; }

    if (/-admin\.sharepoint\.com$/i.test(host))    return 'SharePoint admin host';
    // The MySite host root itself (not individual /personal/* OneDrives).
    if (/-my\.sharepoint\.com$/i.test(host) && (path === '' || path === '/')) return 'OneDrive (MySite) host root';

    // Tenant root site.
    if (/^[a-z0-9-]+\.sharepoint\.com$/i.test(host) && (path === '' || path === '/')) return 'tenant root site';

    // Well-known managed paths that ship with every tenant.
    const wellKnown = {
      '/search': 'search center',
      '/sites/appcatalog': 'app catalog',
      '/sites/contenttypehub': 'content type hub',
      '/sites/compliancepolicycenter': 'compliance policy center',
      '/sites/recordscenter': 'records center',
      '/sites/communityportal': 'community portal',
      '/sites/mysitehost': 'mysite host',
    };
    if (wellKnown[path]) return wellKnown[path];

    return null;
  }

  window.Analysis = { analyzeSite, isSystemSite };
})();
