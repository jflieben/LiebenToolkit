// Checks on URLs and ids from user input or tenant data, done before a token is sent
// anywhere or an id goes into a path.
(() => {
  const HOST = /^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)*$/;
  const GUID = '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}';
  const SITE_ID = new RegExp(`^[a-z0-9.-]+,${GUID},${GUID}$`, 'i');

  // https only, no credentials or port, hostname is a plain DNS name ending in suffix.
  function httpsUrlWithSuffix(value, suffix) {
    let u;
    try { u = new URL(String(value == null ? '' : value).trim()); } catch { return null; }
    if (u.protocol !== 'https:' || u.username || u.password || u.port) return null;
    const host = u.hostname.toLowerCase();
    if (!HOST.test(host) || !host.endsWith(suffix) || host.length <= suffix.length) return null;
    return u;
  }

  const sharePointUrl = (v) => httpsUrlWithSuffix(v, '.sharepoint.com');
  const dataverseUrl = (v) => httpsUrlWithSuffix(v, '.dynamics.com');

  function dataverseOrigin(v) {
    const u = dataverseUrl(v);
    return u ? u.origin : null;
  }

  // Azure DevOps: the profile/accounts host and org hosts under dev.azure.com.
  function isDevOpsUrl(v) {
    let u;
    try { u = new URL(String(v || '')); } catch { return false; }
    if (u.protocol !== 'https:' || u.username || u.password || u.port) return false;
    const host = u.hostname.toLowerCase();
    return host === 'app.vssps.visualstudio.com' || host === 'dev.azure.com' || (host.endsWith('.dev.azure.com') && HOST.test(host));
  }

  // Graph address for a site URL: /sites/{host}:/{path}: with every path segment encoded.
  function graphSitePath(u) {
    const segs = u.pathname.split('/').filter(Boolean).map((s) => {
      let plain = s;
      try { plain = decodeURIComponent(s); } catch { /* keep as typed */ }
      return encodeURIComponent(plain);
    });
    return segs.length ? `/sites/${u.hostname}:/${segs.join('/')}:` : `/sites/${u.hostname}`;
  }

  const isGraphSiteId = (v) => SITE_ID.test(String(v || ''));
  const sameSiteId = (a, b) => !!a && !!b && String(a).toLowerCase() === String(b).toLowerCase();

  // candidates: [{ groupId, rootSiteId }]. Returns the group whose root site is this site.
  function matchGroupSite(siteId, candidates) {
    const hit = (candidates || []).find((c) => c && Http.isGuid(c.groupId) && sameSiteId(c.rootSiteId, siteId));
    return hit ? hit.groupId : null;
  }

  window.Guards = {
    sharePointUrl, dataverseUrl, dataverseOrigin, isDevOpsUrl, graphSitePath,
    isGraphSiteId, sameSiteId, matchGroupSite,
  };
})();
