// SharePoint tenant admin API for SPNRoleMgr: list every site collection, and make the
// signed-in user site collection admin for a while. Needs SharePoint Administrator.
(() => {
  const V1 = 'https://graph.microsoft.com/v1.0';
  let hosts = null; // { spo, admin }

  // SharePoint host from the root site; the initial domain when that is not readable.
  async function resolveHosts() {
    if (hosts) return hosts;
    let spo = null;
    try {
      const root = await Graph.call(`${V1}/sites/root?$select=siteCollection,webUrl`, { kind: 'sites' });
      spo = (root && root.siteCollection && root.siteCollection.hostname) || (root && root.webUrl && new URL(root.webUrl).hostname) || null;
    } catch (e) { Log.warn(`Root site not readable, using the initial domain: ${e.message}`); }
    if (!spo) {
      const org = await Graph.call(`${V1}/organization?$select=verifiedDomains`);
      const domains = (org && org.value && org.value[0] && org.value[0].verifiedDomains) || [];
      const prefix = ((domains.find((d) => d.isInitial) || {}).name || '').split('.')[0];
      if (/^[a-z0-9]+$/i.test(prefix)) spo = `${prefix.toLowerCase()}.sharepoint.com`;
    }
    const admin = Spns.adminHostOf(spo);
    if (!spo || !Spns.isSharePointHost(spo) || !admin) throw new Error('Could not find the SharePoint host of this tenant.');
    hosts = { spo: spo.toLowerCase(), admin };
    Log.info(`SharePoint host ${hosts.spo}, admin host ${hosts.admin}`);
    return hosts;
  }

  // noRedirect: a blocked popup must not navigate away from the open dialog.
  const adminToken = Graph.guarded(async () => {
    const { admin } = await resolveHosts();
    return Auth.getToken([`https://${admin}/.default`], { noRedirect: true });
  });
  const opts = (extra) => ({ label: 'SPO-admin', token: adminToken, ...extra });

  async function formDigest() {
    const { admin } = await resolveHosts();
    const data = await Http.json(`https://${admin}/_api/contextinfo`, opts({ method: 'POST', headers: { Accept: 'application/json;odata=verbose' } }));
    const v = data && ((data.d && data.d.GetContextWebInformation && data.d.GetContextWebInformation.FormDigestValue) || data.FormDigestValue);
    if (!v) throw new Error('No form digest from the SharePoint admin API.');
    return v;
  }

  const SKIP_TEMPLATES = new Set(['redirectsite#0', 'spsmsitehost#0', 'srchcen#0', 'pointpublishinghub#0', 'pointpublishingtopic#0']);

  // Every site collection, also those the user is not a member of. OneDrive excluded.
  async function listSites(onProgress) {
    const { admin } = await resolveHosts();
    const endpoint = `https://${admin}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/GetSitePropertiesFromSharePointByFilters`;
    const digest = await formDigest();
    const out = new Map();
    let startIndex = null;
    for (let guard = 0; guard < 1000; guard++) {
      const speFilter = {
        __metadata: { type: 'Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerableFilter' },
        IncludePersonalSite: 0, StartIndex: startIndex, IncludeDetail: false, GroupIdDefined: 0, Template: '', Filter: '',
      };
      const data = await Http.json(endpoint, opts({
        method: 'POST',
        headers: { 'X-RequestDigest': digest, 'Content-Type': 'application/json;odata=verbose', Accept: 'application/json;odata=verbose' },
        body: JSON.stringify({ speFilter }),
      }));
      const d = (data && data.d && (data.d.GetSitePropertiesFromSharePointByFilters || data.d)) || data || {};
      for (const s of (d.results || d.value || [])) {
        const key = Spns.siteKey(s.Url);
        if (!key || SKIP_TEMPLATES.has(String(s.Template || '').toLowerCase())) continue;
        out.set(key, { id: null, displayName: s.Title || s.Url, webUrl: String(s.Url).replace(/\/+$/, ''), template: s.Template || '' });
      }
      if (onProgress) onProgress(out.size);
      const next = d.NextStartIndexFromSharePoint;
      if (next == null || next === '' || next === '0') break;
      startIndex = String(next);
    }
    return [...out.values()].sort((a, b) => String(a.displayName).localeCompare(String(b.displayName)));
  }

  function xmlEscape(s) {
    return String(s || '').replace(/[<>&"']/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' }[c]));
  }

  // Tenant.SetSiteAdmin(siteUrl, loginName, isSiteAdmin) through CSOM, as SPOMove does.
  async function setSiteAdmin(siteUrl, upn, isSiteAdmin) {
    const { spo, admin } = await resolveHosts();
    const p = Spns.parseSharePointUrl(siteUrl);
    if (!p || p.host !== spo) throw new Error(`${siteUrl} is not a site of this tenant (${spo}).`);
    const login = Spns.claimLoginName(upn);
    const digest = await formDigest().catch(() => null);
    const xml = '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="SPNRoleMgr">' +
      '<Actions><Method Name="SetSiteAdmin" Id="1" ObjectPathId="2"><Parameters>' +
      `<Parameter Type="String">${xmlEscape(siteUrl)}</Parameter>` +
      `<Parameter Type="String">${xmlEscape(login)}</Parameter>` +
      `<Parameter Type="Boolean">${isSiteAdmin ? 'true' : 'false'}</Parameter>` +
      '</Parameters></Method></Actions>' +
      '<ObjectPaths><Constructor Id="2" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>';
    const headers = { 'Content-Type': 'text/xml', Accept: '*/*' };
    if (digest) headers['X-RequestDigest'] = digest;
    Log.info(`SetSiteAdmin(${siteUrl}, ${login}, ${isSiteAdmin})`);
    const res = await Http.request(`https://${admin}/_vti_bin/client.svc/ProcessQuery`, opts({ method: 'POST', headers, body: xml }));
    const body = await res.json().catch(() => null);
    if (Array.isArray(body) && body[0] && body[0].ErrorInfo) {
      const ei = body[0].ErrorInfo;
      const e = new Error(`SetSiteAdmin failed: ${ei.ErrorMessage || 'unknown error'}`);
      if (String(ei.ErrorCode) === '-2147024891' || /unauthorized|access denied/i.test(ei.ErrorMessage || '')) e.status = 403;
      throw e;
    }
  }

  window.Spo = { resolveHosts, listSites, setSiteAdmin };
})();
