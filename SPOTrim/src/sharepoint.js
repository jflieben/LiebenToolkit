// SharePoint REST calls: per-site modification data, owners, search based discovery,
// OneDrive listing and the SetSiteAdmin call used by force scan.
// Tokens are taken silently; app.js asks consent (Auth.consent) before a run starts.
(() => {
  let hosts = null; // { tenant, my, admin } from Sites.hostsFromRoot

  function setHosts(h) { hosts = h; }
  function getHosts() { return hosts; }
  function need() {
    if (!hosts) throw new Error('SharePoint host not detected yet');
    return hosts;
  }
  // Read access to sites, for analysis and search.
  const readScopes = () => [`https://${need().tenant}/AllSites.Read`];
  // Tenant admin API (OneDrive listing, SetSiteAdmin). Needs the SharePoint Administrator role too.
  const adminScopes = () => [`https://${need().admin}/AllSites.FullControl`];
  const readToken = () => Auth.getToken(readScopes(), { interactive: false, noRedirect: true });
  const adminToken = () => Auth.getToken(adminScopes(), { interactive: false, noRedirect: true });

  function siteBase(siteUrl) {
    if (!Sites.isTenantSiteUrl(siteUrl, hosts)) throw new Error(`Not a SharePoint URL of this tenant: ${siteUrl}`);
    return String(siteUrl).replace(/\/+$/, '');
  }

  function rest(siteUrl, path) {
    return Http.json(siteBase(siteUrl) + path, { label: 'SPO', token: readToken, headers: { Accept: 'application/json;odata=nometadata' } });
  }

  function getWebInfo(siteUrl) {
    return rest(siteUrl, '/_api/web?$select=Title,Url,Created,LastItemUserModifiedDate,LastItemModifiedDate');
  }

  // Non-hidden document libraries with their last user modification.
  async function getDocLibsLastModified(siteUrl) {
    const data = await rest(siteUrl, '/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 101&$select=Title,LastItemUserModifiedDate,LastItemModifiedDate,ItemCount');
    return (data && data.value) || [];
  }

  // Owners: members of the associated owner group, or the primary site owner when that
  // group only holds a system principal (group connected sites, OneDrive).
  async function getSiteOwners(siteUrl) {
    const seen = new Map();
    const add = (title, email, loginName) => {
      if (!title && !email) return;
      const key = (email || loginName || title).toLowerCase();
      if (seen.has(key)) return;
      const ln = String(loginName || '');
      if (/_o$|_m$|^c:0\(\.s\|true|^SHAREPOINT\\system$/i.test(ln)) return;
      if (title && /^Everyone(?:\s+except external users)?$/i.test(title)) return;
      seen.set(key, { title: title || '', email: email || '', loginName: ln });
    };
    try {
      const data = await rest(siteUrl, '/_api/web/AssociatedOwnerGroup/Users?$select=Title,Email,LoginName');
      for (const u of (data && data.value) || []) add(u.Title, u.Email, u.LoginName);
    } catch (e) {
      Log.dbg(`Owner group lookup failed for ${siteUrl}: ${e.message}`);
    }
    if (seen.size === 0) {
      try {
        const data = await rest(siteUrl, '/_api/site?$expand=Owner&$select=Owner/Title,Owner/Email,Owner/LoginName');
        const o = data && data.Owner;
        if (o) add(o.Title, o.Email, o.LoginName);
      } catch (e) {
        Log.dbg(`Site owner lookup failed for ${siteUrl}: ${e.message}`);
      }
    }
    return [...seen.values()];
  }

  // All site collections the signed-in user can see, via SharePoint Search.
  async function searchAllSites({ includeOneDrive = true, includeSharePoint = true, onProgress, signal } = {}) {
    const url = `https://${need().tenant}/_api/search/postquery`;
    let qt;
    if (includeOneDrive && includeSharePoint) qt = 'contentclass:STS_Site';
    else if (includeOneDrive) qt = 'contentclass:STS_Site WebTemplate:SPSPERS';
    else qt = 'contentclass:STS_Site -WebTemplate:SPSPERS';

    const select = ['Title', 'Path', 'SiteId', 'WebTemplate', 'WebId', 'CreatedBy', 'LastModifiedTime', 'Created', 'SiteName'];
    const rowLimit = 500;
    let startRow = 0;
    const out = [];
    const seen = new Set();
    for (;;) {
      const data = await Http.json(url, {
        method: 'POST', label: 'SPO search', token: readToken, signal,
        headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' },
        body: {
          request: {
            __metadata: { type: 'Microsoft.Office.Server.Search.REST.SearchRequest' },
            Querytext: qt, RowLimit: rowLimit, StartRow: startRow, TrimDuplicates: false,
            SelectProperties: { results: select }, ClientType: 'SPOTrim',
          },
        },
      });
      const root = (data && data.d && data.d.postquery) || (data && data.PostqueryResult) || data || {};
      const rels = root.PrimaryQueryResult && root.PrimaryQueryResult.RelevantResults;
      if (!rels) break;
      const rows = (rels.Table && rels.Table.Rows && (rels.Table.Rows.results || rels.Table.Rows)) || [];
      for (const row of rows) {
        const obj = {};
        for (const c of (row.Cells && (row.Cells.results || row.Cells)) || []) obj[c.Key] = c.Value;
        const path = obj.Path || '';
        if (!path || seen.has(path.toLowerCase())) continue;
        seen.add(path.toLowerCase());
        const isPersonal = (obj.WebTemplate || '').toUpperCase().startsWith('SPSPERS') || /-my\.sharepoint\.com\/personal\//i.test(path);
        out.push({
          id: obj.SiteId || obj.WebId || path,
          displayName: obj.Title || obj.SiteName || '',
          name: obj.SiteName || '',
          webUrl: path,
          createdDateTime: obj.Created || null,
          lastModifiedDateTime: obj.LastModifiedTime || null,
          isPersonalSite: isPersonal,
          root: { '@odata.type': obj.WebTemplate || '' },
          _source: 'search',
        });
      }
      if (onProgress) onProgress(out.length, rels.TotalRows || rows.length || 0);
      if (rows.length < rowLimit) break;
      startRow += rowLimit;
      if (startRow > 100000) break;
    }
    return out;
  }

  // Every OneDrive in the tenant via the tenant admin API. Search only returns the
  // OneDrives the caller can open; this one returns all of them.
  async function listOneDriveSitesViaTenantApi({ signal } = {}) {
    const url = `https://${need().admin}/_api/SPO.Tenant/GetSitePropertiesFromSharePointByFilters`;
    const out = [];
    const seen = new Set();
    let startIndex = null;
    for (let page = 1; ; page++) {
      let data;
      try {
        data = await Http.json(url, {
          method: 'POST', label: 'SPO admin', token: adminToken, signal,
          headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
          body: { speFilter: { IncludePersonalSite: 1, Template: '', StartIndex: startIndex, IncludeDetail: false } },
        });
      } catch (e) {
        if (e.status === 401 || e.status === 403) throw new Error(`Tenant admin API refused access (HTTP ${e.status}). You need the SharePoint Administrator role.`);
        throw e;
      }
      const props = (data && (data.value || data.Items || (data.SiteProperties && data.SiteProperties.results))) || [];
      let added = 0;
      for (const sp of props) {
        const u = sp.Url || sp.URL || '';
        if (!u) continue;
        const tpl = String(sp.Template || '');
        if (!(tpl.toUpperCase().startsWith('SPSPERS') || /-my\.sharepoint\.com\/personal\//i.test(u))) continue;
        const k = u.toLowerCase().replace(/\/+$/, '');
        if (seen.has(k)) continue;
        seen.add(k);
        out.push({
          id: sp.SiteId || k,
          displayName: sp.Title || sp.Owner || '',
          name: sp.Title || '',
          webUrl: u.replace(/\/+$/, ''),
          createdDateTime: null,
          lastModifiedDateTime: sp.LastContentModifiedDate || null,
          isPersonalSite: true,
          root: { '@odata.type': tpl || 'SPSPERS#10' },
          _source: 'spo-admin-tenant-api',
          _ownerUpn: sp.Owner || sp.OwnerEmail || '',
          _storageUsageMB: sp.StorageUsage || 0,
        });
        added++;
      }
      Log.dbg(`SP admin page ${page}: ${props.length} items, ${added} OneDrives added (total ${out.length})`);
      const next = data && data.NextStartIndexFromSharePoint;
      if (!next || next === startIndex) break;
      startIndex = next;
      if (page >= 200) { Log.warn('SP admin paging stopped at 200 pages'); break; }
    }
    return out;
  }

  // ---------- force scan ----------
  const XML_ESC = { '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' };
  const xmlEscape = (s) => String(s == null ? '' : s).replace(/[<>&"']/g, (c) => XML_ESC[c]);

  function buildClaimLoginName(upn) {
    if (!upn) throw new Error('UPN required to build the claim login name');
    return /^i:0/i.test(upn) ? upn : `i:0#.f|membership|${upn}`;
  }

  // Resolves with the currentuser JSON, rejects with the HttpError (status kept) otherwise.
  function getCurrentUser(siteUrl) {
    return rest(siteUrl, '/_api/web/currentuser?$select=IsSiteAdmin,LoginName');
  }

  function setSiteAdminXml(siteUrl, claim, on) {
    return '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="SPOTrim">' +
      '<Actions><Method Name="SetSiteAdmin" Id="1" ObjectPathId="2"><Parameters>' +
      `<Parameter Type="String">${xmlEscape(siteUrl)}</Parameter>` +
      `<Parameter Type="String">${xmlEscape(claim)}</Parameter>` +
      `<Parameter Type="Boolean">${on ? 'true' : 'false'}</Parameter>` +
      '</Parameters></Method></Actions>' +
      '<ObjectPaths><Constructor Id="2" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>';
  }

  // Add (on=true) or remove a site collection admin through tenant admin CSOM.
  async function setSiteAdmin(siteUrl, upn, on) {
    const site = siteBase(siteUrl);
    const claim = buildClaimLoginName(upn);
    Log.dbg(`SPO SetSiteAdmin(${site}, ${claim}, ${on})`);
    const body = await Http.json(`https://${need().admin}/_vti_bin/client.svc/ProcessQuery`, {
      method: 'POST', label: 'SPO admin', token: adminToken,
      headers: { 'Content-Type': 'text/xml', Accept: 'application/json' },
      body: setSiteAdminXml(site, claim, on),
    });
    const info = Array.isArray(body) && body[0] ? body[0].ErrorInfo : null;
    if (info) {
      const e = new Error(`SetSiteAdmin: ${info.ErrorMessage || JSON.stringify(info)}`);
      e.csom = true;
      throw e;
    }
    // Not marked csom: an odd answer does not prove the change was refused.
    if (!Array.isArray(body)) throw new Error('SetSiteAdmin: unexpected response');
    return true;
  }

  window.SharePoint = {
    setHosts, getHosts, readScopes, adminScopes,
    getWebInfo, getDocLibsLastModified, getSiteOwners,
    searchAllSites, listOneDriveSitesViaTenantApi,
    getCurrentUser, setSiteAdmin, buildClaimLoginName, setSiteAdminXml,
  };
})();
