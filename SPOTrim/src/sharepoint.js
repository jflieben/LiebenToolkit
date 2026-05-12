// SharePoint REST client (per-site requests for LastItemUserModifiedDate, etc.)
(() => {
  async function callSpoRest(siteUrl, relativePath, { attempt = 0 } = {}) {
    const token = await Auth.getSpoToken();
    const url = siteUrl.replace(/\/$/, '') + relativePath;
    Log.dbg(`SPO  GET ${url}`);
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json;odata=nometadata',
      },
    });
    if (res.status === 429 || res.status === 503 || res.status === 504) {
      const ra = parseInt(res.headers.get('Retry-After') || '0', 10);
      const wait = (ra > 0 ? ra : Math.min(60, Math.pow(2, attempt))) * 1000;
      if (attempt >= 4) throw new Error(`SPO ${res.status} after retries: ${url}`);
      Log.warn(`SPO ${res.status} → backing off ${wait}ms (attempt ${attempt + 1})`);
      await Concurrency.sleep(wait);
      return callSpoRest(siteUrl, relativePath, { attempt: attempt + 1 });
    }
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      const e = new Error(`SPO ${res.status}: ${txt.substring(0, 200)}`);
      e.status = res.status;
      throw e;
    }
    return res.json();
  }

  async function getWebInfo(siteUrl) {
    return callSpoRest(siteUrl, "/_api/web?$select=Title,Url,Created,LastItemUserModifiedDate,LastItemModifiedDate");
  }

  async function getDocLibsLastModified(siteUrl) {
    // Query non-hidden document libraries for their LastItemUserModifiedDate
    const path = "/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 101&$select=Title,LastItemUserModifiedDate,LastItemModifiedDate,ItemCount";
    const data = await callSpoRest(siteUrl, path);
    return data.value || [];
  }

  /**
   * Get site owners. For most SharePoint sites this is the membership of the
   * Associated Owner Group. For Microsoft 365 group-connected sites the owner
   * group typically only contains a system principal (groupId_o), so as a
   * fallback we also query the primary site owner via /_api/site?$expand=Owner.
   * For OneDrive personal sites the primary site owner is the user themselves.
   *
   * Returns: array of { title, email, loginName }. Empty array on permission
   * denied or no data. Throws on transient errors only.
   */
  async function getSiteOwners(siteUrl) {
    const seen = new Map(); // key = email|loginName lowercased
    const add = (title, email, loginName) => {
      if (!title && !email) return;
      const key = (email || loginName || title).toLowerCase();
      if (seen.has(key)) return;
      // Skip pure system principals (group_o, SharePoint App, etc.)
      const ln = String(loginName || '');
      if (/_o$|_m$|^c:0\(\.s\|true|^SHAREPOINT\\system$/i.test(ln)) return;
      if (title && /^Everyone(?:\s+except external users)?$/i.test(title)) return;
      seen.set(key, { title: title || '', email: email || '', loginName: ln });
    };

    // Owner group members
    try {
      const data = await callSpoRest(siteUrl, '/_api/web/AssociatedOwnerGroup/Users?$select=Title,Email,LoginName');
      for (const u of (data.value || [])) add(u.Title, u.Email, u.LoginName);
    } catch (e) {
      // Silently ignore: many sites have no owner group readable to the caller.
      Log.dbg(`Owner group lookup failed for ${siteUrl}: ${e.message}`);
    }

    // Primary site owner (works for OneDrive + group-connected)
    if (seen.size === 0) {
      try {
        const data = await callSpoRest(siteUrl, '/_api/site?$expand=Owner&$select=Owner/Title,Owner/Email,Owner/LoginName');
        const o = data.Owner;
        if (o) add(o.Title, o.Email, o.LoginName);
      } catch (e) {
        Log.dbg(`Site owner lookup failed for ${siteUrl}: ${e.message}`);
      }
    }

    return [...seen.values()];
  }

  // Enumerate all site collections via SharePoint Search (works for OneDrive too).
  // Uses the user's SharePoint token. Requires AllSites.Read or higher.
  async function searchAllSites({ includeOneDrive = true, includeSharePoint = true, onProgress } = {}) {
    const tenantHost = Auth.getTenantHost();
    if (!tenantHost) throw new Error('Tenant host not detected');
    const baseUrl = `https://${tenantHost}`;

    // Build query: site contentclass; optionally restrict to SPSPERS for OneDrive only
    let qt;
    if (includeOneDrive && includeSharePoint) qt = 'contentclass:STS_Site';
    else if (includeOneDrive) qt = 'contentclass:STS_Site WebTemplate:SPSPERS';
    else qt = 'contentclass:STS_Site -WebTemplate:SPSPERS';

    const select = ['Title', 'Path', 'SiteId', 'WebTemplate', 'WebId', 'CreatedBy', 'LastModifiedTime', 'Created', 'SiteName'];
    const rowLimit = 500;
    let startRow = 0;
    const out = [];
    const seen = new Set();

    while (true) {
      const body = {
        request: {
          __metadata: { type: 'Microsoft.Office.Server.Search.REST.SearchRequest' },
          Querytext: qt,
          RowLimit: rowLimit,
          StartRow: startRow,
          TrimDuplicates: false,
          SelectProperties: { results: select },
          ClientType: 'SPOTrim',
        },
      };
      const token = await Auth.getSpoToken();
      const url = `${baseUrl}/_api/search/postquery`;
      Log.dbg(`SPO  POST ${url} startRow=${startRow}`);
      const res = await fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
        },
        body: JSON.stringify(body),
      });
      if (res.status === 429 || res.status === 503) {
        const ra = parseInt(res.headers.get('Retry-After') || '5', 10) * 1000;
        Log.warn(`Search throttled, waiting ${ra}ms`);
        await Concurrency.sleep(ra);
        continue;
      }
      if (!res.ok) {
        const txt = await res.text().catch(() => '');
        throw new Error(`Search ${res.status}: ${txt.substring(0, 300)}`);
      }
      const data = await res.json();
      // odata=verbose wraps in d.postquery
      const root = data.d?.postquery || data.PostqueryResult || data;
      const rels = root.PrimaryQueryResult?.RelevantResults;
      if (!rels) break;
      const rows = rels.Table?.Rows?.results || rels.Table?.Rows || [];
      for (const row of rows) {
        const cells = row.Cells?.results || row.Cells || [];
        const obj = {};
        for (const c of cells) obj[c.Key] = c.Value;
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
      const totalRows = rels.TotalRows || rels.Table?.Rows?.length || 0;
      if (onProgress) onProgress(out.length, totalRows);
      if (rows.length < rowLimit) break;
      startRow += rowLimit;
      if (startRow > 100000) break; // safety
    }
    return out;
  }

  /**
   * Enumerate OneDrive sites via the SharePoint Admin Tenant REST API.
   * Requires the signed-in user to be a SharePoint Administrator.
   * Unlike Graph reports, this endpoint returns CORS headers and does not redirect cross-origin,
   * so it works from a browser SPA. Unlike SharePoint Search, it returns ALL personal sites
   * in the tenant (not just sites the user can access).
   *
   * Endpoint: POST {tenant}-admin.sharepoint.com/_api/SPO.Tenant/GetSitePropertiesFromSharePointByFilters
   * Filter: IncludePersonalSite=1, Template='' (any), IncludeDetail=false
   */
  async function listOneDriveSitesViaTenantApi() {
    const adminHost = Auth.getAdminHost();
    if (!adminHost) throw new Error('Admin host not detected (tenant not connected?)');
    const baseUrl = `https://${adminHost}`;
    const url = `${baseUrl}/_api/SPO.Tenant/GetSitePropertiesFromSharePointByFilters`;

    const out = [];
    let startIndex = null;
    let pageNum = 0;
    const seen = new Set();

    while (true) {
      pageNum++;
      const body = {
        speFilter: {
          IncludePersonalSite: 1,         // 1 = include OneDrives
          Template: '',                   // any template (we'll filter to SPSPERS)
          StartIndex: startIndex,         // pagination cursor
          IncludeDetail: false,
        },
      };
      const token = await Auth.getSpoAdminToken();
      Log.dbg(`SPO  POST ${url} (page ${pageNum}, startIndex=${startIndex})`);
      const res = await fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
        },
        body: JSON.stringify(body),
      });
      if (res.status === 401 || res.status === 403) {
        const txt = await res.text().catch(() => '');
        throw new Error(`SP Admin tenant API access denied (HTTP ${res.status}). The signed-in user must be a SharePoint Administrator. ${txt.substring(0, 200)}`);
      }
      if (res.status === 429 || res.status === 503) {
        const ra = parseInt(res.headers.get('Retry-After') || '5', 10) * 1000;
        Log.warn(`SP Admin throttled, waiting ${ra}ms`);
        await Concurrency.sleep(ra);
        continue;
      }
      if (!res.ok) {
        const txt = await res.text().catch(() => '');
        throw new Error(`SP Admin tenant API HTTP ${res.status}: ${txt.substring(0, 300)}`);
      }
      const data = await res.json();
      // Response shape (nometadata): { value: [ { Url, Template, Title, LastContentModifiedDate, StorageUsage, ...}, ... ], NextStartIndexFromSharePoint?: ... }
      const props = data.value || data.Items || data.SiteProperties?.results || [];
      let added = 0;
      for (const sp of props) {
        const u = sp.Url || sp.URL || '';
        if (!u) continue;
        const tpl = (sp.Template || '').toString();
        const isPersonal = tpl.toUpperCase().startsWith('SPSPERS') || /-my\.sharepoint\.com\/personal\//i.test(u);
        if (!isPersonal) continue;        // SPO admin returns ALL site types - we only want OneDrives here
        const k = u.toLowerCase().replace(/\/+$/, '');
        if (seen.has(k)) continue;
        seen.add(k);
        out.push({
          id: sp.SiteId || k,
          displayName: sp.Title || sp.Owner || '',
          name: sp.Title || '',
          webUrl: u.replace(/\/+$/, ''),
          createdDateTime: sp.SiteOwnerName || null, // not provided
          lastModifiedDateTime: sp.LastContentModifiedDate || null,
          isPersonalSite: true,
          root: { '@odata.type': tpl || 'SPSPERS#10' },
          _source: 'spo-admin-tenant-api',
          _ownerUpn: sp.Owner || sp.OwnerEmail || '',
          _storageUsageMB: sp.StorageUsage || 0,
        });
        added++;
      }
      Log.dbg(`SP Admin page ${pageNum}: ${props.length} items, ${added} OneDrives added (running total ${out.length})`);

      // Pagination: NextStartIndexFromSharePoint is included when more rows exist
      const next = data.NextStartIndexFromSharePoint || data.value?.NextStartIndexFromSharePoint;
      if (!next || next === startIndex) break;
      startIndex = next;
      if (pageNum > 200) { Log.warn('SP Admin pagination safety stop at 200 pages'); break; }
    }
    return out;
  }

  // ---------- Force-scan helpers: temporary site collection admin ----------
  // The SetSiteAdmin call goes through SharePoint's CSOM ProcessQuery on the
  // tenant admin endpoint. The signed-in user must hold the SharePoint
  // Administrator role in Entra ID. Login name format expected by SetSiteAdmin
  // is the SharePoint claim, e.g. "i:0#.f|membership|user@tenant.com".
  function _xmlEscape(s) {
    return String(s || '').replace(/[<>&"']/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c]));
  }
  function buildClaimLoginName(upn) {
    if (!upn) throw new Error('UPN required to build claim login name');
    if (/^i:0/i.test(upn)) return upn;
    return `i:0#.f|membership|${upn}`;
  }

  /**
   * Check whether the signed-in user is currently a site collection admin
   * for the given site. Returns:
   *   { hasAccess: bool, isSiteAdmin: bool, error?: string }
   * hasAccess=false means the call returned 401/403 (no access at all).
   */
  async function getCurrentUserAdminStatus(siteUrl) {
    try {
      const data = await callSpoRest(siteUrl, '/_api/web/currentuser?$select=IsSiteAdmin,LoginName');
      return { hasAccess: true, isSiteAdmin: !!data.IsSiteAdmin, loginName: data.LoginName || '' };
    } catch (e) {
      if (e.status === 401 || e.status === 403) return { hasAccess: false, isSiteAdmin: false };
      return { hasAccess: false, isSiteAdmin: false, error: e.message };
    }
  }

  /**
   * Add or remove a site collection admin via the SP Admin tenant CSOM endpoint.
   * isSiteAdmin=true adds, false removes.
   */
  async function setSiteAdmin(siteUrl, claimLoginName, isSiteAdmin) {
    const adminHost = Auth.getAdminHost();
    if (!adminHost) throw new Error('Admin host not detected');
    const url = `https://${adminHost}/_vti_bin/client.svc/ProcessQuery`;
    const token = await Auth.getSpoAdminToken();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="SPOTrim">` +
      `<Actions>` +
      `<Method Name="SetSiteAdmin" Id="1" ObjectPathId="2">` +
      `<Parameters>` +
      `<Parameter Type="String">${_xmlEscape(siteUrl)}</Parameter>` +
      `<Parameter Type="String">${_xmlEscape(claimLoginName)}</Parameter>` +
      `<Parameter Type="Boolean">${isSiteAdmin ? 'true' : 'false'}</Parameter>` +
      `</Parameters>` +
      `</Method>` +
      `</Actions>` +
      `<ObjectPaths>` +
      `<Constructor Id="2" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" />` +
      `</ObjectPaths>` +
      `</Request>`;

    Log.dbg(`SPO  ProcessQuery SetSiteAdmin(${siteUrl}, ${claimLoginName}, ${isSiteAdmin})`);
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'text/xml',
        Accept: '*/*',
      },
      body: xml,
    });
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      throw new Error(`SetSiteAdmin HTTP ${res.status}: ${txt.substring(0, 300)}`);
    }
    const body = await res.json().catch(() => null);
    // CSOM returns an array; the first element is { SchemaVersion, LibraryVersion, ErrorInfo, ... }
    if (Array.isArray(body) && body[0] && body[0].ErrorInfo) {
      const err = body[0].ErrorInfo;
      throw new Error(`SetSiteAdmin error: ${err.ErrorMessage || JSON.stringify(err)}`);
    }
    return true;
  }

  window.SharePoint = {
    getWebInfo, getDocLibsLastModified, getSiteOwners,
    searchAllSites, listOneDriveSitesViaTenantApi,
    getCurrentUserAdminStatus, setSiteAdmin, buildClaimLoginName,
  };
})();
