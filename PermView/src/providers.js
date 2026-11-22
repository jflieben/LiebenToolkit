// Workload providers. Every provider has the same shape so app.js renders them all
// with one form and one results table:
//
//   id, name, icon (i-*), description, notice (the quick-view limits),
//   scopes: [what is asked the first time it runs], atSignIn: true when nothing extra is asked,
//   columns?: [{ key, label }], quickFilters?: [{ id, label, predicate }],
//   form: [{ kind: 'text'|'select', name, label, placeholder?, loadOptions? }],
//   run(values, ctx): streams rows via ctx.onRow(row), ctx.onProgress({ label, current?, total? }),
//                     ctx.setEntity({ name, sub }); ctx.cancelToken stops long runs.
(() => {
  const G = Api.Graph;
  const enc = encodeURIComponent;

  function row({ principal, principalType, role, scope, source, ...rest }) {
    return {
      principal: principal || '(unknown)',
      principalType: principalType || '',
      role: role || '',
      scope: scope || '',
      source: source || '',
      ...rest,
    };
  }

  const odataType = (o) => String((o && o['@odata.type']) || '').replace('#microsoft.graph.', '');

  // Object ids to { displayName, type, upn } via getByIds (max 1000 per call).
  async function resolveDirectoryObjects(ids) {
    const out = new Map();
    const unique = [...new Set(ids.filter((id) => Http.isGuid(id)))];
    for (let i = 0; i < unique.length; i += 1000) {
      try {
        const data = await G('read').json('/directoryObjects/getByIds', {
          method: 'POST',
          body: { ids: unique.slice(i, i + 1000), types: ['user', 'group', 'servicePrincipal', 'device'] },
        });
        for (const obj of (data && data.value) || []) {
          out.set(obj.id, {
            displayName: obj.displayName || obj.userPrincipalName || obj.appDisplayName || obj.id,
            type: odataType(obj),
            upn: obj.userPrincipalName || obj.appId || '',
          });
        }
      } catch (e) {
        Log.warn(`Could not resolve directory objects: ${e.message}`);
      }
    }
    return out;
  }

  function summarizeCredentials(passwordCredentials, keyCredentials, now = Date.now()) {
    const nearDays = 30;
    const creds = [...(passwordCredentials || []), ...(keyCredentials || [])].filter((c) => c && c.endDateTime);
    if (!creds.length) return { state: 'None', severity: 'none', expiryText: 'n/a' };
    const expiries = creds.map((c) => Date.parse(c.endDateTime)).filter(Number.isFinite).sort((a, b) => a - b);
    if (!expiries.length) return { state: 'Unknown', severity: 'medium', expiryText: 'unknown' };
    const next = expiries[0];
    const days = Math.ceil((next - now) / 86400000);
    const date = new Date(next).toISOString().slice(0, 10);
    if (next < now) return { state: 'Expired', severity: 'critical', expiryText: `${date} (${Math.abs(days)}d ago)` };
    if (next < now + nearDays * 86400000) return { state: 'ExpiringSoon', severity: 'high', expiryText: `${date} (in ${Math.max(0, days)}d)` };
    return { state: 'Valid', severity: 'ok', expiryText: `${date} (in ${Math.max(0, days)}d)` };
  }

  function classifyPermissionRisk(permissionName) {
    const p = String(permissionName || '').toLowerCase();
    if (!p) return 'Info';
    if (['fullcontrol', 'readwrite.all', 'directory.readwrite', 'rolemanagement.readwrite', 'approleassignment.readwrite', 'policy.readwrite'].some((k) => p.includes(k))) return 'High';
    if (['read.all', 'manage.all', 'accessasuser.all', 'sites.selected'].some((k) => p.includes(k))) return 'Medium';
    return 'Low';
  }

  function permissionRiskFromCredentialState(state) {
    if (state === 'Expired') return 'High';
    if (state === 'ExpiringSoon') return 'Medium';
    return 'Info';
  }

  function createResourceResolver(graph) {
    const byId = new Map();
    const byAppId = new Map();
    const SELECT = '$select=id,appId,displayName,appRoles,publishedPermissionScopes';
    const remember = (r) => {
      if (r && r.id) byId.set(r.id, r);
      if (r && r.appId) byAppId.set(r.appId, r);
      return r || null;
    };
    async function getById(id) {
      if (!Http.isGuid(id)) return null;
      if (byId.has(id)) return byId.get(id);
      try { return remember(await graph.json(`/servicePrincipals/${enc(id)}?${SELECT}`)); }
      catch { byId.set(id, null); return null; }
    }
    async function getByAppId(appId) {
      if (!Http.isGuid(appId)) return null;
      if (byAppId.has(appId)) return byAppId.get(appId);
      try {
        const data = await graph.json(`/servicePrincipals?$filter=${enc(`appId eq ${Http.odataString(appId)}`)}&${SELECT}`);
        return remember((data && data.value && data.value[0]) || null);
      } catch { byAppId.set(appId, null); return null; }
    }
    function resolvePermissionName(resource, permissionId, accessType) {
      if (!resource || !permissionId) return permissionId;
      const id = String(permissionId).toLowerCase();
      if (accessType === 'Role') {
        const m = (resource.appRoles || []).find((r) => String(r.id || '').toLowerCase() === id);
        return (m && (m.value || m.displayName)) || permissionId;
      }
      const m = (resource.publishedPermissionScopes || []).find((s) => String(s.id || '').toLowerCase() === id);
      return (m && (m.value || m.adminConsentDisplayName)) || permissionId;
    }
    return { getById, getByAppId, resolvePermissionName };
  }

  // Security roles in a Dataverse-backed environment, read from its Web API.
  async function runDataverseRoles(instanceUrl, envName, ctx) {
    ctx.onProgress({ label: 'Reading Dataverse security roles' });
    let done = 0;
    const path = '/systemusers?$select=fullname,internalemailaddress,domainname,azureactivedirectoryobjectid,isdisabled' +
      '&$expand=systemuserroles_association($select=name,roleid)&$filter=isdisabled eq false';
    await Api.dataversePageAll(instanceUrl, path, (users) => {
      for (const u of users) {
        const aadId = u.azureactivedirectoryobjectid;
        const roles = (u.systemuserroles_association || []).map((r) => r.name).filter(Boolean);
        if (!roles.length) continue;
        for (const role of roles) {
          ctx.onRow(row({
            principal: u.fullname || u.internalemailaddress || u.domainname || aadId,
            principalType: aadId ? 'user' : 'system',
            role,
            scope: `${envName} (Dataverse)`,
            source: 'Dataverse security role',
          }));
        }
        done++;
        ctx.onProgress({ label: 'Reading Dataverse users', current: done });
      }
    });
    if (!done) Log.warn('No Dataverse users with roles returned. Reading users needs a Dataverse security role with read access to the User table.');
  }

  // ---------- Administrators ----------
  const Entra = {
    id: 'entra',
    name: 'Administrators',
    icon: 'i-shield',
    atSignIn: true,
    scopes: ['Directory.Read.All'],
    description: 'Tenant-wide administrator roles and who has them.',
    notice: 'Active directory role assignments only. PIM eligible assignments, administrative unit scopes, conditional access and per-app admin consent are not covered. For the full picture use m365permissions.com.',
    form: [],
    async run(_values, ctx) {
      const graph = G('read');
      ctx.onProgress({ label: 'Listing directory roles' });
      const roles = await graph.pageAll('/directoryRoles');
      ctx.setEntity({ name: 'Tenant directory roles', sub: `${roles.length} active roles` });
      let done = 0;
      await Concurrency.pmap(roles, async (role) => {
        const members = await graph.pageAll(`/directoryRoles/${enc(role.id)}/members`);
        for (const m of members) {
          ctx.onRow(row({
            principal: m.displayName || m.userPrincipalName || m.appDisplayName || m.id,
            principalType: odataType(m),
            role: role.displayName,
            scope: 'Tenant',
            source: 'Directory role',
          }));
        }
        done++;
        ctx.onProgress({ label: `Read ${role.displayName}`, current: done, total: roles.length });
      }, { concurrency: 6, cancelToken: ctx.cancelToken });
    },
  };

  // ---------- Entra app permissions ----------
  const EntraAppAccess = {
    id: 'entra-app-access',
    name: 'Entra App Access',
    icon: 'i-key',
    scopes: ['Application.Read.All'],
    description: 'App registrations and enterprise apps with their API permissions, consent type and credential expiry.',
    notice: 'Configured permissions (app registrations) and granted permissions (service principals). No app role assignment conditions, PIM for workload identities or sign-in data. For workload identity governance use m365permissions.com.',
    columns: [
      { key: 'principal', label: 'Identity' },
      { key: 'principalType', label: 'Type' },
      { key: 'permissionMode', label: 'Mode' },
      { key: 'api', label: 'API' },
      { key: 'role', label: 'Permission' },
      { key: 'consent', label: 'Consent / delegated to' },
      { key: 'credentialState', label: 'Credentials' },
      { key: 'credentialExpiry', label: 'Credential expiry' },
      { key: 'riskLevel', label: 'Risk' },
      { key: 'source', label: 'Source' },
    ],
    quickFilters: [
      { id: 'expired', label: 'Expired credentials', predicate: (r) => r.credentialState === 'Expired' },
      { id: 'expiring', label: 'Expiring soon', predicate: (r) => r.credentialState === 'ExpiringSoon' },
      { id: 'app', label: 'Application permissions', predicate: (r) => r.permissionMode === 'Application' },
      { id: 'delegated', label: 'Delegated permissions', predicate: (r) => r.permissionMode === 'Delegated' },
      { id: 'highrisk', label: 'High risk', predicate: (r) => r.riskLevel === 'High' },
    ],
    form: [],
    async run(_values, ctx) {
      const graph = G('apps');
      ctx.onProgress({ label: 'Loading app registrations and service principals' });
      const [applications, servicePrincipals] = await Promise.all([
        graph.pageAll('/applications?$select=id,appId,displayName,passwordCredentials,keyCredentials,requiredResourceAccess'),
        graph.pageAll('/servicePrincipals?$select=id,appId,displayName,servicePrincipalType,passwordCredentials,keyCredentials'),
      ]);
      ctx.setEntity({ name: 'Entra applications and service principals', sub: `${applications.length} applications, ${servicePrincipals.length} service principals` });

      const resolver = createResourceResolver(graph);
      const total = applications.length + servicePrincipals.length;
      let processed = 0;

      for (const app of applications) {
        ctx.cancelToken.throwIfCancelled();
        const cred = summarizeCredentials(app.passwordCredentials, app.keyCredentials);
        const base = {
          principal: app.displayName || app.appId || app.id,
          principalType: 'application',
          credentialState: cred.state,
          credentialExpiry: cred.expiryText,
          credentialSeverity: cred.severity,
          source: 'Application registration',
        };
        const req = app.requiredResourceAccess || [];
        if (!req.length) {
          ctx.onRow(row({ ...base, permissionMode: '-', api: '-', role: '(no configured API permissions)', consent: '-', riskLevel: permissionRiskFromCredentialState(cred.state) }));
        }
        for (const ra of req) {
          const resource = await resolver.getByAppId(ra.resourceAppId);
          const apiName = (resource && resource.displayName) || ra.resourceAppId || '(unknown API)';
          for (const access of ra.resourceAccess || []) {
            const permName = resolver.resolvePermissionName(resource, access.id, access.type);
            ctx.onRow(row({
              ...base,
              permissionMode: access.type === 'Scope' ? 'Delegated' : 'Application',
              api: apiName,
              role: permName,
              consent: 'Configured in app registration',
              riskLevel: classifyPermissionRisk(permName),
            }));
          }
        }
        processed++;
        ctx.onProgress({ label: 'Reading app registrations', current: processed, total });
      }

      await Concurrency.pmap(servicePrincipals, async (sp) => {
        const cred = summarizeCredentials(sp.passwordCredentials, sp.keyCredentials);
        const base = {
          principal: sp.displayName || sp.appId || sp.id,
          principalType: 'servicePrincipal',
          credentialState: cred.state,
          credentialExpiry: cred.expiryText,
          credentialSeverity: cred.severity,
        };
        let appAssignments = [];
        let delegatedGrants = [];
        try { appAssignments = await graph.pageAll(`/servicePrincipals/${enc(sp.id)}/appRoleAssignments?$select=appRoleId,resourceId`); }
        catch (e) { Log.warn(`Could not read app role assignments for ${sp.displayName}: ${e.message}`); }
        try { delegatedGrants = await graph.pageAll(`/oauth2PermissionGrants?$filter=${enc(`clientId eq ${Http.odataString(sp.id)}`)}&$select=scope,consentType,principalId,resourceId`); }
        catch (e) { Log.warn(`Could not read delegated grants for ${sp.displayName}: ${e.message}`); }

        const principalMap = await resolveDirectoryObjects(delegatedGrants.map((g) => g.principalId).filter(Boolean));

        if (!appAssignments.length && !delegatedGrants.length) {
          ctx.onRow(row({ ...base, permissionMode: '-', api: '-', role: '(no granted API permissions found)', consent: '-', riskLevel: permissionRiskFromCredentialState(cred.state), source: 'Enterprise application' }));
        }
        for (const a of appAssignments) {
          const resource = await resolver.getById(a.resourceId);
          const perm = resolver.resolvePermissionName(resource, a.appRoleId, 'Role');
          ctx.onRow(row({
            ...base,
            permissionMode: 'Application',
            api: (resource && resource.displayName) || a.resourceId || '(unknown API)',
            role: perm,
            consent: 'Assigned to service principal',
            riskLevel: classifyPermissionRisk(perm),
            source: 'Enterprise application grant',
          }));
        }
        for (const g of delegatedGrants) {
          const resource = await resolver.getById(g.resourceId);
          const scopes = String(g.scope || '').split(/\s+/).filter(Boolean);
          const delegatedTo = g.principalId
            ? (principalMap.get(g.principalId) && principalMap.get(g.principalId).displayName) || g.principalId
            : 'All users (tenant-wide consent)';
          for (const scope of scopes.length ? scopes : ['(empty scope)']) {
            ctx.onRow(row({
              ...base,
              permissionMode: 'Delegated',
              api: (resource && resource.displayName) || g.resourceId || '(unknown API)',
              role: scope,
              consent: delegatedTo,
              riskLevel: classifyPermissionRisk(scope),
              source: g.consentType === 'Principal' ? 'Delegated grant (single principal)' : 'Delegated grant (admin consent)',
            }));
          }
        }
        processed++;
        ctx.onProgress({ label: 'Reading enterprise apps', current: processed, total });
      }, { concurrency: 6, cancelToken: ctx.cancelToken });
    },
  };

  // ---------- SharePoint site ----------
  // The M365 group behind a site is only trusted when that group's own root site is
  // this site. Name matches (mailNickname) are candidates, never the answer.
  async function findSiteGroup(graph, site) {
    const candidates = new Set();
    try {
      const drive = await graph.json(`/sites/${site.id}/drive?$select=id,owner`);
      const gid = drive && drive.owner && drive.owner.group && drive.owner.group.id;
      if (Http.isGuid(gid)) candidates.add(gid);
    } catch (e) { Log.dbg(`No default library owner for ${site.webUrl}: ${e.message}`); }
    if (site.name) {
      try {
        const data = await graph.json(`/groups?$filter=${enc(`mailNickname eq ${Http.odataString(site.name)}`)}&$select=id`);
        for (const g of (data && data.value) || []) if (Http.isGuid(g.id)) candidates.add(g.id);
      } catch (e) { Log.dbg(`Group lookup by alias failed: ${e.message}`); }
    }
    const checked = [];
    for (const groupId of candidates) {
      try {
        const root = await graph.json(`/groups/${enc(groupId)}/sites/root?$select=id`);
        checked.push({ groupId, rootSiteId: root && root.id });
      } catch (e) { Log.dbg(`Group ${groupId} has no readable root site: ${e.message}`); }
    }
    return Guards.matchGroupSite(site.id, checked);
  }

  const SharePoint = {
    id: 'sharepoint',
    name: 'SharePoint site',
    icon: 'i-folder',
    scopes: ['Sites.Read.All'],
    description: 'Top-level permissions on a SharePoint site: owners and members of the connected M365 group and app permissions granted through Graph.',
    notice: 'Read through Microsoft Graph only. Classic SharePoint groups with custom permission levels, item-level sharing and unique permissions on subsites and lists are not visible here. For that use SPOTrim or m365permissions.com.',
    form: [
      { kind: 'text', name: 'siteUrl', label: 'Site URL', placeholder: 'https://contoso.sharepoint.com/sites/marketing' },
    ],
    async run({ siteUrl }, ctx) {
      if (!siteUrl) throw new Error('Enter a site URL');
      const u = Guards.sharePointUrl(siteUrl);
      if (!u) throw new Error('Enter an https://<tenant>.sharepoint.com site URL');
      const graph = G('sites');
      ctx.onProgress({ label: 'Resolving site' });
      const site = await graph.json(Guards.graphSitePath(u));
      if (!site || !Guards.isGraphSiteId(site.id)) throw new Error('Site not found');
      ctx.setEntity({ name: site.displayName || site.name, sub: site.webUrl });

      ctx.onProgress({ label: 'Reading site app permissions' });
      try {
        const perms = await graph.pageAll(`/sites/${site.id}/permissions`);
        for (const p of perms) {
          const apps = (p.grantedToIdentitiesV2 || p.grantedToIdentities || []).map((g) => g.application).filter(Boolean);
          ctx.onRow(row({
            principal: apps.map((a) => a.displayName).join(', ') || '(unknown app)',
            principalType: 'application',
            role: (p.roles || []).join(', '),
            scope: 'Site',
            source: 'Graph site permission',
          }));
        }
      } catch (e) {
        Log.warn(`Could not list site permissions: ${e.message}`);
      }

      ctx.onProgress({ label: 'Looking for the connected M365 group' });
      const groupId = await findSiteGroup(graph, site);
      if (!groupId) { Log.info(`No M365 group connected to ${site.webUrl}`); return; }
      ctx.onProgress({ label: 'Reading M365 group owners and members' });
      const sel = '$select=id,displayName,userPrincipalName,appDisplayName';
      const [owners, members] = await Promise.all([
        graph.pageAll(`/groups/${enc(groupId)}/owners?${sel}`),
        graph.pageAll(`/groups/${enc(groupId)}/members?${sel}`),
      ]);
      const add = (list, role) => list.forEach((m) => ctx.onRow(row({
        principal: m.displayName || m.userPrincipalName || m.appDisplayName || m.id,
        principalType: odataType(m) || 'user',
        role,
        scope: 'Site (via M365 group)',
        source: 'M365 group',
      })));
      add(owners, 'Owner');
      add(members, 'Member');
    },
  };

  // ---------- OneDrive ----------
  const OneDrive = {
    id: 'onedrive',
    name: 'OneDrive',
    icon: 'i-file',
    scopes: ['Files.Read.All'],
    description: 'Sharing on the root of a user\'s OneDrive.',
    notice: 'Only sharing on the OneDrive root. Per-folder and per-file sharing is not listed. For a deeper view use SPOTrim or m365permissions.com.',
    form: [
      { kind: 'text', name: 'upn', label: 'User principal name', placeholder: 'jane@contoso.com' },
    ],
    async run({ upn }, ctx) {
      const id = String(upn || '').trim();
      if (!id) throw new Error('Enter a UPN');
      const graph = G('files');
      ctx.onProgress({ label: 'Resolving drive' });
      const drive = await graph.json(`/users/${enc(id)}/drive`);
      if (!drive || !drive.id) throw new Error('No drive for this user');
      ctx.setEntity({ name: drive.owner && drive.owner.user ? drive.owner.user.displayName : id, sub: drive.webUrl || '' });
      ctx.onProgress({ label: 'Reading sharing on the drive root' });
      const perms = await graph.pageAll(`/drives/${enc(drive.id)}/root/permissions`);
      for (const p of perms) {
        const ids = p.grantedToIdentitiesV2 || (p.grantedToV2 ? [p.grantedToV2] : []);
        for (const g of ids.length ? ids : [{}]) {
          const subj = g.user || g.group || g.application || g.siteUser || {};
          ctx.onRow(row({
            principal: subj.displayName || subj.email || (p.link && p.link.scope ? `(link: ${p.link.scope})` : '(anonymous link)'),
            principalType: g.user ? 'user' : g.group ? 'group' : g.application ? 'application' : 'unknown',
            role: (p.roles || []).join(', '),
            scope: 'Drive root',
            source: p.link ? 'Sharing link' : 'Direct grant',
          }));
        }
      }
    },
  };

  // ---------- Mailbox (calendar) ----------
  const Exchange = {
    id: 'exchange',
    name: 'Mailbox',
    icon: 'i-calendar',
    scopes: ['Calendars.Read.Shared'],
    description: 'Calendar sharing permissions on a mailbox.',
    notice: 'Calendar sharing only, the one mailbox permission Microsoft Graph exposes. FullAccess, SendAs, SendOnBehalf and folder permissions live in Exchange Online and need another API. For full mailbox audits use m365permissions.com.',
    form: [
      { kind: 'text', name: 'upn', label: 'Mailbox UPN', placeholder: 'jane@contoso.com' },
    ],
    async run({ upn }, ctx) {
      const id = String(upn || '').trim();
      if (!id) throw new Error('Enter a UPN');
      ctx.setEntity({ name: id, sub: 'Default calendar' });
      ctx.onProgress({ label: 'Reading calendar permissions' });
      const perms = await G('calendars').pageAll(`/users/${enc(id)}/calendar/calendarPermissions`);
      for (const p of perms) {
        const e = p.emailAddress || {};
        ctx.onRow(row({
          principal: e.name || e.address || (p.isInsideOrganization === false ? '(external)' : '(default)'),
          principalType: p.isInsideOrganization === false ? 'external' : 'user',
          role: p.role || (Array.isArray(p.allowedRoles) ? p.allowedRoles.join(', ') : ''),
          scope: 'Calendar',
          source: 'Graph calendar permission',
        }));
      }
    },
  };

  // ---------- Azure subscription ----------
  const Azure = {
    id: 'azure',
    name: 'Azure',
    icon: 'i-server',
    scopes: ['https://management.azure.com/user_impersonation'],
    description: 'Role assignments at subscription scope.',
    notice: 'Assignments at or inherited by the subscription only. Resource group and resource level assignments, classic administrators and PIM eligibility are not listed. For the full picture use m365permissions.com.',
    form: [
      { kind: 'select', name: 'subscriptionId', label: 'Subscription', loadOptions: async () => {
        const subs = await Api.armPageAll('/subscriptions?api-version=2022-12-01');
        return subs.map((s) => ({ value: s.subscriptionId, label: `${s.displayName} (${s.subscriptionId})` }));
      } },
    ],
    async run({ subscriptionId }, ctx) {
      if (!Http.isGuid(subscriptionId)) throw new Error('Pick a subscription');
      ctx.setEntity({ name: 'Subscription', sub: subscriptionId });
      ctx.onProgress({ label: 'Loading role assignments and definitions' });
      const base = `/subscriptions/${enc(subscriptionId)}/providers/Microsoft.Authorization`;
      const [assignments, defs] = await Promise.all([
        Api.armPageAll(`${base}/roleAssignments?api-version=2022-04-01`),
        Api.armPageAll(`${base}/roleDefinitions?api-version=2022-04-01`),
      ]);
      const defMap = new Map(defs.map((d) => [String(d.id).toLowerCase(), d.properties && d.properties.roleName]));
      ctx.onProgress({ label: 'Resolving principals in Entra' });
      const directory = await resolveDirectoryObjects(assignments.map((a) => a.properties && a.properties.principalId));
      for (const a of assignments) {
        const p = a.properties || {};
        const dir = directory.get(p.principalId);
        ctx.onRow(row({
          principal: dir ? dir.displayName : p.principalId,
          principalType: dir ? dir.type : (p.principalType || 'unknown'),
          role: defMap.get(String(p.roleDefinitionId).toLowerCase()) || p.roleDefinitionId,
          scope: p.scope,
          source: 'Azure RBAC',
        }));
      }
    },
  };

  // ---------- Azure DevOps ----------
  const DevOps = {
    id: 'devops',
    name: 'Azure DevOps',
    icon: 'i-pipeline',
    scopes: ['499b84ac-1321-427f-aa17-267ca6975798/user_impersonation'],
    description: 'Organisation-level security groups and their members.',
    notice: 'Organisation-level groups only. Project-level memberships and security namespaces are not listed. For a full audit use m365permissions.com.',
    form: [
      { kind: 'select', name: 'org', label: 'Organisation', loadOptions: async () => {
        const me = await Api.DevOps.json('https://app.vssps.visualstudio.com/_apis/profile/profiles/me?api-version=7.1');
        const accounts = await Api.DevOps.json(`https://app.vssps.visualstudio.com/_apis/accounts?memberId=${enc(me.id)}&api-version=7.1`);
        return ((accounts && accounts.value) || []).map((a) => ({ value: a.accountName, label: a.accountName }));
      } },
    ],
    async run({ org }, ctx) {
      if (!org) throw new Error('Pick an organisation');
      const vssps = `https://vssps.dev.azure.com/${enc(org)}/_apis/graph`;
      ctx.setEntity({ name: 'Azure DevOps organisation', sub: org });
      ctx.onProgress({ label: 'Listing security groups' });
      const groups = [];
      let cont = '';
      do {
        // The continuation token comes in a response header (body as fallback).
        const res = await Api.DevOps.request(`${vssps}/groups?api-version=7.1-preview.1${cont ? `&continuationToken=${enc(cont)}` : ''}`);
        const data = await res.json().catch(() => null);
        if (data && Array.isArray(data.value)) groups.push(...data.value);
        cont = res.headers.get('X-MS-ContinuationToken') || (data && data.continuationToken) || '';
        if (Array.isArray(cont)) cont = cont[0];
        ctx.cancelToken.throwIfCancelled();
      } while (cont);

      const subjectCache = new Map();
      function resolveSubject(descriptor) {
        if (!subjectCache.has(descriptor)) {
          subjectCache.set(descriptor, Api.DevOps.json(`${vssps}/subjects/${enc(descriptor)}?api-version=7.1-preview.1`).catch(() => null));
        }
        return subjectCache.get(descriptor);
      }

      const rawItems = [];
      let done = 0;
      await Concurrency.pmap(groups, async (g) => {
        try {
          const data = await Api.DevOps.json(`${vssps}/Memberships/${enc(g.descriptor)}?direction=down&api-version=7.1-preview.1`);
          await Concurrency.pmap((data && data.value) || [], async (m) => {
            const sub = await resolveSubject(m.memberDescriptor);
            if (sub && sub.origin !== 'vsts') rawItems.push({ sub, role: g.displayName || g.principalName });
          }, { concurrency: 4, cancelToken: ctx.cancelToken });
        } catch (e) {
          Log.warn(`Could not list members of ${g.displayName}: ${e.message}`);
        }
        done++;
        ctx.onProgress({ label: `Read group ${g.displayName}`, current: done, total: groups.length });
      }, { concurrency: 3, cancelToken: ctx.cancelToken });

      ctx.onProgress({ label: 'Resolving Entra identities' });
      const directory = await resolveDirectoryObjects(rawItems.map((r) => r.sub.origin === 'aad' && r.sub.originId).filter(Boolean));
      for (const { sub, role } of rawItems) {
        const dir = sub.origin === 'aad' ? directory.get(sub.originId) : null;
        ctx.onRow(row({
          principal: (dir && dir.displayName) || sub.displayName || sub.principalName || sub.mailAddress || '(unresolved)',
          principalType: dir ? dir.type : (sub.subjectKind || 'unknown'),
          role,
          scope: org,
          source: sub.origin === 'aad' ? 'Entra (via DevOps)' : 'Azure DevOps',
        }));
      }
    },
  };

  // ---------- Power BI workspace ----------
  const PowerBI = {
    id: 'powerbi',
    name: 'Power BI',
    icon: 'i-bars',
    scopes: ['https://analysis.windows.net/powerbi/api/Workspace.Read.All'],
    description: 'Role assignments on a Power BI workspace.',
    notice: 'Only workspaces you have access to are listed, not the admin view of all workspaces. Dataset and report level RLS is not listed. For tenant-wide audits use m365permissions.com.',
    form: [
      { kind: 'select', name: 'workspaceId', label: 'Workspace', loadOptions: async () => {
        const data = await Api.PowerBI.json('/groups');
        return ((data && data.value) || []).map((g) => ({ value: g.id, label: g.name }));
      } },
    ],
    async run({ workspaceId }, ctx) {
      if (!Http.isGuid(workspaceId)) throw new Error('Pick a workspace');
      ctx.setEntity({ name: 'Power BI workspace', sub: workspaceId });
      ctx.onProgress({ label: 'Reading workspace users' });
      const data = await Api.PowerBI.json(`/groups/${enc(workspaceId)}/users`);
      for (const u of (data && data.value) || []) {
        ctx.onRow(row({
          principal: u.displayName || u.identifier || u.emailAddress,
          principalType: u.principalType || 'user',
          role: u.groupUserAccessRight || u.userType || '',
          scope: 'Workspace',
          source: 'Power BI workspace',
        }));
      }
    },
  };

  // ---------- Power Platform ----------
  // Environments from the last loadOptions() call, so run() has the Dataverse metadata.
  const envCache = new Map();
  const BAP_ENVS = '/providers/Microsoft.BusinessAppPlatform';

  const PowerPlatform = {
    id: 'powerplatform',
    name: 'Power Platform',
    icon: 'i-bolt',
    scopes: ['https://service.powerapps.com/User', 'Dataverse user_impersonation (per environment)'],
    description: 'Role assignments on a Power Platform environment.',
    notice: 'Lists all environments when you are a Power Platform admin, otherwise only yours. Dataverse environments are read from the Dataverse Web API. App, flow and connector permissions are not listed. For full Power Platform audits use m365permissions.com.',
    form: [
      { kind: 'select', name: 'envName', label: 'Environment', loadOptions: async () => {
        let envs;
        try {
          envs = ((await Api.PowerPlatform.json(`${BAP_ENVS}/scopes/admin/environments?api-version=2016-11-01`)) || {}).value || [];
        } catch (e) {
          Log.warn(`Admin environment list failed, falling back to your own environments: ${e.message}`);
          envs = ((await Api.PowerPlatform.json(`${BAP_ENVS}/environments?api-version=2016-11-01`)) || {}).value || [];
        }
        envCache.clear();
        for (const env of envs) envCache.set(env.name, env);
        return envs
          .filter((env) => {
            const rt = env.properties && env.properties.states && env.properties.states.runtime;
            return !rt || String(rt.id || '').toLowerCase() !== 'disabled';
          })
          .map((env) => ({ value: env.name, label: (env.properties && env.properties.displayName) || env.name }));
      } },
    ],
    async run({ envName }, ctx) {
      if (!envName) throw new Error('Pick an environment');
      ctx.setEntity({ name: 'Power Platform environment', sub: envName });
      ctx.onProgress({ label: 'Reading role assignments' });
      let list;
      try {
        const data = await Api.PowerPlatform.json(`${BAP_ENVS}/scopes/admin/environments/${enc(envName)}/roleAssignments?api-version=2016-11-01`);
        list = (data && data.value) || [];
      } catch (e) {
        // Dataverse environments answer 403 LinkedEnvironmentForbiddenOperation here.
        if (!(e.status === 403 && /LinkedEnvironment/i.test(e.message))) throw e;
        const env = envCache.get(envName);
        const lem = env && env.properties && env.properties.linkedEnvironmentMetadata;
        const instanceUrl = lem && (lem.instanceApiUrl || lem.instanceUrl);
        if (!instanceUrl) throw new Error('Dataverse environment, but the environment metadata has no instance URL.');
        await runDataverseRoles(instanceUrl, envName, ctx);
        return;
      }
      ctx.onProgress({ label: 'Resolving principals in Entra' });
      const directory = await resolveDirectoryObjects(list.map((r) => r.properties && r.properties.principal && r.properties.principal.id));
      for (const r of list) {
        const p = r.properties || {};
        const principalId = p.principal && p.principal.id;
        const dir = directory.get(principalId);
        ctx.onRow(row({
          principal: dir ? dir.displayName : (p.principal && p.principal.email) || principalId,
          principalType: dir ? dir.type : (p.principal && p.principal.type) || 'unknown',
          role: (p.roleDefinition && (p.roleDefinition.name || p.roleDefinition.id)) || '',
          scope: envName,
          source: 'Power Platform role',
        }));
      }
    },
  };

  window.Providers = {
    list: [Entra, EntraAppAccess, SharePoint, OneDrive, Exchange, Azure, DevOps, PowerBI, PowerPlatform],
    byId(id) { return this.list.find((p) => p.id === id); },
    summarizeCredentials, classifyPermissionRisk,
  };
})();
