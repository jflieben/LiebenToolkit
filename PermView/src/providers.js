// Workload providers for PermView. Each provider exposes a uniform shape so
// app.js can render any of them with the same form / table UI.
//
// Provider shape:
//   id          - stable string id used for the tab
//   name        - human label
//   icon        - 2-3 letters shown in the tab badge
//   description - short blurb shown above the form
//   notice      - extra warning / "quick version" caveat
//   form        - array of input descriptors:
//                   { kind: 'text'|'select', name, label, placeholder?, loadOptions? }
//   run(values) - async function returning:
//                   { entity: { name, sub? }, rows: [{ principal, principalType,
//                     role, scope, source, link? }] }
//
// Rows are deliberately small; PermView is the "quick" view, not the deep
// audit. We always include a top-of-result note pointing power users to
// m365permissions.com.

(() => {

  // ---------- helpers ----------

  function row({ principal, principalType, role, scope, source, link, ...rest }) {
    return {
      principal: principal || '(unknown)',
      principalType: principalType || '',
      role: role || '',
      scope: scope || '',
      source: source || '',
      link: link || null,
      ...rest,
    };
  }

  // Resolve a list of object ids to display names via Graph getByIds.
  // Returns a Map id -> { displayName, type, upn }.
  async function resolveDirectoryObjects(ids) {
    const out = new Map();
    const unique = Array.from(new Set(ids.filter(Boolean)));
    if (!unique.length) return out;
    // getByIds takes max 1000 ids per call.
    for (let i = 0; i < unique.length; i += 1000) {
      const batch = unique.slice(i, i + 1000);
      try {
        const data = await Api.Graph.call('/directoryObjects/getByIds', {
          method: 'POST',
          body: { ids: batch, types: ['user', 'group', 'servicePrincipal', 'device'] },
        });
        for (const obj of (data.value || [])) {
          const t = (obj['@odata.type'] || '').replace('#microsoft.graph.', '');
          out.set(obj.id, {
            displayName: obj.displayName || obj.userPrincipalName || obj.appDisplayName || obj.id,
            type: t,
            upn: obj.userPrincipalName || obj.appId || '',
          });
        }
      } catch (e) {
        Log.warn(`Failed to resolve directory objects: ${e.message}`);
      }
    }
    return out;
  }

  // Encode site URL into Graph addressable path: /sites/{host}:/sites/{name}
  function siteAddressFromUrl(url) {
    const u = new URL(url);
    const host = u.hostname;
    const path = u.pathname.replace(/\/$/, '');
    if (!path || path === '/') return `/sites/${host}`;
    return `/sites/${host}:${path}:`;
  }

  function summarizeCredentials(passwordCredentials, keyCredentials) {
    const now = Date.now();
    const nearDays = 30;
    const creds = [];
    for (const c of (passwordCredentials || [])) {
      if (c && c.endDateTime) creds.push({ kind: 'secret', endDateTime: c.endDateTime });
    }
    for (const c of (keyCredentials || [])) {
      if (c && c.endDateTime) creds.push({ kind: 'certificate', endDateTime: c.endDateTime });
    }
    if (!creds.length) {
      return {
        state: 'None',
        severity: 'none',
        expiryText: 'n/a',
      };
    }

    const expiries = creds
      .map(c => ({ ...c, ts: Date.parse(c.endDateTime) }))
      .filter(c => Number.isFinite(c.ts))
      .sort((a, b) => a.ts - b.ts);

    if (!expiries.length) {
      return {
        state: 'Unknown',
        severity: 'medium',
        expiryText: 'unknown',
      };
    }

    const expired = expiries.filter(c => c.ts < now);
    const near = expiries.filter(c => c.ts >= now && c.ts < (now + nearDays * 86400000));
    const next = expiries[0];
    const days = Math.ceil((next.ts - now) / 86400000);
    const date = new Date(next.ts).toISOString().slice(0, 10);

    if (expired.length) {
      return {
        state: 'Expired',
        severity: 'critical',
        expiryText: `${date} (${Math.abs(days)}d ago)`,
      };
    }
    if (near.length) {
      return {
        state: 'ExpiringSoon',
        severity: 'high',
        expiryText: `${date} (in ${Math.max(0, days)}d)`,
      };
    }
    return {
      state: 'Valid',
      severity: 'ok',
      expiryText: `${date} (in ${Math.max(0, days)}d)`,
    };
  }

  function classifyPermissionRisk(permissionName) {
    const p = String(permissionName || '').toLowerCase();
    if (!p) return 'Info';
    if (
      p.includes('fullcontrol') ||
      p.includes('readwrite.all') ||
      p.includes('directory.readwrite') ||
      p.includes('rolemanagement.readwrite') ||
      p.includes('approleassignment.readwrite') ||
      p.includes('policy.readwrite')
    ) return 'High';
    if (
      p.includes('read.all') ||
      p.includes('manage.all') ||
      p.includes('accessasuser.all') ||
      p.includes('sites.selected')
    ) return 'Medium';
    return 'Low';
  }

  function permissionRiskFromCredentialState(state) {
    if (state === 'Expired') return 'High';
    if (state === 'ExpiringSoon') return 'Medium';
    return 'Info';
  }

  function quoteOData(v) {
    return String(v || '').replace(/'/g, "''");
  }

  function createResourceResolver() {
    const byId = new Map();
    const byAppId = new Map();

    function mapResource(resource) {
      if (!resource) return null;
      if (resource.id) byId.set(resource.id, resource);
      if (resource.appId) byAppId.set(resource.appId, resource);
      return resource;
    }

    async function getById(resourceId) {
      if (!resourceId) return null;
      if (byId.has(resourceId)) return byId.get(resourceId);
      try {
        const resource = await Api.Graph.call(`/servicePrincipals/${resourceId}?$select=id,appId,displayName,appRoles,publishedPermissionScopes`);
        return mapResource(resource);
      } catch {
        byId.set(resourceId, null);
        return null;
      }
    }

    async function getByAppId(appId) {
      if (!appId) return null;
      if (byAppId.has(appId)) return byAppId.get(appId);
      try {
        const data = await Api.Graph.call(`/servicePrincipals?$filter=appId eq '${quoteOData(appId)}'&$select=id,appId,displayName,appRoles,publishedPermissionScopes`);
        const resource = (data && data.value && data.value[0]) || null;
        return mapResource(resource);
      } catch {
        byAppId.set(appId, null);
        return null;
      }
    }

    function resolvePermissionName(resource, permissionId, accessType) {
      if (!resource || !permissionId) return permissionId;
      const id = String(permissionId).toLowerCase();
      if (accessType === 'Role') {
        const match = (resource.appRoles || []).find(r => String(r.id || '').toLowerCase() === id);
        return (match && (match.value || match.displayName)) || permissionId;
      }
      const match = (resource.publishedPermissionScopes || []).find(s => String(s.id || '').toLowerCase() === id);
      return (match && (match.value || match.adminConsentDisplayName)) || permissionId;
    }

    return { getById, getByAppId, resolvePermissionName };
  }

  // Read security roles for a Dataverse-linked Power Platform environment.
  // Streams rows through ctx as users + their assigned roles are resolved.
  async function runDataverseRoles(instanceUrl, envName, ctx) {
    ctx && ctx.onProgress({ label: 'Reading Dataverse security roles...' });
    // systemuserroles_association links systemusers to roles.
    // We page via @odata.nextLink. $expand returns the role rows inline.
    const userMap = new Map(); // azureactivedirectoryobjectid -> { name, upn, roles[] }
    let next = `/systemusers?$select=fullname,internalemailaddress,domainname,azureactivedirectoryobjectid,isdisabled&$expand=systemuserroles_association($select=name,roleid)&$filter=isdisabled eq false`;
    let done = 0;
    while (next) {
      const data = await Api.Dataverse.call(instanceUrl, next);
      const users = (data && data.value) || [];
      for (const u of users) {
        const aadId = u.azureactivedirectoryobjectid;
        const roles = (u.systemuserroles_association || []).map(r => r.name).filter(Boolean);
        if (!roles.length) continue;
        for (const role of roles) {
          ctx && ctx.onRow(row({
            principal: u.fullname || u.internalemailaddress || u.domainname || aadId || '(unknown)',
            principalType: aadId ? 'user' : 'system',
            role,
            scope: envName + ' (Dataverse)',
            source: 'Dataverse security role',
          }));
        }
        done++;
        ctx && ctx.onProgress({ label: 'Streaming Dataverse users', current: done });
      }
      next = data && data['@odata.nextLink'];
    }
    if (!done) {
      Log.warn('No Dataverse users with roles returned. You may need a Dataverse security role with read access to the User entity.');
    }
  }

  // ---------- 1) Entra (directory roles) ----------

  const Entra = {
    id: 'entra',
    name: 'Administrators',
    icon: 'EN',
    description: 'Tenant-wide administrator roles and the principals assigned to them.',
    notice: 'Shows active directory role assignments only. PIM eligible assignments, custom RBAC scopes, conditional access and per-app admin consent are not covered here. For a full picture, see <a href="https://www.m365permissions.com" target="_blank" rel="noopener">m365permissions.com</a>.',
    form: [],
    async run(_values, ctx) {
      ctx && ctx.onProgress({ label: 'Listing directory roles...' });
      const roles = await Api.Graph.pageAll('/directoryRoles');
      ctx && ctx.setEntity({ name: 'Tenant directory roles', sub: `${roles.length} roles` });
      let done = 0;
      await Concurrency.pmap(roles, async (role) => {
        const members = await Api.Graph.pageAll(`/directoryRoles/${role.id}/members`);
        for (const m of members) {
          const t = (m['@odata.type'] || '').replace('#microsoft.graph.', '');
          ctx && ctx.onRow(row({
            principal: m.displayName || m.userPrincipalName || m.appDisplayName || m.id,
            principalType: t,
            role: role.displayName,
            scope: 'Tenant',
            source: 'Directory role',
          }));
        }
        done++;
        ctx && ctx.onProgress({ label: `Resolving ${role.displayName}`, current: done, total: roles.length });
      }, { concurrency: 6 });
    },
  };

  // ---------- 1b) Entra app permissions ----------

  const EntraAppAccess = {
    id: 'entra-app-access',
    name: 'Entra App Access',
    icon: 'EA',
    description: 'Application registrations and enterprise apps with API permissions, consent model and credential hygiene signals.',
    notice: 'This view combines app registrations (configured permissions) and service principals (granted permissions). It does not include app role assignment conditions, PIM for workload identities, or sign-in telemetry. Use m365permissions.com for deep workload identity governance.',
    columns: [
      { key: 'principal', label: 'Identity' },
      { key: 'principalType', label: 'Type' },
      { key: 'permissionMode', label: 'Mode' },
      { key: 'api', label: 'API' },
      { key: 'role', label: 'Permission' },
      { key: 'consent', label: 'Consent / Delegated To' },
      { key: 'credentialState', label: 'Credentials' },
      { key: 'credentialExpiry', label: 'Credential Expiry' },
      { key: 'riskLevel', label: 'Risk' },
      { key: 'source', label: 'Source' },
    ],
    quickFilters: [
      { id: 'expired', label: 'Expired Credentials', predicate: r => r.credentialState === 'Expired' },
      { id: 'expiring', label: 'Expiring Soon', predicate: r => r.credentialState === 'ExpiringSoon' },
      { id: 'app', label: 'Application Permissions', predicate: r => r.permissionMode === 'Application' },
      { id: 'delegated', label: 'Delegated Permissions', predicate: r => r.permissionMode === 'Delegated' },
      { id: 'highrisk', label: 'High Risk', predicate: r => r.riskLevel === 'High' },
    ],
    form: [],
    async run(_values, ctx) {
      ctx && ctx.onProgress({ label: 'Loading app registrations and service principals...' });
      const [applications, servicePrincipals] = await Promise.all([
        Api.Graph.pageAll('/applications?$select=id,appId,displayName,passwordCredentials,keyCredentials,requiredResourceAccess'),
        Api.Graph.pageAll('/servicePrincipals?$select=id,appId,displayName,servicePrincipalType,passwordCredentials,keyCredentials'),
      ]);
      ctx && ctx.setEntity({
        name: 'Entra applications and service principals',
        sub: `${applications.length} applications, ${servicePrincipals.length} service principals`,
      });

      const resolver = createResourceResolver();
      let processed = 0;

      // App registrations: configured permissions (Role/Scope) by API.
      ctx && ctx.onProgress({ label: 'Reading application registration permissions...', current: 0, total: applications.length + servicePrincipals.length });
      for (const app of applications) {
        const cred = summarizeCredentials(app.passwordCredentials, app.keyCredentials);
        const req = app.requiredResourceAccess || [];
        if (!req.length) {
          ctx && ctx.onRow(row({
            principal: app.displayName || app.appId || app.id,
            principalType: 'application',
            permissionMode: '-',
            api: '-',
            role: '(no configured API permissions)',
            consent: '-',
            credentialState: cred.state,
            credentialExpiry: cred.expiryText,
            riskLevel: permissionRiskFromCredentialState(cred.state),
            source: 'Application registration',
            credentialSeverity: cred.severity,
          }));
        } else {
          for (const resourceAccess of req) {
            const resource = await resolver.getByAppId(resourceAccess.resourceAppId);
            const apiName = (resource && resource.displayName) || resourceAccess.resourceAppId || '(unknown API)';
            for (const access of (resourceAccess.resourceAccess || [])) {
              const mode = access.type === 'Scope' ? 'Delegated' : 'Application';
              const permName = resolver.resolvePermissionName(resource, access.id, access.type);
              ctx && ctx.onRow(row({
                principal: app.displayName || app.appId || app.id,
                principalType: 'application',
                permissionMode: mode,
                api: apiName,
                role: permName,
                consent: 'Configured in app registration',
                credentialState: cred.state,
                credentialExpiry: cred.expiryText,
                riskLevel: classifyPermissionRisk(permName),
                source: 'Application registration',
                credentialSeverity: cred.severity,
              }));
            }
          }
        }
        processed++;
        ctx && ctx.onProgress({ label: 'Reading app registrations', current: processed, total: applications.length + servicePrincipals.length });
      }

      // Service principals: granted app roles + delegated OAuth grants.
      await Concurrency.pmap(servicePrincipals, async (sp) => {
        const cred = summarizeCredentials(sp.passwordCredentials, sp.keyCredentials);
        let appAssignments = [];
        let delegatedGrants = [];

        try {
          appAssignments = await Api.Graph.pageAll(`/servicePrincipals/${sp.id}/appRoleAssignments?$select=appRoleId,resourceId`);
        } catch (e) {
          Log.warn(`Could not read app role assignments for ${sp.displayName}: ${e.message}`);
        }

        try {
          delegatedGrants = await Api.Graph.pageAll(`/oauth2PermissionGrants?$filter=clientId eq '${quoteOData(sp.id)}'&$select=scope,consentType,principalId,resourceId`);
        } catch (e) {
          Log.warn(`Could not read delegated grants for ${sp.displayName}: ${e.message}`);
        }

        const principalIds = delegatedGrants.map(g => g.principalId).filter(Boolean);
        const principalMap = await resolveDirectoryObjects(principalIds);

        if (!appAssignments.length && !delegatedGrants.length) {
          ctx && ctx.onRow(row({
            principal: sp.displayName || sp.appId || sp.id,
            principalType: 'servicePrincipal',
            permissionMode: '-',
            api: '-',
            role: '(no granted API permissions found)',
            consent: '-',
            credentialState: cred.state,
            credentialExpiry: cred.expiryText,
            riskLevel: permissionRiskFromCredentialState(cred.state),
            source: 'Enterprise application',
            credentialSeverity: cred.severity,
          }));
        }

        for (const a of appAssignments) {
          const resource = await resolver.getById(a.resourceId);
          const apiName = (resource && resource.displayName) || a.resourceId || '(unknown API)';
          const perm = resolver.resolvePermissionName(resource, a.appRoleId, 'Role');
          ctx && ctx.onRow(row({
            principal: sp.displayName || sp.appId || sp.id,
            principalType: 'servicePrincipal',
            permissionMode: 'Application',
            api: apiName,
            role: perm,
            consent: 'Assigned to service principal',
            credentialState: cred.state,
            credentialExpiry: cred.expiryText,
            riskLevel: classifyPermissionRisk(perm),
            source: 'Enterprise application grant',
            credentialSeverity: cred.severity,
          }));
        }

        for (const g of delegatedGrants) {
          const resource = await resolver.getById(g.resourceId);
          const apiName = (resource && resource.displayName) || g.resourceId || '(unknown API)';
          const scopes = String(g.scope || '').split(/\s+/).filter(Boolean);
          const delegatedTo = g.principalId
            ? (principalMap.get(g.principalId) && principalMap.get(g.principalId).displayName) || g.principalId
            : 'All users (tenant-wide consent)';
          for (const scope of (scopes.length ? scopes : ['(empty scope)'])) {
            ctx && ctx.onRow(row({
              principal: sp.displayName || sp.appId || sp.id,
              principalType: 'servicePrincipal',
              permissionMode: 'Delegated',
              api: apiName,
              role: scope,
              consent: delegatedTo,
              credentialState: cred.state,
              credentialExpiry: cred.expiryText,
              riskLevel: classifyPermissionRisk(scope),
              source: g.consentType === 'Principal' ? 'Delegated grant (single principal)' : 'Delegated grant (admin consent)',
              credentialSeverity: cred.severity,
            }));
          }
        }

        processed++;
        ctx && ctx.onProgress({ label: 'Reading enterprise apps', current: processed, total: applications.length + servicePrincipals.length });
      }, { concurrency: 6 });
    },
  };

  // ---------- 2) SharePoint site ----------

  const SharePoint = {
    id: 'sharepoint',
    name: 'SharePoint site',
    icon: 'SP',
    description: 'Top-level permissions on a SharePoint site (owners, members, visitors and Graph-granted app permissions).',
    notice: 'PermView reads SharePoint via Microsoft Graph only. It will not see classic SharePoint groups with custom permission levels, item-level sharing, or unique permissions on subsites and lists. For that, use SPOTrim or m365permissions.com.',
    form: [
      { kind: 'text', name: 'siteUrl', label: 'Site URL', placeholder: 'https://contoso.sharepoint.com/sites/marketing' },
    ],
    async run({ siteUrl }, ctx) {
      if (!siteUrl) throw new Error('Enter a site URL');
      ctx && ctx.onProgress({ label: 'Resolving site...' });
      const addr = siteAddressFromUrl(siteUrl);
      const site = await Api.Graph.call(addr);
      if (!site || !site.id) throw new Error('Site not found');
      ctx && ctx.setEntity({ name: site.displayName || site.name, sub: site.webUrl });

      // Graph site permissions = application permissions granted to this site.
      ctx && ctx.onProgress({ label: 'Reading site app permissions...' });
      try {
        const perms = await Api.Graph.pageAll(`/sites/${site.id}/permissions`);
        for (const p of perms) {
          const apps = (p.grantedToIdentitiesV2 || p.grantedToIdentities || []).map(g => g.application).filter(Boolean);
          const appLabel = apps.map(a => a.displayName).join(', ') || '(unknown app)';
          ctx && ctx.onRow(row({
            principal: appLabel,
            principalType: 'application',
            role: (p.roles || []).join(', '),
            scope: 'Site',
            source: 'Graph site permission',
          }));
        }
      } catch (e) {
        Log.warn('Could not list site permissions: ' + e.message);
      }

      // M365 group backed site? List owners + members of the group.
      ctx && ctx.onProgress({ label: 'Checking for M365 group backing...' });
      try {
        const groups = await Api.Graph.call(`/groups?$filter=mailNickname eq '${(site.name || '').replace(/'/g, "''")}'&$select=id,displayName`);
        const grp = groups && groups.value && groups.value[0];
        if (grp) {
          ctx && ctx.onProgress({ label: 'Reading M365 group owners and members...' });
          const [owners, members] = await Promise.all([
            Api.Graph.pageAll(`/groups/${grp.id}/owners?$select=id,displayName,userPrincipalName`),
            Api.Graph.pageAll(`/groups/${grp.id}/members?$select=id,displayName,userPrincipalName`),
          ]);
          for (const u of owners) ctx && ctx.onRow(row({
            principal: u.displayName || u.userPrincipalName,
            principalType: 'user',
            role: 'Owner',
            scope: 'Site (via M365 group)',
            source: 'M365 group',
          }));
          for (const u of members) ctx && ctx.onRow(row({
            principal: u.displayName || u.userPrincipalName,
            principalType: 'user',
            role: 'Member',
            scope: 'Site (via M365 group)',
            source: 'M365 group',
          }));
        }
      } catch (e) { Log.dbg('No group-backed site detected: ' + e.message); }
    },
  };

  // ---------- 3) OneDrive ----------

  const OneDrive = {
    id: 'onedrive',
    name: 'OneDrive',
    icon: 'OD',
    description: 'Sharing on the root of a user OneDrive.',
    notice: 'Only sharing on the OneDrive root is shown. Per-folder or per-file sharing is not enumerated here. For a deeper view, use SPOTrim or m365permissions.com.',
    form: [
      { kind: 'text', name: 'upn', label: 'User principal name', placeholder: 'jane@contoso.com' },
    ],
    async run({ upn }, ctx) {
      if (!upn) throw new Error('Enter a UPN');
      ctx && ctx.onProgress({ label: 'Resolving drive...' });
      const drive = await Api.Graph.call(`/users/${encodeURIComponent(upn)}/drive`);
      if (!drive || !drive.id) throw new Error('No drive for this user');
      ctx && ctx.setEntity({ name: drive.owner && drive.owner.user ? drive.owner.user.displayName : upn, sub: drive.webUrl || '' });
      ctx && ctx.onProgress({ label: 'Reading drive root sharing...' });
      const perms = await Api.Graph.pageAll(`/drives/${drive.id}/root/permissions`);
      for (const p of perms) {
        const ids = p.grantedToIdentitiesV2 || (p.grantedToV2 ? [p.grantedToV2] : []);
        const list = ids.length ? ids : [{}];
        for (const g of list) {
          const subj = g.user || g.group || g.application || g.siteUser || {};
          const principalType = g.user ? 'user' : g.group ? 'group' : g.application ? 'application' : 'unknown';
          ctx && ctx.onRow(row({
            principal: subj.displayName || subj.email || (p.link && p.link.scope ? `(link: ${p.link.scope})` : '(anonymous link)'),
            principalType,
            role: (p.roles || []).join(', '),
            scope: 'Drive root',
            source: p.link ? 'Sharing link' : 'Direct grant',
          }));
        }
      }
    },
  };

  // ---------- 4) Mailbox (calendar permissions) ----------

  const Exchange = {
    id: 'exchange',
    name: 'Mailbox',
    icon: 'EX',
    description: 'Calendar sharing permissions on a mailbox.',
    notice: 'PermView reads calendar sharing only - it is the only mailbox-level permission Microsoft Graph exposes. FullAccess, SendAs, SendOnBehalf and Inbox folder permissions live in Exchange Online and need a different API. For full mailbox audits, use m365permissions.com.',
    form: [
      { kind: 'text', name: 'upn', label: 'Mailbox UPN', placeholder: 'jane@contoso.com' },
    ],
    async run({ upn }, ctx) {
      if (!upn) throw new Error('Enter a UPN');
      ctx && ctx.setEntity({ name: upn, sub: 'Default calendar' });
      ctx && ctx.onProgress({ label: 'Reading calendar permissions...' });
      const perms = await Api.Graph.pageAll(`/users/${encodeURIComponent(upn)}/calendar/calendarPermissions`);
      ctx && ctx.onProgress({ label: 'Rendering permissions', current: 0, total: perms.length });
      let i = 0;
      for (const p of perms) {
        const e = p.emailAddress || {};
        const role = p.role || (Array.isArray(p.allowedRoles) ? p.allowedRoles.join(', ') : '');
        ctx && ctx.onRow(row({
          principal: e.name || e.address || (p.isInsideOrganization === false ? '(external)' : '(default / unknown)'),
          principalType: p.isInsideOrganization === false ? 'external' : 'user',
          role,
          scope: 'Calendar',
          source: 'Graph calendar permission',
        }));
        i++;
        ctx && ctx.onProgress({ label: 'Rendering permissions', current: i, total: perms.length });
      }
    },
  };

  // ---------- 5) Azure subscription ----------

  const Azure = {
    id: 'azure',
    name: 'Azure',
    icon: 'AZ',
    description: 'Role assignments at the subscription scope.',
    notice: 'Only assignments scoped at, or inherited to, the subscription are shown. Resource-group and resource-level assignments, classic administrators, and Azure AD PIM eligibility are not enumerated here. For a full picture, use m365permissions.com.',
    form: [
      { kind: 'select', name: 'subscriptionId', label: 'Subscription', loadOptions: async () => {
        const subs = await Api.Arm.pageAll('/subscriptions?api-version=2022-12-01');
        return subs.map(s => ({ value: s.subscriptionId, label: `${s.displayName} (${s.subscriptionId})` }));
      } },
    ],
    async run({ subscriptionId }, ctx) {
      if (!subscriptionId) throw new Error('Pick a subscription');
      ctx && ctx.setEntity({ name: 'Subscription', sub: subscriptionId });
      ctx && ctx.onProgress({ label: 'Loading role assignments and definitions...' });
      const [assignments, defs] = await Promise.all([
        Api.Arm.pageAll(`/subscriptions/${subscriptionId}/providers/Microsoft.Authorization/roleAssignments?api-version=2022-04-01`),
        Api.Arm.pageAll(`/subscriptions/${subscriptionId}/providers/Microsoft.Authorization/roleDefinitions?api-version=2022-04-01`),
      ]);
      const defMap = new Map(defs.map(d => [d.id, d.properties && d.properties.roleName]));
      ctx && ctx.onProgress({ label: 'Resolving principals via Entra...' });
      const principalIds = assignments.map(a => a.properties && a.properties.principalId);
      const directory = await resolveDirectoryObjects(principalIds);
      ctx && ctx.onProgress({ label: 'Rendering assignments', current: 0, total: assignments.length });
      let i = 0;
      for (const a of assignments) {
        const p = a.properties || {};
        const dir = directory.get(p.principalId);
        ctx && ctx.onRow(row({
          principal: dir ? dir.displayName : p.principalId,
          principalType: dir ? dir.type : (p.principalType || 'unknown'),
          role: defMap.get(p.roleDefinitionId) || p.roleDefinitionId,
          scope: p.scope,
          source: 'Azure RBAC',
        }));
        i++;
        ctx && ctx.onProgress({ label: 'Rendering assignments', current: i, total: assignments.length });
      }
    },
  };

  // ---------- 6) Azure DevOps ----------

  const DevOps = {
    id: 'devops',
    name: 'Azure DevOps',
    icon: 'AD',
    description: 'Org-level security groups and their members in an Azure DevOps organisation.',
    notice: 'Only organisation-level groups are listed. Project-level group memberships and project security descriptors are not enumerated here. For a full audit see m365permissions.com.',
    form: [
      { kind: 'select', name: 'org', label: 'Organisation', loadOptions: async () => {
        // Get current user id, then list their orgs.
        const me = await Api.DevOps.call('https://app.vssps.visualstudio.com/_apis/profile/profiles/me?api-version=7.1');
        const accounts = await Api.DevOps.call(`https://app.vssps.visualstudio.com/_apis/accounts?memberId=${me.id}&api-version=7.1`);
        return (accounts.value || []).map(a => ({ value: a.accountName, label: a.accountName }));
      } },
    ],
    async run({ org }, ctx) {
      if (!org) throw new Error('Pick an org');
      ctx && ctx.setEntity({ name: 'Azure DevOps org', sub: org });
      ctx && ctx.onProgress({ label: 'Listing security groups...' });
      const groups = [];
      let cont = '';
      while (true) {
        const url = `https://vssps.dev.azure.com/${encodeURIComponent(org)}/_apis/graph/groups?api-version=7.1-preview.1${cont ? `&continuationToken=${encodeURIComponent(cont)}` : ''}`;
        const data = await Api.DevOps.call(url);
        if (data && Array.isArray(data.value)) groups.push(...data.value);
        cont = data && data.continuationToken;
        if (!cont) break;
      }

      const subjectCache = new Map();
      async function resolveSubject(descriptor) {
        if (subjectCache.has(descriptor)) return subjectCache.get(descriptor);
        try {
          const sub = await Api.DevOps.call(`https://vssps.dev.azure.com/${encodeURIComponent(org)}/_apis/graph/subjects/${descriptor}?api-version=7.1-preview.1`);
          subjectCache.set(descriptor, sub);
          return sub;
        } catch (e) {
          subjectCache.set(descriptor, null);
          return null;
        }
      }

      // Phase 1: discover members per group, accumulate raw items.
      const rawItems = [];
      let done = 0;
      await Concurrency.pmap(groups, async (g) => {
        try {
          const memUrl = `https://vssps.dev.azure.com/${encodeURIComponent(org)}/_apis/graph/Memberships/${g.descriptor}?direction=down&api-version=7.1-preview.1`;
          const data = await Api.DevOps.call(memUrl);
          const members = (data && data.value) || [];
          await Concurrency.pmap(members, async (m) => {
            const sub = await resolveSubject(m.memberDescriptor);
            if (!sub) return;
            if (sub.origin === 'vsts') return;
            rawItems.push({ sub, role: g.displayName || g.principalName });
          }, { concurrency: 4 });
        } catch (e) {
          Log.warn(`Could not enumerate ${g.displayName}: ${e.message}`);
        }
        done++;
        ctx && ctx.onProgress({ label: `Walking group ${g.displayName}`, current: done, total: groups.length });
      }, { concurrency: 3 });

      // Phase 2: enrich AAD-backed subjects with Entra display names.
      ctx && ctx.onProgress({ label: 'Resolving Entra identities...' });
      const aadIds = rawItems.map(r => r.sub.origin === 'aad' && r.sub.originId).filter(Boolean);
      const directory = await resolveDirectoryObjects(aadIds);

      ctx && ctx.onProgress({ label: 'Rendering rows', current: 0, total: rawItems.length });
      let i = 0;
      for (const { sub, role } of rawItems) {
        const dir = sub.origin === 'aad' ? directory.get(sub.originId) : null;
        const principal = (dir && dir.displayName) || sub.displayName || sub.principalName || sub.mailAddress || '(unresolved)';
        const principalType = dir ? dir.type : (sub.subjectKind || 'unknown');
        ctx && ctx.onRow(row({
          principal,
          principalType,
          role,
          scope: org,
          source: sub.origin === 'aad' ? 'Entra (via DevOps)' : 'Azure DevOps',
        }));
        i++;
        ctx && ctx.onProgress({ label: 'Rendering rows', current: i, total: rawItems.length });
      }
    },
  };

  // ---------- 7) Power BI workspace ----------

  const PowerBI = {
    id: 'powerbi',
    name: 'Power BI',
    icon: 'PB',
    description: 'Role assignments on a Power BI workspace.',
    notice: 'Only workspaces you have access to are listed (not the tenant-admin "all workspaces" view). Dataset / report-level RLS is not enumerated. For tenant-wide audits use m365permissions.com.',
    form: [
      { kind: 'select', name: 'workspaceId', label: 'Workspace', loadOptions: async () => {
        const data = await Api.PowerBI.call('/groups');
        return (data.value || []).map(g => ({ value: g.id, label: g.name }));
      } },
    ],
    async run({ workspaceId }, ctx) {
      if (!workspaceId) throw new Error('Pick a workspace');
      ctx && ctx.setEntity({ name: 'Power BI workspace', sub: workspaceId });
      ctx && ctx.onProgress({ label: 'Reading workspace users...' });
      const data = await Api.PowerBI.call(`/groups/${workspaceId}/users`);
      const users = data.value || [];
      ctx && ctx.onProgress({ label: 'Rendering rows', current: 0, total: users.length });
      let i = 0;
      for (const u of users) {
        ctx && ctx.onRow(row({
          principal: u.displayName || u.identifier || u.emailAddress,
          principalType: u.principalType || 'user',
          role: u.groupUserAccessRight || u.userType || '',
          scope: 'Workspace',
          source: 'Power BI workspace',
        }));
        i++;
        ctx && ctx.onProgress({ label: 'Rendering rows', current: i, total: users.length });
      }
    },
  };

  // ---------- 8) Power Platform / Power Automate ----------

  // Cache of env objects from the most recent loadOptions() call, so run()
  // can look up linked-Dataverse metadata without refetching.
  const _envCache = new Map();

  const PowerPlatform = {
    id: 'powerplatform',
    name: 'Power Platform',
    icon: 'PP',
    description: 'Role assignments on a Power Platform environment.',
    notice: 'Requires you to be a Power Platform admin (the BAP admin scope is used). For Dataverse-linked environments, security roles are read directly from the Dataverse Web API. App, flow and connector-level permissions are not enumerated here. For full Power Platform audits, use m365permissions.com.',
    form: [
      { kind: 'select', name: 'envName', label: 'Environment', loadOptions: async () => {
        // Admin endpoint enumerates ALL tenant environments (requires PP admin).
        // Falls back to user-scope environments if admin call fails.
        let envs = [];
        try {
          const data = await Api.PowerPlatform.call('/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2016-11-01');
          envs = data.value || [];
        } catch (e) {
          Log.warn('Admin environments call failed: ' + e.message + ' — falling back to user environments');
          try {
            const data = await Api.PowerPlatform.call('/providers/Microsoft.BusinessAppPlatform/environments?api-version=2016-11-01');
            envs = data.value || [];
          } catch (e2) {
            Log.err('User environments call also failed: ' + e2.message);
            throw e2;
          }
        }
        _envCache.clear();
        for (const e of envs) _envCache.set(e.name, e);
        return envs
          .filter(e => {
            // Skip disabled environments.
            const rt = e.properties && e.properties.states && e.properties.states.runtime;
            return !rt || (rt.id || '').toLowerCase() !== 'disabled';
          })
          .map(e => ({ value: e.name, label: (e.properties && e.properties.displayName) || e.name }));
      } },
    ],
    async run({ envName }, ctx) {
      if (!envName) throw new Error('Pick an environment');
      ctx && ctx.setEntity({ name: 'Power Platform environment', sub: envName });
      ctx && ctx.onProgress({ label: 'Reading role assignments...' });

      let list = null;
      try {
        const data = await Api.PowerPlatform.call(`/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${envName}/roleAssignments?api-version=2016-11-01`);
        list = data.value || [];
      } catch (e) {
        // Dataverse-linked environments reject this call:
        //   403 LinkedEnvironmentForbiddenOperation
        // Roles must be read from the Dataverse Web API instead.
        const linkedEnv = e.status === 403 && /LinkedEnvironment/i.test(e.message);
        if (!linkedEnv) throw e;
        const env = _envCache.get(envName);
        const lem = env && env.properties && env.properties.linkedEnvironmentMetadata;
        const instanceUrl = lem && (lem.instanceApiUrl || lem.instanceUrl);
        if (!instanceUrl) {
          throw new Error('Linked Dataverse environment but no instance URL found in environment metadata.');
        }
        await runDataverseRoles(instanceUrl, envName, ctx);
        return;
      }

      ctx && ctx.onProgress({ label: 'Resolving principals via Entra...' });
      const principalIds = list.map(r => r.properties && r.properties.principal && r.properties.principal.id).filter(Boolean);
      const directory = await resolveDirectoryObjects(principalIds);
      ctx && ctx.onProgress({ label: 'Rendering rows', current: 0, total: list.length });
      let i = 0;
      for (const r of list) {
        const p = r.properties || {};
        const principalId = p.principal && p.principal.id;
        const dir = directory.get(principalId);
        const roleDef = (p.roleDefinition && (p.roleDefinition.name || p.roleDefinition.id)) || '';
        ctx && ctx.onRow(row({
          principal: dir ? dir.displayName : (p.principal && p.principal.email) || principalId,
          principalType: dir ? dir.type : (p.principal && p.principal.type) || 'unknown',
          role: roleDef,
          scope: envName,
          source: 'Power Platform role',
        }));
        i++;
        ctx && ctx.onProgress({ label: 'Rendering rows', current: i, total: list.length });
      }
    },
  };

  window.Providers = {
    list: [Entra, EntraAppAccess, SharePoint, OneDrive, Exchange, Azure, DevOps, PowerBI, PowerPlatform],
    byId(id) { return this.list.find(p => p.id === id); },
  };
})();
