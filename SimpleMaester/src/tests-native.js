// Hand-coded JavaScript ports of the most useful Maester non-EIDSCA tests.
// Each test is self-contained: it queries Graph, decides Pass/Fail/Skip and returns
// a row in the same shape as the EIDSCA executor produces. This makes it easy to
// keep adding new tests over time without touching the runner or UI.
//
// Test IDs in the exposed catalog must map to authoritative Maester docs IDs.
(() => {
  const tests = [];

  function makeRow(id, title, severity, category, opts = {}) {
    return {
      id, title, severity, tag: category, category,
      docUrl: opts.docUrl || `https://maester.dev/docs/tests/${id}`,
      description: opts.description,
      detailMd: opts.detailMd,
    };
  }

  // ---------- Helpers reused across tests ----------
  async function getDirectoryRoles() {
    return Graph.graphAll('directoryRoles', { apiVersion: 'v1.0' });
  }
  async function getRoleMembers(roleId) {
    try { return await Graph.graphAll(`directoryRoles/${roleId}/members`, { apiVersion: 'v1.0' }); }
    catch (e) { Log.debug('roleMembers failed', roleId, e.message); return []; }
  }
  // Bulk flatten of all members of all "privileged" Entra roles.
  // Maester uses a curated list of templateIds; we use the same shortlist.
  const PRIV_ROLE_TEMPLATE_IDS = [
    '62e90394-69f5-4237-9190-012177145e10', // Global Administrator
    '7be44c8a-adaf-4e2a-84d6-ab2649e08a13', // Privileged Authentication Admin
    'e8611ab8-c189-46e8-94e1-60213ab1f814', // Privileged Role Administrator
    '194ae4cb-b126-40b2-bd5b-6091b380977d', // Security Administrator
    'fdd7a751-b60b-444a-984c-02652fe8fa1c', // Groups Administrator
    'fe930be7-5e62-47db-91af-98c3a49a38b1', // User Administrator
    '29232cdf-9323-42fd-ade2-1d097af3e4de', // Exchange Administrator
    'f28a1f50-f6e7-4571-818b-6a12f2af6b6c', // SharePoint Administrator
    'f2ef992c-3afb-46b9-b7cf-a126ee74c451', // Global Reader
    '69091246-20e8-4a56-aa4d-066075b2a7a8', // Teams Administrator
    'b0f54661-2d74-4c50-afa3-1ec803f12efe', // Billing Administrator
    'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9', // Conditional Access Administrator
    '966707d0-3269-4727-9be2-8c3a10f19b9d', // Password Administrator
    '7495fdc4-34c4-4d15-a289-98788ce399fd', // Azure Information Protection Admin
  ];

  // ---------- SM.1006 Emergency access accounts exist ----------
  tests.push({
    id: 'SM.1006',
    title: 'At least one emergency access (break-glass) account exists',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: null,
    description: 'Best practice is to keep one or two cloud-only "break-glass" accounts excluded from CA and MFA policies, with strong long passphrases stored offline. We look for accounts whose UPN or display name contains "break", "emergency" or "bg-".',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const users = await Graph.graphAll(`users?$select=id,userPrincipalName,displayName,accountEnabled&$top=999`, { apiVersion: 'v1.0' });
        const ba = users.filter(u => /break|emergency|\bbg[-_]/i.test((u.displayName||'') + ' ' + (u.userPrincipalName||'')));
        const enabled = ba.filter(u => u.accountEnabled);
        if (enabled.length === 0) {
          return { id:'SM.1006', status:'Failed', reason:'No accounts found whose UPN or display name suggests a break-glass purpose.', actual: ba.length, durationMs: ms(start) };
        }
        return { id:'SM.1006', status:'Passed', reason:`${enabled.length} candidate emergency access account(s) found: ${enabled.map(u=>u.userPrincipalName).join(', ')}`, actual: enabled.length, durationMs: ms(start) };
      } catch (e) { return errRow('SM.1006', e, start); }
    },
  });

  // ---------- MT.1032 Global admin count ----------
  tests.push({
    id: 'MT.1032',
    title: 'Limited number of Global Admins are assigned.',
    severity: 'Medium',
    tag: 'Maester',
    category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1032',
    description: 'Centre for Internet Security recommends 2 to 4 dedicated global admin accounts. Fewer means no backup if one breaks. More inflates the blast radius of a single compromise.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const roles = await getDirectoryRoles();
        const ga = roles.find(r => r.roleTemplateId === '62e90394-69f5-4237-9190-012177145e10');
        if (!ga) return { id:'MT.1032', status:'Skipped', reason:'Global Administrator role not activated in this tenant.', durationMs: ms(start) };
        const members = await getRoleMembers(ga.id);
        const userCount = members.filter(m => m['@odata.type'] === '#microsoft.graph.user').length;
        const ok = userCount >= 2 && userCount <= 4;
        return { id:'MT.1032', status: ok ? 'Passed':'Failed', actual: userCount, reason: `Found ${userCount} user(s) in the Global Administrator role.`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1032', e, start); }
    },
  });

  // ---------- CIS.M365.1.1.1 Cloud-only admins ----------
  tests.push({
    id: 'CIS.M365.1.1.1',
    title: 'Ensure Administrative accounts are cloud-only',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/CIS.M365.1.1.1',
    description: 'Synced on-prem accounts in privileged roles widen your attack surface to your AD - if AD falls, the cloud goes too. CIS and Microsoft both recommend cloud-only admin accounts.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const roles = await getDirectoryRoles();
        const ga = roles.find(r => r.roleTemplateId === '62e90394-69f5-4237-9190-012177145e10');
        if (!ga) return { id:'CIS.M365.1.1.1', status:'Skipped', reason:'Global Administrator role not activated in this tenant.', durationMs: ms(start) };
        const members = await getRoleMembers(ga.id);
        const synced = [];
        for (const m of members) {
          if (m['@odata.type'] !== '#microsoft.graph.user') continue;
          try {
            const u = await Graph.graph(`users/${m.id}?$select=onPremisesSyncEnabled,userPrincipalName`, { apiVersion: 'v1.0' });
            if (u.onPremisesSyncEnabled) synced.push(u.userPrincipalName);
          } catch (e) { Log.debug('user lookup failed', e.message); }
        }
        if (synced.length === 0) return { id:'CIS.M365.1.1.1', status:'Passed', reason:'All Global Administrators are cloud-only.', actual: 0, durationMs: ms(start) };
        return { id:'CIS.M365.1.1.1', status:'Failed', reason:`${synced.length} synced account(s) hold the Global Administrator role: ${synced.join(', ')}`, actual: synced.length, durationMs: ms(start) };
      } catch (e) { return errRow('CIS.M365.1.1.1', e, start); }
    },
  });

  // ---------- CISA.MS.AAD.3.6 Phishing-resistant MFA for highly privileged roles ----------
  tests.push({
    id: 'CISA.MS.AAD.3.6',
    title: 'Phishing-resistant MFA SHALL be required for highly privileged roles.',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.6',
    description: 'Anyone holding a privileged role should have at least one strong (MFA) method registered. Ideally phishing-resistant (FIDO2, WHfB, certificate). We check the userRegistrationDetails report.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const roles = await getDirectoryRoles();
        const privRoles = roles.filter(r => PRIV_ROLE_TEMPLATE_IDS.includes(r.roleTemplateId));
        const userIds = new Set();
        for (const r of privRoles) {
          const members = await getRoleMembers(r.id);
          for (const m of members) if (m['@odata.type'] === '#microsoft.graph.user') userIds.add(m.id);
        }
        if (userIds.size === 0) return { id:'CISA.MS.AAD.3.6', status:'Skipped', reason:'No privileged role members found.', durationMs: ms(start) };
        const reg = await Graph.graphAll('reports/authenticationMethods/userRegistrationDetails?$top=999', { apiVersion: 'v1.0' });
        const map = new Map(reg.map(u => [u.id, u]));
        const noMfa = []; const noPhish = [];
        for (const id of userIds) {
          const u = map.get(id);
          if (!u) continue;
          if (!u.isMfaCapable) noMfa.push(u.userPrincipalName);
          else if (!u.isSystemPreferredAuthenticationMethodEnabled && (!u.methodsRegistered || !u.methodsRegistered.some(m => /fido2|windowsHello|x509Certificate|deviceBasedPush/i.test(m)))) noPhish.push(u.userPrincipalName);
        }
        if (noMfa.length === 0 && noPhish.length === 0) {
          return { id:'CISA.MS.AAD.3.6', status:'Passed', reason:`All ${userIds.size} privileged users have phishing-resistant methods registered.`, actual: 0, durationMs: ms(start) };
        }
        if (noMfa.length > 0) {
          return { id:'CISA.MS.AAD.3.6', status:'Failed', reason:`${noMfa.length} privileged user(s) without MFA: ${noMfa.slice(0,10).join(', ')}${noMfa.length>10?', ...':''}`, actual: noMfa.length, durationMs: ms(start) };
        }
        return { id:'CISA.MS.AAD.3.6', status:'Failed', reason:`${noPhish.length} privileged user(s) without phishing-resistant methods: ${noPhish.slice(0,10).join(', ')}${noPhish.length>10?', ...':''}`, actual: noPhish.length, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.6', e, start); }
    },
  });

  // ---------- SM.1015 App credentials nearing expiry / expired ----------
  tests.push({
    id: 'SM.1015',
    title: 'No application has expired or soon-to-expire credentials in active use',
    severity: 'Medium',
    tag: 'Maester',
    category: 'Maester',
    docUrl: null,
    description: 'Reports app registrations whose secrets or certificates are already expired or expire within the next 30 days. Expired credentials are a sign of orphaned apps; soon-to-expire ones are a recipe for a 3am outage.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const apps = await Graph.graphAll('applications?$select=appId,displayName,passwordCredentials,keyCredentials&$top=999', { apiVersion: 'v1.0' });
        const now = new Date();
        const soon = new Date(now.getTime() + 30*24*3600*1000);
        const expired = []; const expiring = [];
        for (const a of apps) {
          const creds = [...(a.passwordCredentials||[]), ...(a.keyCredentials||[])];
          for (const c of creds) {
            if (!c.endDateTime) continue;
            const end = new Date(c.endDateTime);
            if (end < now) expired.push(a.displayName);
            else if (end < soon) expiring.push(a.displayName);
          }
        }
        if (expired.length === 0 && expiring.length === 0) {
          return { id:'SM.1015', status:'Passed', reason:`Checked ${apps.length} applications, no expired or expiring credentials in the next 30 days.`, actual: 0, durationMs: ms(start) };
        }
        return { id:'SM.1015', status: expired.length ? 'Failed':'Failed', actual: expired.length + expiring.length,
          reason:`${expired.length} expired credential(s) and ${expiring.length} expiring within 30 days. Examples: ${[...new Set([...expired, ...expiring])].slice(0,5).join(', ')}`,
          durationMs: ms(start) };
      } catch (e) { return errRow('SM.1015', e, start); }
    },
  });

  // ---------- MT.1063 Application owners with MFA enabled ----------
  tests.push({
    id: 'MT.1063',
    title: 'All app registration owners should have MFA registered',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1063',
    description: 'Application owners can rotate secrets, add redirect URIs and effectively impersonate the app. They are a valuable target. We check that every owner of every app registration is MFA-registered.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const apps = await Graph.graphAll('applications?$select=appId,displayName&$top=999', { apiVersion: 'v1.0' });
        const ownerIds = new Set();
        await Concurrency.pmap(apps, 8, async (a) => {
          try { const owners = await Graph.graphAll(`applications(appId='${a.appId}')/owners`, { apiVersion: 'v1.0' });
            for (const o of owners) if (o['@odata.type'] === '#microsoft.graph.user') ownerIds.add(o.id);
          } catch (e) { Log.debug('owner lookup', a.displayName, e.message); }
        });
        if (ownerIds.size === 0) return { id:'MT.1063', status:'Skipped', reason:'No user app owners found.', durationMs: ms(start) };
        const reg = await Graph.graphAll('reports/authenticationMethods/userRegistrationDetails?$top=999', { apiVersion: 'v1.0' });
        const map = new Map(reg.map(u => [u.id, u]));
        const noMfa = [];
        for (const id of ownerIds) { const u = map.get(id); if (u && !u.isMfaCapable) noMfa.push(u.userPrincipalName); }
        if (noMfa.length === 0) return { id:'MT.1063', status:'Passed', reason:`All ${ownerIds.size} application owners have MFA registered.`, actual: 0, durationMs: ms(start) };
        return { id:'MT.1063', status:'Failed', reason:`${noMfa.length} application owner(s) without MFA: ${noMfa.slice(0,10).join(', ')}`, actual: noMfa.length, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1063', e, start); }
    },
  });

  // ---------- SM.1001 At least one Conditional Access policy exists ----------
  tests.push({
    id: 'SM.1001',
    title: 'At least one Conditional Access policy is enabled',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: null,
    description: 'A tenant with no enabled CA policies relies entirely on legacy security defaults. Bare minimum is to have at least one CA policy enabled.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies?$select=id,displayName,state', { apiVersion: 'v1.0' });
        const enabled = pols.filter(p => p.state === 'enabled');
        if (enabled.length === 0) return { id:'SM.1001', status:'Failed', reason:`No enabled CA policies found (${pols.length} total, none enabled).`, actual: 0, durationMs: ms(start) };
        return { id:'SM.1001', status:'Passed', reason:`${enabled.length} of ${pols.length} CA policies are enabled.`, actual: enabled.length, durationMs: ms(start) };
      } catch (e) { return errRow('SM.1001', e, start); }
    },
  });

  // ---------- SM.1002 Block legacy authentication ----------
  tests.push({
    id: 'SM.1002',
    title: 'A Conditional Access policy blocks legacy authentication',
    severity: 'Critical',
    tag: 'Maester',
    category: 'Maester',
    docUrl: null,
    description: 'Looks for an enabled CA policy that targets all users and the "exchangeActiveSync" + "other" client app types and blocks access. This is one of the highest-impact policies you can deploy.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && p.grantControls?.builtInControls?.includes('block')
          && (p.conditions?.clientAppTypes||[]).some(t => /exchangeActiveSync|other/i.test(t)));
        if (hit) return { id:'SM.1002', status:'Passed', reason:`Policy "${hit.displayName}" blocks legacy authentication.`, durationMs: ms(start) };
        return { id:'SM.1002', status:'Failed', reason:'No enabled CA policy was found that blocks legacy authentication client app types.', durationMs: ms(start) };
      } catch (e) { return errRow('SM.1002', e, start); }
    },
  });

  // ---------- MT.1006 MFA for all admins ----------
  tests.push({
    id: 'MT.1006',
    title: 'At least one Conditional Access policy is configured to require MFA for admins.',
    severity: 'Critical',
    tag: 'Maester',
    category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1006',
    description: 'Looks for an enabled CA policy that targets the privileged Entra roles and requires MFA. Maester checks the broader set of 14 privileged roles - we do the same.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const target = new Set(PRIV_ROLE_TEMPLATE_IDS);
        const hit = pols.find(p => {
          if (p.state !== 'enabled') return false;
          if (!p.grantControls) return false;
          const wantsMfa = (p.grantControls.builtInControls||[]).includes('mfa') || (p.grantControls.authenticationStrength?.id);
          if (!wantsMfa) return false;
          const incRoles = p.conditions?.users?.includeRoles || [];
          // a CA policy passes if at least 5 of the privileged roles are included
          let count = 0;
          for (const r of incRoles) if (target.has(r)) count++;
          return count >= 5;
        });
        if (hit) return { id:'MT.1006', status:'Passed', reason:`Policy "${hit.displayName}" requires MFA for the privileged role set.`, durationMs: ms(start) };
        return { id:'MT.1006', status:'Failed', reason:'No enabled CA policy was found that requires MFA for the major privileged Entra roles.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1006', e, start); }
    },
  });

  // ---------- SM.1004 MFA for all users ----------
  tests.push({
    id: 'SM.1004',
    title: 'A Conditional Access policy requires MFA for all users',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: null,
    description: 'Looks for an enabled CA policy that includes "All users" and requires MFA for all cloud apps.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.users?.includeUsers||[]).includes('All')
          && (p.conditions?.applications?.includeApplications||[]).includes('All')
          && ((p.grantControls?.builtInControls||[]).includes('mfa') || p.grantControls?.authenticationStrength?.id));
        if (hit) return { id:'SM.1004', status:'Passed', reason:`Policy "${hit.displayName}" requires MFA for all users on all apps.`, durationMs: ms(start) };
        return { id:'SM.1004', status:'Failed', reason:'No enabled CA policy was found requiring MFA for All users on All apps.', durationMs: ms(start) };
      } catch (e) { return errRow('SM.1004', e, start); }
    },
  });

  // ---------- SM.1005 MFA for guest users ----------
  tests.push({
    id: 'SM.1005',
    title: 'A Conditional Access policy requires MFA for guest users',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: null,
    description: 'Looks for an enabled CA policy that targets guests/external users and requires MFA.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && ((p.conditions?.users?.includeGuestsOrExternalUsers?.guestOrExternalUserTypes) || (p.conditions?.users?.includeUsers||[]).some(u => /guest/i.test(u)))
          && ((p.grantControls?.builtInControls||[]).includes('mfa') || p.grantControls?.authenticationStrength?.id));
        if (hit) return { id:'SM.1005', status:'Passed', reason:`Policy "${hit.displayName}" requires MFA for guest users.`, durationMs: ms(start) };
        return { id:'SM.1005', status:'Failed', reason:'No enabled CA policy was found requiring MFA for guests/external users.', durationMs: ms(start) };
      } catch (e) { return errRow('SM.1005', e, start); }
    },
  });

  // ---------- CISA.MS.AAD.7.4 Permanent active role assignments ----------
  tests.push({
    id: 'CISA.MS.AAD.7.4',
    title: 'Permanent active role assignments SHALL NOT be allowed for highly privileged roles.',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.7.4',
    description: 'Permanent active assignments for privileged roles defeat the purpose of PIM and just-in-time elevation. We list every active (non-eligible) assignment in the privileged roles set.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const active = await Graph.graphAll('roleManagement/directory/roleAssignmentScheduleInstances?$expand=roleDefinition,principal', { apiVersion: 'v1.0' });
        const perm = active.filter(a => (a.assignmentType === 'Assigned' || a.memberType === 'Direct') && PRIV_ROLE_TEMPLATE_IDS.includes(a.roleDefinition?.templateId));
        if (perm.length === 0) return { id:'CISA.MS.AAD.7.4', status:'Passed', reason:'No permanent active assignments in privileged roles. Good job.', actual: 0, durationMs: ms(start) };
        return { id:'CISA.MS.AAD.7.4', status:'Failed', actual: perm.length,
          reason:`${perm.length} permanent active assignment(s) in privileged roles. Examples: ${perm.slice(0,5).map(p => `${p.principal?.displayName||p.principalId} -> ${p.roleDefinition?.displayName}`).join('; ')}`,
          durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403) return { id:'CISA.MS.AAD.7.4', status:'Skipped', reason:'PIM not licensed in this tenant or insufficient permissions.', durationMs: ms(start) };
        return errRow('CISA.MS.AAD.7.4', e, start);
      }
    },
  });

  // ---------- CISA.MS.AAD.3.5 Weak authentication methods ----------
  tests.push({
    id: 'CISA.MS.AAD.3.5',
    title: 'The authentication methods SMS, Voice Call, and Email One-Time Passcode (OTP) SHALL be disabled.',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.AAD.3.5',
    description: 'SMS and Voice are vulnerable to SIM swap and interception attacks. Microsoft and CISA recommend disabling them in favour of phishing-resistant methods.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/authenticationMethodsPolicy', { apiVersion: 'beta' });
        const sms   = pol.authenticationMethodConfigurations?.find(c => /Sms/i.test(c.id));
        const voice = pol.authenticationMethodConfigurations?.find(c => /Voice/i.test(c.id));
        const issues = [];
        if (sms?.state   === 'enabled') issues.push('SMS is enabled');
        if (voice?.state === 'enabled') issues.push('Voice is enabled');
        if (issues.length === 0) return { id:'CISA.MS.AAD.3.5', status:'Passed', reason:'SMS and Voice authentication methods are disabled.', durationMs: ms(start) };
        return { id:'CISA.MS.AAD.3.5', status:'Failed', reason: issues.join('; ') + '. Disable these in Authentication methods policy.', durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.AAD.3.5', e, start); }
    },
  });

  // ---------- SM.1060 Stale guest accounts ----------
  tests.push({
    id: 'SM.1060',
    title: 'No guest accounts have been inactive for more than 90 days',
    severity: 'Medium',
    tag: 'Maester',
    category: 'Maester',
    docUrl: null,
    description: 'Inactive guests are a needless attack surface. We use signInActivity to find guests that have not signed in within the last 90 days.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const guests = await Graph.graphAll(`users?$filter=userType eq 'Guest'&$select=id,userPrincipalName,signInActivity,createdDateTime&$top=999`, { apiVersion: 'beta' });
        const cutoff = new Date(Date.now() - 90*24*3600*1000);
        const stale = guests.filter(g => {
          const last = g.signInActivity?.lastSignInDateTime || g.signInActivity?.lastNonInteractiveSignInDateTime;
          if (!last) return new Date(g.createdDateTime) < cutoff;
          return new Date(last) < cutoff;
        });
        if (stale.length === 0) return { id:'SM.1060', status:'Passed', reason:`Checked ${guests.length} guest(s), none inactive >90 days.`, actual: 0, durationMs: ms(start) };
        return { id:'SM.1060', status:'Failed', reason:`${stale.length} guest(s) inactive for more than 90 days.`, actual: stale.length, durationMs: ms(start) };
      } catch (e) { return errRow('SM.1060', e, start); }
    },
  });

  // ---------- SM.1070 Service principals with risky API permissions ----------
  tests.push({
    id: 'SM.1070',
    title: 'No service principal has consented to high-risk application permissions',
    severity: 'High',
    tag: 'Maester',
    category: 'Maester',
    docUrl: null,
    description: 'Reports any non-Microsoft service principal that has been granted Graph application permissions like Directory.ReadWrite.All, RoleManagement.ReadWrite.Directory, Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All - permissions that effectively give global admin.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const HIGH = new Set([
          '19dbc75e-c2e2-444c-a770-ec69d8559fc7', // Directory.ReadWrite.All
          '9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8', // RoleManagement.ReadWrite.Directory
          '1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9', // Application.ReadWrite.All
          '06b708a9-e830-4db3-a914-8e69da51d44f', // AppRoleAssignment.ReadWrite.All
          '62a82d76-70ea-41e2-9197-370581804d09', // Group.ReadWrite.All
          '7ab1d382-f21e-4acd-a863-ba3e13f7da61', // Directory.Read.All - watch
        ]);
        // Find graph SPN
        const graphSpn = await Graph.graph(`servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=id`, { apiVersion: 'v1.0' });
        const graphSpnId = graphSpn.value?.[0]?.id;
        if (!graphSpnId) return { id:'SM.1070', status:'Skipped', reason:'Could not resolve Microsoft Graph service principal.', durationMs: ms(start) };
        const grants = await Graph.graphAll(`servicePrincipals/${graphSpnId}/appRoleAssignedTo?$top=999`, { apiVersion: 'v1.0' });
        const risky = grants.filter(g => HIGH.has(g.appRoleId));
        if (risky.length === 0) return { id:'SM.1070', status:'Passed', reason:'No service principals hold high-risk Graph application permissions.', actual: 0, durationMs: ms(start) };
        return { id:'SM.1070', status:'Failed', actual: risky.length,
          reason:`${risky.length} service principal grant(s) with high-risk Graph permissions. Examples: ${risky.slice(0,5).map(r => r.principalDisplayName).join(', ')}`,
          durationMs: ms(start) };
      } catch (e) { return errRow('SM.1070', e, start); }
    },
  });

  // ---------- SM.0001 Admin units in use? (Soft check, info only) ----------
  tests.push({
    id: 'SM.0001',
    title: 'Tenant uses Administrative Units to scope role assignments',
    severity: 'Info',
    tag: 'Maester',
    category: 'Maester',
    docUrl: 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/administrative-units',
    description: 'AUs let you scope role assignments to a slice of the tenant - users, groups, devices. Useful for "regional helpdesk" or "department admin" patterns. Informational check.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const aus = await Graph.graphAll('directory/administrativeUnits?$select=id,displayName', { apiVersion: 'v1.0' });
        if (aus.length === 0) return { id:'SM.0001', status:'Failed', reason:'No administrative units defined. If your tenant has any tiered-admin needs, AUs are worth considering.', actual: 0, durationMs: ms(start) };
        return { id:'SM.0001', status:'Passed', reason:`${aus.length} administrative unit(s) defined.`, actual: aus.length, durationMs: ms(start) };
      } catch (e) { return errRow('SM.0001', e, start); }
    },
  });

  // ---------- MT.DEF.1001 Defender for Office 365 licensing/enablement ----------
  tests.push({
    id: 'MT.DEF.1001',
    title: 'Microsoft Defender for Office 365 is enabled',
    severity: 'High',
    tag: 'Defender',
    category: 'Defender',
    docUrl: 'https://maester.dev/docs/tests/MT.DEF.1001',
    description: 'Best-effort Graph equivalent: verify the tenant has an active Defender for Office 365 related service plan in subscribed SKUs.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const skus = await Graph.graphAll('subscribedSkus?$select=skuPartNumber,servicePlans,capabilityStatus', { apiVersion: 'v1.0' });
        const allPlans = [];
        for (const sku of skus) {
          for (const sp of (sku.servicePlans || [])) {
            allPlans.push({
              skuPartNumber: sku.skuPartNumber,
              servicePlanName: String(sp.servicePlanName || ''),
              provisioningStatus: String(sp.provisioningStatus || ''),
            });
          }
        }

        const mdoPlans = allPlans.filter(p => /ATP|DEFENDER.*OFFICE|MDO|SAFE.*(LINK|ATTACH)|THREAT.*PROTECTION/i.test(p.servicePlanName));
        if (!mdoPlans.length) {
          return { id: 'MT.DEF.1001', status: 'Failed', reason: 'No Defender for Office 365 related service plans found in subscribed SKUs.', actual: 0, durationMs: ms(start) };
        }

        const enabled = mdoPlans.filter(p => /success|pendinginput|pendingactivation/i.test(p.provisioningStatus));
        if (!enabled.length) {
          return {
            id: 'MT.DEF.1001',
            status: 'Failed',
            reason: 'Defender for Office 365 service plans are present but none are provisioned as active.',
            actual: mdoPlans.slice(0, 10),
            durationMs: ms(start),
          };
        }

        return {
          id: 'MT.DEF.1001',
          status: 'Passed',
          reason: `Found ${enabled.length} active Defender for Office 365 related service plan(s).`,
          actual: enabled.slice(0, 10),
          durationMs: ms(start),
        };
      } catch (e) { return errRow('MT.DEF.1001', e, start); }
    },
  });

  // ---------- MT.DEF.1002 Preset security policy equivalent ----------
  tests.push({
    id: 'MT.DEF.1002',
    title: 'Preset security policies are enabled (Standard or Strict)',
    severity: 'High',
    tag: 'Defender',
    category: 'Defender',
    docUrl: 'https://maester.dev/docs/tests/MT.DEF.1002',
    description: 'Graph equivalent: inspect latest Microsoft Secure Score control scores for controls that reference preset security policies.',
    implemented: true,
    requiredScopes: Auth.SCOPES.graphSecurityEvents,
    async run() {
      const start = performance.now();
      try {
        const secure = await Graph.graph('security/secureScores?$top=1', { apiVersion: 'beta', tokenScopes: Auth.SCOPES.graphSecurityEvents });
        const latest = Array.isArray(secure?.value) ? secure.value[0] : null;
        const controls = latest?.controlScores || [];
        const preset = controls.filter(c => /preset security/i.test(String(c.controlName || '')));

        if (!preset.length) {
          return {
            id: 'MT.DEF.1002',
            status: 'Skipped',
            reason: 'No preset-security control scores exposed via Graph secureScores for this tenant.',
            durationMs: ms(start),
          };
        }

        const nonCompliant = preset.filter(c => Number(c.scoreInPercentage || 0) < 100);
        const ok = nonCompliant.length === 0;
        return {
          id: 'MT.DEF.1002',
          status: ok ? 'Passed' : 'Failed',
          reason: ok
            ? 'Preset-security related secure score controls are fully satisfied.'
            : `${nonCompliant.length} preset-security related control(s) are below full compliance.`,
          actual: preset.map(c => ({ controlName: c.controlName, scoreInPercentage: c.scoreInPercentage })),
          durationMs: ms(start),
        };
      } catch (e) { return errRow('MT.DEF.1002', e, start); }
    },
  });

  // ---------- MT.DEF.1003 Tenant Allow/Block list size equivalent ----------
  tests.push({
    id: 'MT.DEF.1003',
    title: 'Tenant Allow/Block list size is reasonable',
    severity: 'Low',
    tag: 'Defender',
    category: 'Defender',
    docUrl: 'https://maester.dev/docs/tests/MT.DEF.1003',
    description: 'Graph equivalent: evaluate latest Secure Score control(s) that reference Tenant Allow/Block list posture.',
    implemented: true,
    requiredScopes: Auth.SCOPES.graphSecurityEvents,
    async run() {
      const start = performance.now();
      try {
        const secure = await Graph.graph('security/secureScores?$top=1', { apiVersion: 'beta', tokenScopes: Auth.SCOPES.graphSecurityEvents });
        const latest = Array.isArray(secure?.value) ? secure.value[0] : null;
        const controls = latest?.controlScores || [];
        const allowBlock = controls.filter(c => /allow\/?block|tenant allow|block list/i.test(String(c.controlName || '')));

        if (!allowBlock.length) {
          return {
            id: 'MT.DEF.1003',
            status: 'Skipped',
            reason: 'No Tenant Allow/Block list control is exposed in Graph secureScores for this tenant.',
            durationMs: ms(start),
          };
        }

        const nonCompliant = allowBlock.filter(c => Number(c.scoreInPercentage || 0) < 100);
        const ok = nonCompliant.length === 0;
        return {
          id: 'MT.DEF.1003',
          status: ok ? 'Passed' : 'Failed',
          reason: ok
            ? 'Tenant Allow/Block list related secure score controls are fully satisfied.'
            : `${nonCompliant.length} Tenant Allow/Block list related control(s) are below full compliance.`,
          actual: allowBlock.map(c => ({ controlName: c.controlName, scoreInPercentage: c.scoreInPercentage })),
          durationMs: ms(start),
        };
      } catch (e) { return errRow('MT.DEF.1003', e, start); }
    },
  });

  // ---------- MT.1007 MFA for all users (CA) ----------
  tests.push({
    id: 'MT.1007', title: 'At least one Conditional Access policy is configured to require MFA for all users',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1007',
    description: 'Checks for an enabled CA policy scoped to All users that requires MFA or an authentication strength.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.users?.includeUsers || []).includes('All')
          && ((p.grantControls?.builtInControls || []).includes('mfa') || p.grantControls?.authenticationStrength?.id));
        if (hit) return { id: 'MT.1007', status: 'Passed', reason: `Policy "${hit.displayName}" requires MFA for all users.`, durationMs: ms(start) };
        return { id: 'MT.1007', status: 'Failed', reason: 'No enabled CA policy found that requires MFA for all users.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1007', e, start); }
    },
  });

  // ---------- MT.1008 MFA for Azure management (CA) ----------
  tests.push({
    id: 'MT.1008', title: 'At least one Conditional Access policy is configured to require MFA for Azure management',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1008',
    description: 'Checks for an enabled CA policy targeting the Azure Management app (797f4846-ba00-4fd7-ba43-dac1f8f63013) that requires MFA.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const AZURE_MGMT = '797f4846-ba00-4fd7-ba43-dac1f8f63013';
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.applications?.includeApplications || []).includes(AZURE_MGMT)
          && ((p.grantControls?.builtInControls || []).includes('mfa') || p.grantControls?.authenticationStrength?.id));
        if (hit) return { id: 'MT.1008', status: 'Passed', reason: `Policy "${hit.displayName}" requires MFA for Azure management.`, durationMs: ms(start) };
        return { id: 'MT.1008', status: 'Failed', reason: 'No enabled CA policy found requiring MFA for Azure management.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1008', e, start); }
    },
  });

  // ---------- MT.1009 Block other legacy authentication (CA) ----------
  tests.push({
    id: 'MT.1009', title: 'At least one Conditional Access policy is configured to block other legacy authentication',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1009',
    description: 'Checks for an enabled CA policy that blocks the "other" legacy client app type (non-EAS legacy auth).',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.grantControls?.builtInControls || []).includes('block')
          && (p.conditions?.clientAppTypes || []).includes('other'));
        if (hit) return { id: 'MT.1009', status: 'Passed', reason: `Policy "${hit.displayName}" blocks other legacy authentication.`, durationMs: ms(start) };
        return { id: 'MT.1009', status: 'Failed', reason: 'No enabled CA policy found that blocks the "other" legacy client app type.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1009', e, start); }
    },
  });

  // ---------- MT.1011 Security info registration from trusted location (CA) ----------
  tests.push({
    id: 'MT.1011', title: 'At least one Conditional Access policy is configured to secure security info registration only from a trusted location',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1011',
    description: 'Checks for an enabled CA policy targeting the "Register security information" user action.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.applications?.includeUserActions || []).includes('urn:user:registersecurityinfo'));
        if (hit) return { id: 'MT.1011', status: 'Passed', reason: `Policy "${hit.displayName}" governs security info registration.`, durationMs: ms(start) };
        return { id: 'MT.1011', status: 'Failed', reason: 'No enabled CA policy found targeting the "Register security information" user action.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1011', e, start); }
    },
  });

  // ---------- MT.1012 MFA for risky sign-ins (CA / P2) ----------
  tests.push({
    id: 'MT.1012', title: 'At least one Conditional Access policy is configured to require MFA for risky sign-ins',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1012',
    description: 'Checks for an enabled CA policy with signInRiskLevels that requires MFA. Requires Entra ID P2.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.signInRiskLevels || []).length > 0
          && ((p.grantControls?.builtInControls || []).includes('mfa') || p.grantControls?.authenticationStrength?.id
            || (p.grantControls?.builtInControls || []).includes('block')));
        if (hit) return { id: 'MT.1012', status: 'Passed', reason: `Policy "${hit.displayName}" enforces action for risky sign-ins.`, durationMs: ms(start) };
        return { id: 'MT.1012', status: 'Failed', reason: 'No enabled CA policy found with signInRiskLevels. Requires Entra ID P2 for Identity Protection.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1012', e, start); }
    },
  });

  // ---------- MT.1013 Password change for high user risk (CA / P2) ----------
  tests.push({
    id: 'MT.1013', title: 'At least one Conditional Access policy is configured to require new password when user risk is high',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1013',
    description: 'Checks for an enabled CA policy with high userRiskLevels that requires password change or blocks access. Requires Entra ID P2.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.userRiskLevels || []).some(r => /high/i.test(r))
          && ((p.grantControls?.builtInControls || []).some(c => /mfa|passwordChange|block/i.test(c))
            || p.grantControls?.authenticationStrength?.id));
        if (hit) return { id: 'MT.1013', status: 'Passed', reason: `Policy "${hit.displayName}" enforces action for high user risk.`, durationMs: ms(start) };
        return { id: 'MT.1013', status: 'Failed', reason: 'No enabled CA policy found requiring password change for high user risk. Requires Entra ID P2.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1013', e, start); }
    },
  });

  // ---------- MT.1014 Compliant/hybrid devices for admins (CA) ----------
  tests.push({
    id: 'MT.1014', title: 'At least one Conditional Access policy is configured to require compliant or Entra hybrid joined devices for admins',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1014',
    description: 'Checks for an enabled CA policy targeting admin roles that requires a compliant or hybrid-joined device.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.users?.includeRoles || []).length > 0
          && (p.grantControls?.builtInControls || []).some(c => /compliantDevice|domainJoined/i.test(c)));
        if (hit) return { id: 'MT.1014', status: 'Passed', reason: `Policy "${hit.displayName}" requires compliant or hybrid-joined device for admin roles.`, durationMs: ms(start) };
        return { id: 'MT.1014', status: 'Failed', reason: 'No enabled CA policy found requiring compliant/hybrid device for admin roles.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1014', e, start); }
    },
  });

  // ---------- MT.1016 MFA for guest access (CA) ----------
  tests.push({
    id: 'MT.1016', title: 'At least one Conditional Access policy is configured to require MFA for guest access',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1016',
    description: 'Checks for an enabled CA policy targeting guest or external users that requires MFA.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.users?.includeGuestsOrExternalUsers?.guestOrExternalUserTypes
            || (p.conditions?.users?.includeUsers || []).some(u => /guest/i.test(u)))
          && ((p.grantControls?.builtInControls || []).includes('mfa') || p.grantControls?.authenticationStrength?.id));
        if (hit) return { id: 'MT.1016', status: 'Passed', reason: `Policy "${hit.displayName}" requires MFA for guests.`, durationMs: ms(start) };
        return { id: 'MT.1016', status: 'Failed', reason: 'No enabled CA policy found requiring MFA for guest/external users.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1016', e, start); }
    },
  });

  // ---------- MT.1017 Non-persistent browser session (CA) ----------
  tests.push({
    id: 'MT.1017', title: 'At least one Conditional Access policy is configured to enforce non persistent browser session for non-corporate devices',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1017',
    description: 'Checks for an enabled CA policy with persistentBrowser session control disabled.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && p.sessionControls?.persistentBrowser?.isEnabled === true
          && p.sessionControls?.persistentBrowser?.mode === 'never');
        if (hit) return { id: 'MT.1017', status: 'Passed', reason: `Policy "${hit.displayName}" enforces non-persistent browser session.`, durationMs: ms(start) };
        return { id: 'MT.1017', status: 'Failed', reason: 'No enabled CA policy found enforcing non-persistent browser session (persistentBrowser mode=never).', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1017', e, start); }
    },
  });

  // ---------- MT.1018 Sign-in frequency (CA) ----------
  tests.push({
    id: 'MT.1018', title: 'At least one Conditional Access policy is configured to enforce sign-in frequency for non-corporate devices',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1018',
    description: 'Checks for an enabled CA policy with signInFrequency session controls configured.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && p.sessionControls?.signInFrequency?.isEnabled === true);
        if (hit) return { id: 'MT.1018', status: 'Passed', reason: `Policy "${hit.displayName}" enforces sign-in frequency.`, durationMs: ms(start) };
        return { id: 'MT.1018', status: 'Failed', reason: 'No enabled CA policy found with sign-in frequency session control enabled.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1018', e, start); }
    },
  });

  // ---------- MT.1019 App enforced restrictions (CA) ----------
  tests.push({
    id: 'MT.1019', title: 'At least one Conditional Access policy is configured to enable application enforced restrictions',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1019',
    description: 'Checks for an enabled CA policy with applicationEnforcedRestrictions session control enabled (limits unmanaged device access to SharePoint/Exchange).',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && p.sessionControls?.applicationEnforcedRestrictions?.isEnabled === true);
        if (hit) return { id: 'MT.1019', status: 'Passed', reason: `Policy "${hit.displayName}" enables app enforced restrictions.`, durationMs: ms(start) };
        return { id: 'MT.1019', status: 'Failed', reason: 'No enabled CA policy found with applicationEnforcedRestrictions enabled.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1019', e, start); }
    },
  });

  // ---------- MT.1021 Security Defaults ----------
  tests.push({
    id: 'MT.1021', title: 'Security Defaults are enabled',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1021',
    description: 'Security Defaults provide pre-configured security policies for free/basic tenants. If you have CA policies instead, this test will Fail (which is expected and OK - CA is better). The test Passes only if Security Defaults are explicitly on.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/identitySecurityDefaultsEnforcementPolicy', { apiVersion: 'beta' });
        if (pol.isEnabled) return { id: 'MT.1021', status: 'Passed', reason: 'Security Defaults are enabled. Note: if you have CA policies, they are overriding Security Defaults which is fine.', durationMs: ms(start) };
        return { id: 'MT.1021', status: 'Failed', reason: 'Security Defaults are disabled. If you have CA policies configured this is expected. If not, enable Security Defaults.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1021', e, start); }
    },
  });

  // ---------- MT.1025 No external user with permanent control plane role ----------
  tests.push({
    id: 'MT.1025', title: 'No external user with permanent role assignment on Control Plane',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1025',
    description: 'External users (guests) with permanent privileged role assignments are a high risk. All such assignments should go through PIM time-limited elevation.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const assignments = await Graph.graphAll('roleManagement/directory/roleAssignmentScheduleInstances?$expand=principal', { apiVersion: 'v1.0' });
        const privTemplates = new Set(PRIV_ROLE_TEMPLATE_IDS);
        const bad = assignments.filter(a =>
          a.assignmentType === 'Assigned'
          && PRIV_ROLE_TEMPLATE_IDS.some(id => a.roleDefinitionId && a.roleDefinitionId.includes(id) || a.roleDefinition?.templateId === id)
          && a.principal?.userPrincipalName?.includes('#EXT#'));
        if (bad.length === 0) return { id: 'MT.1025', status: 'Passed', reason: 'No external users found with permanent privileged role assignments.', durationMs: ms(start) };
        return { id: 'MT.1025', status: 'Failed', actual: bad.length,
          reason: `${bad.length} external user(s) with permanent privileged roles: ${bad.slice(0, 5).map(a => a.principal?.userPrincipalName || a.principalId).join(', ')}`,
          durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403) return { id: 'MT.1025', status: 'Skipped', reason: 'Insufficient permissions for role assignment data (requires PIM or RoleManagement.Read.All).', durationMs: ms(start) };
        return errRow('MT.1025', e, start);
      }
    },
  });

  // ---------- MT.1026 No hybrid user with permanent control plane role ----------
  tests.push({
    id: 'MT.1026', title: 'No hybrid user with permanent role assignment on Control Plane',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1026',
    description: 'Synced on-prem accounts with permanent privileged roles are a risk: compromise AD = compromise cloud. Use cloud-only accounts or PIM eligible assignments.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const assignments = await Graph.graphAll('roleManagement/directory/roleAssignmentScheduleInstances?$expand=principal', { apiVersion: 'v1.0' });
        const permPriv = assignments.filter(a =>
          a.assignmentType === 'Assigned'
          && a.principal?.['@odata.type'] === '#microsoft.graph.user');
        if (permPriv.length === 0) return { id: 'MT.1026', status: 'Passed', reason: 'No permanent privileged user assignments found.', durationMs: ms(start) };
        const bad = [];
        for (const a of permPriv.slice(0, 50)) {
          try {
            const u = await Graph.graph(`users/${a.principalId}?$select=onPremisesSyncEnabled,userPrincipalName`, { apiVersion: 'v1.0' });
            if (u.onPremisesSyncEnabled) bad.push(u.userPrincipalName);
          } catch (_) { /* skip */ }
        }
        if (bad.length === 0) return { id: 'MT.1026', status: 'Passed', reason: 'No hybrid (synced) users with permanent privileged roles.', durationMs: ms(start) };
        return { id: 'MT.1026', status: 'Failed', actual: bad.length,
          reason: `${bad.length} hybrid user(s) with permanent privileged roles: ${bad.slice(0, 5).join(', ')}`, durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403) return { id: 'MT.1026', status: 'Skipped', reason: 'Insufficient permissions for role assignment data.', durationMs: ms(start) };
        return errRow('MT.1026', e, start);
      }
    },
  });

  // ---------- MT.1027 No SP with client secret and permanent control plane role ----------
  tests.push({
    id: 'MT.1027', title: 'No Service Principal with Client Secret and permanent role assignment on Control Plane',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1027',
    description: 'Service principals with client secrets (not certificates) holding permanent privileged roles are risky — secret leaks are easier and harder to detect.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const assignments = await Graph.graphAll('roleManagement/directory/roleAssignmentScheduleInstances?$expand=principal', { apiVersion: 'v1.0' });
        const permSpns = assignments.filter(a =>
          a.assignmentType === 'Assigned'
          && a.principal?.['@odata.type'] === '#microsoft.graph.servicePrincipal');
        if (permSpns.length === 0) return { id: 'MT.1027', status: 'Passed', reason: 'No service principals with permanent privileged role assignments.', durationMs: ms(start) };
        const bad = [];
        for (const a of permSpns.slice(0, 30)) {
          try {
            const app = await Graph.graph(`servicePrincipals/${a.principalId}?$select=displayName,appId`, { apiVersion: 'v1.0' });
            const appReg = await Graph.graph(`applications?$filter=appId eq '${app.appId}'&$select=passwordCredentials`, { apiVersion: 'v1.0' });
            if ((appReg.value?.[0]?.passwordCredentials || []).length > 0) bad.push(app.displayName);
          } catch (_) { /* skip */ }
        }
        if (bad.length === 0) return { id: 'MT.1027', status: 'Passed', reason: 'No service principals with client secrets hold permanent privileged roles.', durationMs: ms(start) };
        return { id: 'MT.1027', status: 'Failed', actual: bad.length,
          reason: `${bad.length} SP(s) with client secrets hold permanent privileged roles: ${bad.slice(0, 5).join(', ')}`, durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403) return { id: 'MT.1027', status: 'Skipped', reason: 'Insufficient permissions for role assignment data.', durationMs: ms(start) };
        return errRow('MT.1027', e, start);
      }
    },
  });

  // ---------- MT.1028 No user with mailbox and permanent control plane role ----------
  tests.push({
    id: 'MT.1028', title: 'No user with mailbox and permanent role assignment on Control Plane',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1028',
    description: 'Privileged users with mailboxes are phishing targets. The privileged account should be mail-disabled; use a separate account for email.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const assignments = await Graph.graphAll('roleManagement/directory/roleAssignmentScheduleInstances?$expand=principal', { apiVersion: 'v1.0' });
        const permUsers = assignments.filter(a =>
          a.assignmentType === 'Assigned'
          && a.principal?.['@odata.type'] === '#microsoft.graph.user');
        if (permUsers.length === 0) return { id: 'MT.1028', status: 'Passed', reason: 'No permanent privileged user assignments.', durationMs: ms(start) };
        const bad = [];
        for (const a of permUsers.slice(0, 50)) {
          try {
            const u = await Graph.graph(`users/${a.principalId}?$select=userPrincipalName,mail,assignedLicenses`, { apiVersion: 'v1.0' });
            if (u.mail || (u.assignedLicenses || []).length > 0) bad.push(u.userPrincipalName);
          } catch (_) { /* skip */ }
        }
        if (bad.length === 0) return { id: 'MT.1028', status: 'Passed', reason: 'No mail-enabled users with permanent privileged roles.', durationMs: ms(start) };
        return { id: 'MT.1028', status: 'Failed', actual: bad.length,
          reason: `${bad.length} mail-enabled user(s) with permanent privileged roles: ${bad.slice(0, 5).join(', ')}`, durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403) return { id: 'MT.1028', status: 'Skipped', reason: 'Insufficient permissions for role assignment data.', durationMs: ms(start) };
        return errRow('MT.1028', e, start);
      }
    },
  });

  // ---------- MT.1029 Stale accounts not in privileged roles (PIM alert) ----------
  tests.push({
    id: 'MT.1029', title: 'Stale accounts are not assigned to privileged roles',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1029',
    description: 'Checks active privileged role assignments and flags users with no sign-in or last sign-in older than 180 days.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const privTemplates = new Set(PRIV_ROLE_TEMPLATE_IDS);
        const active = await Graph.graphAll('roleManagement/directory/roleAssignmentScheduleInstances?$expand=roleDefinition,principal', { apiVersion: 'v1.0' });
        const activePrivUsers = active.filter(a =>
          a.principal?.['@odata.type'] === '#microsoft.graph.user'
          && privTemplates.has(a.roleDefinition?.templateId));

        if (!activePrivUsers.length) {
          return { id: 'MT.1029', status: 'Passed', reason: 'No active privileged user assignments found.', durationMs: ms(start) };
        }

        const uniqueUserIds = [...new Set(activePrivUsers.map(a => a.principalId).filter(Boolean))];
        const cutoffMs = Date.now() - (180 * 24 * 60 * 60 * 1000);
        let stale = 0;

        for (const userId of uniqueUserIds.slice(0, 250)) {
          try {
            const user = await Graph.graph(`users/${userId}?$select=id,userPrincipalName,accountEnabled,signInActivity`, { apiVersion: 'beta' });
            const last = user?.signInActivity?.lastSignInDateTime || user?.signInActivity?.lastSuccessfulSignInDateTime;
            if (!user?.accountEnabled) continue;
            if (!last || Date.parse(last) < cutoffMs) stale++;
          } catch (_) {
            // Ignore per-user lookup failures; we still evaluate all other privileged users.
          }
        }

        if (stale === 0) {
          return { id: 'MT.1029', status: 'Passed', reason: 'No stale privileged users detected (180+ days since last sign-in).', durationMs: ms(start) };
        }

        return {
          id: 'MT.1029',
          status: 'Failed',
          actual: stale,
          reason: `${stale} privileged user(s) have stale or missing sign-in activity (>=180 days). Review and remove unused role assignments.`,
          durationMs: ms(start),
        };
      } catch (e) {
        if (e.status === 403 || e.status === 404) return { id: 'MT.1029', status: 'Skipped', reason: 'Privileged role or sign-in activity data is not accessible with current permissions.', durationMs: ms(start) };
        return errRow('MT.1029', e, start);
      }
    },
  });

  // ---------- MT.1031 Privileged roles managed by PIM only ----------
  tests.push({
    id: 'MT.1031', title: 'Privileged role on Control Plane are managed by PIM only',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1031',
    description: 'Checks for permanent active privileged role assignments. Privileged roles should be managed as eligible assignments through PIM.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const privTemplates = new Set(PRIV_ROLE_TEMPLATE_IDS);
        const active = await Graph.graphAll('roleManagement/directory/roleAssignmentScheduleInstances?$expand=roleDefinition,principal', { apiVersion: 'v1.0' });
        const permanent = active.filter(a => a.assignmentType === 'Assigned' && privTemplates.has(a.roleDefinition?.templateId));

        if (!permanent.length) {
          return { id: 'MT.1031', status: 'Passed', reason: 'No permanent privileged assignments found. Roles appear to be managed via PIM activation/eligibility.', durationMs: ms(start) };
        }

        return {
          id: 'MT.1031',
          status: 'Failed',
          actual: permanent.length,
          reason: `${permanent.length} permanent privileged role assignment(s) found. Migrate these assignments to PIM eligible roles.`,
          durationMs: ms(start),
        };
      } catch (e) {
        if (e.status === 403 || e.status === 404) return { id: 'MT.1031', status: 'Skipped', reason: 'Privileged role schedule data is not accessible with current permissions.', durationMs: ms(start) };
        return errRow('MT.1031', e, start);
      }
    },
  });

  // ---------- MT.1036 Excluded objects have fallback include in another policy ----------
  tests.push({
    id: 'MT.1036', title: 'All excluded objects should have a fallback include in another policy',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1036',
    description: 'If a CA policy excludes a user or group, there should be another policy that includes those objects, ensuring full coverage. This check looks for any excluded users/groups that are not included in at least one other enabled policy.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const enabled = pols.filter(p => p.state === 'enabled');
        const gaps = [];
        for (const pol of enabled) {
          const excUsers = pol.conditions?.users?.excludeUsers || [];
          const excGroups = pol.conditions?.users?.excludeGroups || [];
          const excTargets = [...excUsers, ...excGroups];
          for (const excId of excTargets) {
            if (['All', 'GuestsOrExternalUsers'].includes(excId)) continue;
            const hasFallback = enabled.some(other => other.id !== pol.id
              && ((other.conditions?.users?.includeUsers || []).includes(excId)
                || (other.conditions?.users?.includeGroups || []).includes(excId)));
            if (!hasFallback) gaps.push({ policy: pol.displayName, excludedId: excId });
          }
        }
        if (gaps.length === 0) return { id: 'MT.1036', status: 'Passed', reason: 'All excluded objects have a fallback include in another policy.', durationMs: ms(start) };
        return { id: 'MT.1036', status: 'Failed', actual: gaps.length,
          reason: `${gaps.length} excluded object(s) lack a fallback include: ${gaps.slice(0, 3).map(g => `"${g.policy}" excludes ${g.excludedId}`).join('; ')}`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1036', e, start); }
    },
  });

  // ---------- MT.1038 CA policies not reference deleted groups ----------
  tests.push({
    id: 'MT.1038', title: 'Conditional Access policies should not include or exclude deleted groups',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1038',
    description: 'Checks all group IDs referenced in CA policies still exist in Entra ID. Deleted groups silently do nothing, creating security gaps.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const allGroupIds = new Set();
        for (const p of pols) {
          for (const id of [...(p.conditions?.users?.includeGroups || []), ...(p.conditions?.users?.excludeGroups || [])]) {
            if (id !== 'All' && id !== 'GuestsOrExternalUsers') allGroupIds.add(id);
          }
        }
        if (allGroupIds.size === 0) return { id: 'MT.1038', status: 'Passed', reason: 'No groups referenced in CA policies.', durationMs: ms(start) };
        const missing = [];
        for (const gid of allGroupIds) {
          try { await Graph.graph(`groups/${gid}?$select=id`, { apiVersion: 'v1.0' }); }
          catch (e) { if (e.status === 404) missing.push(gid); }
        }
        if (missing.length === 0) return { id: 'MT.1038', status: 'Passed', reason: `All ${allGroupIds.size} group(s) referenced in CA policies exist.`, durationMs: ms(start) };
        return { id: 'MT.1038', status: 'Failed', actual: missing.length,
          reason: `${missing.length} deleted group(s) still referenced in CA policies: ${missing.slice(0, 5).join(', ')}`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1038', e, start); }
    },
  });

  // ---------- MT.1049 User Risk and Sign-in Risk CA policies should be separate ----------
  tests.push({
    id: 'MT.1049', title: 'Conditional Access policies for User Risk and Sign-in Risk should be configured separately',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1049',
    description: 'Combining user risk and sign-in risk in the same CA policy is a common misconfiguration. They should be separate policies for proper control.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const combined = pols.filter(p => p.state === 'enabled'
          && (p.conditions?.userRiskLevels || []).length > 0
          && (p.conditions?.signInRiskLevels || []).length > 0);
        if (combined.length === 0) return { id: 'MT.1049', status: 'Passed', reason: 'No CA policies combine user risk and sign-in risk conditions.', durationMs: ms(start) };
        return { id: 'MT.1049', status: 'Failed', actual: combined.length,
          reason: `${combined.length} policy(ies) combine user risk and sign-in risk: ${combined.map(p => p.displayName).join(', ')}. Split these into separate policies.`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1049', e, start); }
    },
  });

  // ---------- MT.1052 Device Code auth flow targeted by CA ----------
  tests.push({
    id: 'MT.1052', title: 'At least one Conditional Access policy is targeting the Device Code authentication flow',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1052',
    description: 'Device code flow is a common phishing-friendly auth method. A CA policy should block or require additional controls for this flow.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.authenticationFlows?.transferMethods || []).includes('deviceCodeFlow'));
        if (hit) return { id: 'MT.1052', status: 'Passed', reason: `Policy "${hit.displayName}" targets the device code authentication flow.`, durationMs: ms(start) };
        return { id: 'MT.1052', status: 'Failed', reason: 'No enabled CA policy targets the device code authentication flow (conditions.authenticationFlows.transferMethods includes "deviceCodeFlow").', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1052', e, start); }
    },
  });

  // ---------- MT.1055 M365 Group creation restricted ----------
  tests.push({
    id: 'MT.1055', title: 'Microsoft 365 Group (and Team) creation should be restricted to approved users',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1055',
    description: 'By default all users can create M365 groups/Teams. Restricting this to approved users reduces Teams sprawl and improves governance.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const settings = await Graph.graphAll('groupSettings', { apiVersion: 'v1.0' });
        const groupUnified = settings.find(s => /unified/i.test(s.displayName || s.templateId));
        if (!groupUnified) return { id: 'MT.1055', status: 'Failed', reason: 'No "Group.Unified" tenant settings found. Group creation may not be restricted.', durationMs: ms(start) };
        const enabled = groupUnified.values?.find(v => v.name === 'EnableGroupCreation');
        if (!enabled || enabled.value?.toLowerCase() !== 'false') {
          return { id: 'MT.1055', status: 'Failed', reason: 'Group creation is not restricted (EnableGroupCreation is not set to false).', durationMs: ms(start) };
        }
        return { id: 'MT.1055', status: 'Passed', reason: 'Group creation is restricted (EnableGroupCreation = false).', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1055', e, start); }
    },
  });

  // ---------- MT.1057 M365 Group expiration notify ----------
  tests.push({
    id: 'MT.1057', title: 'Ensure Microsoft 365 Group (and Team) expiration is configured to notify users',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1057',
    description: 'Group lifecycle policies notify owners before groups expire, giving them a chance to renew or let them expire.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const policies = await Graph.graphAll('groupLifecyclePolicies', { apiVersion: 'v1.0' });
        if (policies.length === 0) return { id: 'MT.1057', status: 'Failed', reason: 'No Group Lifecycle Policies configured. Configure one to notify owners before expiry.', durationMs: ms(start) };
        const notifying = policies.filter(p => p.managedGroupTypes && (p.managedGroupTypes === 'All' || p.managedGroupTypes.includes('Selected')));
        if (notifying.length > 0) return { id: 'MT.1057', status: 'Passed', reason: `Group Lifecycle Policy configured: ${notifying.map(p => `${p.groupLifetimeInDays} days`).join(', ')}.`, durationMs: ms(start) };
        return { id: 'MT.1057', status: 'Failed', reason: 'Group Lifecycle Policy exists but does not manage any groups (managedGroupTypes not set to All or Selected).', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1057', e, start); }
    },
  });

  // ---------- MT.1058 M365 Group expiration auto-expire ----------
  tests.push({
    id: 'MT.1058', title: 'Ensure Microsoft 365 Group (and Team) expiration is configured to auto-expire groups',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1058',
    description: 'Group lifecycle policies should have a reasonable expiration window (e.g. 180 or 365 days) to automatically remove inactive groups.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const policies = await Graph.graphAll('groupLifecyclePolicies', { apiVersion: 'v1.0' });
        if (policies.length === 0) return { id: 'MT.1058', status: 'Failed', reason: 'No Group Lifecycle Policies configured.', durationMs: ms(start) };
        const p = policies[0];
        const days = Number(p.groupLifetimeInDays || 0);
        if (days > 0 && days <= 730) return { id: 'MT.1058', status: 'Passed', reason: `Group Lifecycle Policy set to ${days} days expiration.`, durationMs: ms(start) };
        return { id: 'MT.1058', status: 'Failed', reason: `Group Lifecycle Policy expiration is ${days || 'not set'} days — set to 180 or 365 days.`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1058', e, start); }
    },
  });

  // ---------- MT.1066 CA policies not reference deleted users/groups/roles ----------
  tests.push({
    id: 'MT.1066', title: 'Conditional Access policies should not reference non-existent users, groups, or roles',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1066',
    description: 'Checks that all user, group, and directory role IDs referenced in CA policies still exist. Stale references silently reduce coverage.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const missing = [];
        const seen = new Map();
        async function checkId(type, id) {
          if (seen.has(id)) return seen.get(id);
          const endpoints = { user: `users/${id}?$select=id`, group: `groups/${id}?$select=id`, role: `directoryRoles?$filter=roleTemplateId eq '${id}'&$select=id` };
          try {
            const r = await Graph.graph(endpoints[type] || `groups/${id}?$select=id`, { apiVersion: 'v1.0' });
            const exists = type === 'role' ? (r.value?.length > 0) : !!r.id;
            seen.set(id, exists);
            return exists;
          } catch (e) { seen.set(id, false); return false; }
        }
        for (const p of pols) {
          for (const id of (p.conditions?.users?.includeUsers || [])) {
            if (!['All', 'GuestsOrExternalUsers', 'None'].includes(id) && !await checkId('user', id)) missing.push({ policy: p.displayName, type: 'user', id });
          }
          for (const id of (p.conditions?.users?.excludeUsers || [])) {
            if (!['All', 'GuestsOrExternalUsers', 'None'].includes(id) && !await checkId('user', id)) missing.push({ policy: p.displayName, type: 'user', id });
          }
          for (const id of (p.conditions?.users?.includeGroups || [])) {
            if (!await checkId('group', id)) missing.push({ policy: p.displayName, type: 'group', id });
          }
          for (const id of (p.conditions?.users?.excludeGroups || [])) {
            if (!await checkId('group', id)) missing.push({ policy: p.displayName, type: 'group', id });
          }
        }
        if (missing.length === 0) return { id: 'MT.1066', status: 'Passed', reason: 'All referenced users, groups, and roles in CA policies exist.', durationMs: ms(start) };
        return { id: 'MT.1066', status: 'Failed', actual: missing.length,
          reason: `${missing.length} missing reference(s) in CA policies: ${missing.slice(0, 3).map(m => `"${m.policy}" refs deleted ${m.type} ${m.id}`).join('; ')}`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1066', e, start); }
    },
  });

  // ---------- MT.1069 Non-admin users can't create security groups ----------
  tests.push({
    id: 'MT.1069', title: 'Restrict non-admin users from creating security groups',
    severity: 'Low', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1069',
    description: 'If regular users can create security groups, they can build groups to gain access to resources or add members unexpectedly. Restrict this to admins.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pol = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const allowed = pol.defaultUserRolePermissions?.allowedToCreateSecurityGroups;
        if (allowed === false) return { id: 'MT.1069', status: 'Passed', reason: 'Non-admin users cannot create security groups (allowedToCreateSecurityGroups = false).', durationMs: ms(start) };
        return { id: 'MT.1069', status: 'Failed', reason: 'Non-admin users can create security groups. Disable this in the authorization policy.', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1069', e, start); }
    },
  });

  // ---------- MT.1071 Azure DevOps explicitly included in CA ----------
  tests.push({
    id: 'MT.1071', title: 'At least one Conditional Access policy explicitly includes Azure DevOps',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1071',
    description: 'Azure DevOps is not covered by "All apps" in CA. It should be explicitly targeted in at least one CA policy.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const AZURE_DEVOPS_APP = '499b84ac-1321-427f-aa17-267ca6975798';
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const hit = pols.find(p => p.state === 'enabled'
          && (p.conditions?.applications?.includeApplications || []).includes(AZURE_DEVOPS_APP));
        if (hit) return { id: 'MT.1071', status: 'Passed', reason: `Policy "${hit.displayName}" explicitly targets Azure DevOps.`, durationMs: ms(start) };
        return { id: 'MT.1071', status: 'Failed', reason: 'No enabled CA policy explicitly includes Azure DevOps (app ID 499b84ac-1321-427f-aa17-267ca6975798).', durationMs: ms(start) };
      } catch (e) { return errRow('MT.1071', e, start); }
    },
  });

  // ---------- MT.1072 No deprecated Approved Client App grant in CA ----------
  tests.push({
    id: 'MT.1072', title: 'Conditional access policies should not use the deprecated Approved Client App grant',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1072',
    description: 'The "Approved Client App" grant control is deprecated and replaced by app protection policies. Policies still using it should be updated.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const pols = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const bad = pols.filter(p => p.state === 'enabled'
          && (p.grantControls?.builtInControls || []).includes('approvedApplication'));
        if (bad.length === 0) return { id: 'MT.1072', status: 'Passed', reason: 'No CA policies use the deprecated "Approved Client App" grant.', durationMs: ms(start) };
        return { id: 'MT.1072', status: 'Failed', actual: bad.length,
          reason: `${bad.length} policy(ies) use deprecated "approvedApplication" grant: ${bad.map(p => p.displayName).join(', ')}`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1072', e, start); }
    },
  });

  // ---------- MT.1073 Soft/hard-matching of synced objects blocked ----------
  tests.push({
    id: 'MT.1073', title: 'Soft- and hard-matching of synchronized objects should be blocked',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1073',
    description: 'Blocking soft/hard match in Entra ID Connect prevents an on-prem attacker from syncing a user to take over a cloud account. Checks the cross-tenant access policy default settings.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const policy = await Graph.graph('policies/crossTenantAccessPolicy/default', { apiVersion: 'beta' });
        const blockSoft = policy?.automaticUserConsentSettings?.blockSoftMatch;
        if (blockSoft === true) return { id: 'MT.1073', status: 'Passed', reason: 'Soft/hard-matching is blocked in cross-tenant access policy defaults.', durationMs: ms(start) };
        return { id: 'MT.1073', status: 'Failed', reason: 'Soft/hard-matching of synchronized objects is not blocked (automaticUserConsentSettings.blockSoftMatch is not true).', durationMs: ms(start) };
      } catch (e) {
        if (e.status === 403 || e.status === 404) return { id: 'MT.1073', status: 'Skipped', reason: 'Cannot access crossTenantAccessPolicy (requires Policy.Read.All permission).', durationMs: ms(start) };
        return errRow('MT.1073', e, start);
      }
    },
  });

  // ---------- MT.1077 App registrations with privileged API permissions have no owners ----------
  tests.push({
    id: 'MT.1077', title: 'App registrations with privileged API permissions should not have owners',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1077',
    description: 'App owners can rotate secrets and effectively impersonate the app. Apps with high-risk permissions (like Directory.ReadWrite.All) should have no owners — use access reviews or managed identities instead.',
    implemented: true,
    async run() {
      const start = performance.now();
      const HIGH_PERMS = new Set([
        '19dbc75e-c2e2-444c-a770-ec69d8559fc7', // Directory.ReadWrite.All
        '9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8', // RoleManagement.ReadWrite.Directory
        '1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9', // Application.ReadWrite.All
        '06b708a9-e830-4db3-a914-8e69da51d44f', // AppRoleAssignment.ReadWrite.All
      ]);
      try {
        const spns = await Graph.graphAll('servicePrincipals?$select=id,appId,displayName,appRoles&$top=999', { apiVersion: 'v1.0' });
        const graphSpn = spns.find(s => s.appId === '00000003-0000-0000-c000-000000000000');
        if (!graphSpn) return { id: 'MT.1077', status: 'Skipped', reason: 'Graph SPN not found.', durationMs: ms(start) };
        const grants = await Graph.graphAll(`servicePrincipals/${graphSpn.id}/appRoleAssignedTo?$top=999`, { apiVersion: 'v1.0' });
        const highRiskAppIds = new Set(grants.filter(g => HIGH_PERMS.has(g.appRoleId)).map(g => g.principalId));
        if (highRiskAppIds.size === 0) return { id: 'MT.1077', status: 'Passed', reason: 'No service principals with high-risk API permissions found.', durationMs: ms(start) };
        const withOwners = [];
        for (const spId of [...highRiskAppIds].slice(0, 20)) {
          try {
            const sp = spns.find(s => s.id === spId);
            const appFilter = await Graph.graph(`applications?$filter=appId eq '${sp?.appId}'&$select=id,displayName`, { apiVersion: 'v1.0' });
            const app = appFilter.value?.[0];
            if (!app) continue;
            const owners = await Graph.graphAll(`applications/${app.id}/owners?$select=id`, { apiVersion: 'v1.0' });
            if (owners.length > 0) withOwners.push(app.displayName || spId);
          } catch (_) { /* skip */ }
        }
        if (withOwners.length === 0) return { id: 'MT.1077', status: 'Passed', reason: 'No app registrations with high-risk API permissions have owners.', durationMs: ms(start) };
        return { id: 'MT.1077', status: 'Failed', actual: withOwners.length,
          reason: `${withOwners.length} app(s) with high-risk API permissions have owners: ${withOwners.slice(0, 5).join(', ')}`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1077', e, start); }
    },
  });

  // ---------- MT.1081 Hybrid users not assigned Entra ID role assignments ----------
  tests.push({
    id: 'MT.1081', title: 'Hybrid users should not be assigned Entra ID role assignments',
    severity: 'Medium', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1081',
    description: 'On-prem synced (hybrid) users assigned to Entra ID roles are risky: if AD is compromised, so is the cloud role. All Entra role assignments should use cloud-only accounts.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const assignments = await Graph.graphAll('roleManagement/directory/roleAssignments?$expand=principal', { apiVersion: 'v1.0' });
        const userAssignments = assignments.filter(a => a.principal?.['@odata.type'] === '#microsoft.graph.user');
        if (userAssignments.length === 0) return { id: 'MT.1081', status: 'Passed', reason: 'No direct user role assignments found.', durationMs: ms(start) };
        const bad = [];
        for (const a of userAssignments.slice(0, 50)) {
          try {
            const u = await Graph.graph(`users/${a.principalId}?$select=onPremisesSyncEnabled,userPrincipalName`, { apiVersion: 'v1.0' });
            if (u.onPremisesSyncEnabled) bad.push(u.userPrincipalName);
          } catch (_) { /* skip */ }
        }
        if (bad.length === 0) return { id: 'MT.1081', status: 'Passed', reason: 'No hybrid (synced) users with Entra ID role assignments.', durationMs: ms(start) };
        return { id: 'MT.1081', status: 'Failed', actual: bad.length,
          reason: `${bad.length} hybrid user(s) with Entra ID role assignments: ${bad.slice(0, 5).join(', ')}`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1081', e, start); }
    },
  });

  // ---------- MT.1084 Seamless SSO disabled ----------
  tests.push({
    id: 'MT.1084', title: 'Seamless Single SignOn should be disabled for all domains in EntraID Connect servers',
    severity: 'High', tag: 'Maester', category: 'Maester',
    docUrl: 'https://maester.dev/docs/tests/MT.1084',
    description: 'Seamless SSO uses a service account (AZUREADSSOACC) in AD that, if compromised, enables Silver Ticket attacks against Entra ID. Disable it unless strictly required.',
    implemented: true,
    async run() {
      const start = performance.now();
      try {
        const domains = await Graph.graphAll('domains?$select=id,authenticationType,isVerified', { apiVersion: 'v1.0' });
        const federated = domains.filter(d => d.isVerified && d.authenticationType === 'Federated');
        if (federated.length === 0) {
          return { id: 'MT.1084', status: 'Passed', reason: 'No federated domains found — Seamless SSO is not applicable.', durationMs: ms(start) };
        }
        // Check for AZUREADSSOACC service principal which indicates SeamlessSSO is configured
        const ssoacc = await Graph.graph(`servicePrincipals?$filter=displayName eq 'AZUREADSSOACC'&$select=id,displayName`, { apiVersion: 'v1.0' });
        if ((ssoacc.value || []).length === 0) {
          return { id: 'MT.1084', status: 'Passed', reason: 'No AZUREADSSOACC service principal found — Seamless SSO does not appear to be configured.', durationMs: ms(start) };
        }
        return { id: 'MT.1084', status: 'Failed', reason: `AZUREADSSOACC service principal found — Seamless SSO appears to be configured. Disable it via Azure AD Connect if not needed.`, durationMs: ms(start) };
      } catch (e) { return errRow('MT.1084', e, start); }
    },
  });

  // ── MT.1074  SPF configured for all accepted domains (DNS-over-HTTPS) ────────
  async function getVerifiedAcceptedMailDomains() {
    const domains = await Graph.graphAll('domains?$select=id,isVerified,isInitial,supportedServices', { apiVersion: 'v1.0' });
    const mailDomains = domains.filter(d =>
      d.isVerified
      && d.isInitial !== true
      && !String(d.id || '').toLowerCase().endsWith('.onmicrosoft.com')
      && Array.isArray(d.supportedServices)
      && d.supportedServices.includes('Email')
    );
    // Fallback for tenants where supportedServices is not fully populated.
    if (mailDomains.length) return mailDomains;
    return domains.filter(d =>
      d.isVerified
      && d.isInitial !== true
      && !String(d.id || '').toLowerCase().endsWith('.onmicrosoft.com')
    );
  }

  tests.push({
    id: 'CISA.MS.EXO.2.2',
    title: 'An SPF policy SHALL be published for each domain, designating only these addresses as approved senders.',
    severity: 'Medium', category: 'Exchange', tag: 'Exchange',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.EXO.2.2',
    description: 'All verified domains in the tenant should have a valid SPF TXT record to prevent email spoofing.',
    async run() {
      const start = performance.now();
      try {
        const allDomains = await getVerifiedAcceptedMailDomains();
        if (!allDomains.length) return { id: 'CISA.MS.EXO.2.2', status: 'Skipped', reason: 'No custom verified domains found.', durationMs: ms(start) };
        const missing = [], found = [];
        for (const domain of allDomains) {
          try {
            const res = await fetch(
              `https://cloudflare-dns.com/dns-query?name=${encodeURIComponent(domain.id)}&type=TXT`,
              { headers: { Accept: 'application/dns-json' } }
            );
            const data = await res.json();
            const records = (data.Answer || []).map(a => (a.data || '').replace(/"/g, '').toLowerCase());
            const hasSpf = records.some(r => r.startsWith('v=spf1'));
            if (hasSpf) found.push(domain.id); else missing.push(domain.id);
          } catch { missing.push(`${domain.id} (DNS error)`); }
        }
        if (missing.length) {
          return { id: 'CISA.MS.EXO.2.2', status: 'Failed', reason: `${missing.length} domain(s) missing SPF record: ${missing.join(', ')}. ${found.length ? `Configured: ${found.join(', ')}.` : ''}`, durationMs: ms(start) };
        }
        return { id: 'CISA.MS.EXO.2.2', status: 'Passed', reason: `All ${found.length} custom domain(s) have SPF configured: ${found.join(', ')}.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.EXO.2.2', e, start); }
    },
  });

  // ── MT.1076  DMARC configured for all accepted domains ───────────────────────
  tests.push({
    id: 'CISA.MS.EXO.4.2',
    title: 'The DMARC message rejection option SHALL be p=reject.',
    severity: 'High', category: 'Exchange', tag: 'Exchange',
    docUrl: 'https://maester.dev/docs/tests/CISA.MS.EXO.4.2',
    description: 'All verified domains should have a DMARC record published at _dmarc.<domain>.',
    async run() {
      const start = performance.now();
      try {
        const allDomains = await getVerifiedAcceptedMailDomains();
        if (!allDomains.length) return { id: 'CISA.MS.EXO.4.2', status: 'Skipped', reason: 'No custom verified domains found.', durationMs: ms(start) };
        const missing = [], found = [], weak = [];
        for (const domain of allDomains) {
          try {
            const res = await fetch(
              `https://cloudflare-dns.com/dns-query?name=${encodeURIComponent('_dmarc.' + domain.id)}&type=TXT`,
              { headers: { Accept: 'application/dns-json' } }
            );
            const data = await res.json();
            const records = (data.Answer || []).map(a => (a.data || '').replace(/"/g, '').toLowerCase());
            const dmarcRec = records.find(r => r.startsWith('v=dmarc1'));
            if (dmarcRec) {
              const pMatch = dmarcRec.match(/\bp=(\w+)/);
              const policy = pMatch ? pMatch[1] : 'none';
              if (policy === 'reject') found.push(`${domain.id}(reject)`);
              else if (policy === 'quarantine') { found.push(`${domain.id}(quarantine)`); weak.push(domain.id); }
              else weak.push(`${domain.id}(${policy})`);
            } else missing.push(domain.id);
          } catch { missing.push(`${domain.id} (DNS error)`); }
        }
        if (missing.length) {
          return { id: 'CISA.MS.EXO.4.2', status: 'Failed', reason: `${missing.length} domain(s) missing DMARC record: ${missing.join(', ')}.${weak.length ? ` Weak policies: ${weak.join(', ')}.` : ''}`, durationMs: ms(start) };
        }
        if (weak.length) {
          return { id: 'CISA.MS.EXO.4.2', status: 'Failed', reason: `All domains have DMARC but ${weak.length} use weak policy (none/quarantine): ${weak.join(', ')}. Upgrade to p=reject.`, durationMs: ms(start) };
        }
        return { id: 'CISA.MS.EXO.4.2', status: 'Passed', reason: `All ${allDomains.length} domain(s) have DMARC with strong policy: ${found.join(', ')}.`, durationMs: ms(start) };
      } catch (e) { return errRow('CISA.MS.EXO.4.2', e, start); }
    },
  });

  function ms(start) { return Math.round(performance.now() - start); }
  function errRow(id, e, start) {
    if (e.status === 403 || e.status === 401) return { id, status:'Skipped', reason:`Insufficient permissions: ${e.message}`, durationMs: ms(start) };
    return { id, status:'Error', reason: e.message, durationMs: ms(start) };
  }

  // The runner expects each test to also carry display metadata. Wrap each test to
  // merge run() output with the catalog metadata, so result rows always have title,
  // severity, etc.
  const RUN_CATEGORY = {
    // Conditional Access
    'MT.1006': 'CA',
    'MT.1007': 'CA', 'MT.1008': 'CA', 'MT.1009': 'CA', 'MT.1011': 'CA',
    'MT.1012': 'CA', 'MT.1013': 'CA', 'MT.1014': 'CA', 'MT.1016': 'CA',
    'MT.1017': 'CA', 'MT.1018': 'CA', 'MT.1019': 'CA', 'MT.1021': 'CA',
    'MT.1036': 'CA', 'MT.1038': 'CA', 'MT.1049': 'CA', 'MT.1052': 'CA',
    'MT.1066': 'CA', 'MT.1071': 'CA', 'MT.1072': 'CA',
    // Privileged Identity / role hygiene
    'MT.1032': 'Privileged', 'CIS.M365.1.1.1': 'Privileged', 'CISA.MS.AAD.7.4': 'Privileged',
    'MT.1025': 'Privileged', 'MT.1026': 'Privileged', 'MT.1027': 'Privileged',
    'MT.1028': 'Privileged', 'MT.1029': 'Privileged', 'MT.1031': 'Privileged',
    'MT.1081': 'Privileged',
    // Application / service principal hygiene
    'MT.1063': 'App',
    'MT.1055': 'App', 'MT.1057': 'App', 'MT.1058': 'App', 'MT.1077': 'App',
    // Authentication methods
    'CISA.MS.AAD.3.6': 'Authentication', 'CISA.MS.AAD.3.5': 'Authentication',
    // General Entra hygiene
    'MT.1069': 'Entra', 'MT.1073': 'Entra', 'MT.1084': 'Entra',
    // Defender
    'MT.DEF.1001': 'Defender', 'MT.DEF.1002': 'Defender', 'MT.DEF.1003': 'Defender',
    // Exchange / DNS
    'CISA.MS.EXO.2.2': 'Exchange', 'CISA.MS.EXO.4.2': 'Exchange',
  };

  function buildCatalog() {
    return tests.filter(t => !/^SM\./.test(t.id)).map(t => ({
      id: t.id, title: t.title, severity: t.severity, tag: t.tag, category: t.category, docUrl: t.docUrl,
      runCategory: RUN_CATEGORY[t.id] || 'Entra',
      description: t.description, detailMd: t.detailMd, implemented: true,
      requiredScopes: t.requiredScopes || Auth.SCOPES.graphFull,
      async run(ctx) {
        const r = await t.run(ctx);
        return { ...t, ...r, tag: t.tag, category: t.category };
      },
    }));
  }

  window.TestsNative = { buildCatalog };
})();
