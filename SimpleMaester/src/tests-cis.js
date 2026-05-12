// CIS Microsoft 365 Foundations Benchmark v6.0.1 tests, ported from the
// upstream Maester PowerShell module (powershell/public/cis/Test-MtCis*.ps1).
//
// Tests run purely via Microsoft Graph and PKCE delegated auth in the browser.
//
// Tests that require Exchange admin REST (outlook.office365.com/adminapi/...) are
// CORS-blocked when called from a browser SPA - Microsoft does not add CORS headers
// to the Exchange admin endpoint. Those tests return Skipped with a note to run the
// Maester PowerShell equivalent instead.
//
// Tests that require Microsoft Teams admin REST are also CORS-blocked / not exposed
// for browser PKCE and return Skipped for the same reason.
(() => {
  const tests = [];

  function ms(start) { return Math.round(performance.now() - start); }
  function row(id, status, reason, extra) { return Object.assign({ id, status, reason }, extra || {}); }
  function errRow(id, e, start) {
    if (e.status === 401 || e.status === 403) {
      return row(id, 'Skipped', `Insufficient permissions: ${e.message}`, { durationMs: ms(start) });
    }
    return row(id, 'Error', e.message || String(e), { durationMs: ms(start) });
  }

  function pickAny(obj, keys) {
    if (!obj) return undefined;
    for (const k of keys) {
      if (obj[k] !== undefined) return obj[k];
    }
    return undefined;
  }

  function unwrapGraphValue(payload) {
    if (!payload || Array.isArray(payload)) return payload;
    if (payload.value && !Array.isArray(payload.value)) return payload.value;
    return payload;
  }

  async function getFirstAvailablePolicy(candidates) {
    let lastErr = null;
    for (const c of candidates) {
      const candidate = typeof c === 'string'
        ? { path: c, apiVersion: 'beta', tokenScopes: Auth.SCOPES.graphFull }
        : {
            path: c.path,
            apiVersion: c.apiVersion || 'beta',
            tokenScopes: c.tokenScopes || Auth.SCOPES.graphFull,
          };
      try {
        const data = await Graph.graph(candidate.path, {
          apiVersion: candidate.apiVersion,
          tokenScopes: candidate.tokenScopes,
        });
        return unwrapGraphValue(data);
      } catch (e) {
        const msg = String(e.message || '').toLowerCase();
        const notFoundSegment = e.status === 400 && (
          msg.includes('resource not found for the segment') ||
          msg.includes('resource not found') ||
          msg.includes('could not find a property named')
        );
        if (e.status === 404 || e.status === 405 || e.status === 501 || notFoundSegment) {
          lastErr = e;
          continue;
        }
        throw e;
      }
    }
    const err = new Error(`No supported Teams policy endpoint found (${candidates.join(', ')})`);
    err.status = (lastErr && lastErr.status) || 404;
    throw err;
  }

  function teamsSkip(id, e, start) {
    if (e.status === 401 || e.status === 403) {
      return row(id, 'Skipped', `Insufficient permissions for Teams policy API: ${e.message}`, { durationMs: ms(start) });
    }
    if (e.status === 400 || e.status === 404 || e.status === 405 || e.status === 501) {
      return row(id, 'Skipped', `Teams policy API not available in this tenant/app: ${e.message}`, { durationMs: ms(start) });
    }
    return row(id, 'Error', e.message || String(e), { durationMs: ms(start) });
  }

  // Exchange admin REST (outlook.office365.com/adminapi/...) is CORS-blocked when
  // called from a browser SPA. Microsoft does not add CORS headers to that endpoint.
  // Tests in this category that need Exchange return Skipped instead of failing.
  function exoStub(id, cmdlets) {
    return row(id, 'Skipped',
      `Not available from a browser SPA: requires Exchange admin REST (CORS-blocked by outlook.office365.com). ` +
      `Equivalent PowerShell: ${cmdlets}. Run this check in the Maester PowerShell module instead.`);
  }

  function teamsStub(id) {
    return row(id, 'Skipped',
      'Not available from a browser SPA: requires Microsoft Teams admin REST which is not exposed for delegated PKCE auth. ' +
      'Run the equivalent Test-MtCis* in the Maester PowerShell module instead.');
  }

  // ---------- 1.1.1 CloudAdmin ----------
  tests.push({
    id: 'CIS.M365.1.1.1', title: '(L1) Ensure administrative accounts are cloud-only',
    severity: 'High',
    description: 'Synced on-prem accounts in any privileged Entra role widen the blast radius to your local AD. CIS recommends cloud-only admin accounts.',
    async run() {
      const start = performance.now();
      try {
        const roles = await Graph.graphAll('directoryRoles?$expand=members', { apiVersion: 'v1.0' });
        const synced = [];
        for (const r of roles) {
          for (const m of (r.members || [])) {
            if (m['@odata.type'] !== '#microsoft.graph.user') continue;
            try {
              const u = await Graph.graph(`users/${m.id}?$select=onPremisesSyncEnabled,userPrincipalName`, { apiVersion: 'v1.0' });
              if (u.onPremisesSyncEnabled) synced.push(`${u.userPrincipalName} (${r.displayName})`);
            } catch (_) { /* skip */ }
          }
        }
        if (synced.length === 0) return row('CIS.M365.1.1.1', 'Passed', 'All administrative accounts are cloud-only.', { actual: 0, durationMs: ms(start) });
        return row('CIS.M365.1.1.1', 'Failed', `${synced.length} synced admin account(s): ${synced.slice(0, 5).join('; ')}`, { actual: synced.length, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.1.1.1', e, start); }
    },
  });

  // ---------- 1.1.3 GlobalAdminCount ----------
  tests.push({
    id: 'CIS.M365.1.1.3', title: '(L1) Ensure between two and four Global Administrators are designated',
    severity: 'Medium',
    description: 'Fewer than 2 means no fallback if one account breaks; more than 4 inflates the blast radius of compromise.',
    async run() {
      const start = performance.now();
      try {
        const roles = await Graph.graphAll('directoryRoles', { apiVersion: 'v1.0' });
        const ga = roles.find(r => r.roleTemplateId === '62e90394-69f5-4237-9190-012177145e10');
        if (!ga) return row('CIS.M365.1.1.3', 'Skipped', 'Global Administrator role not activated in this tenant.', { durationMs: ms(start) });
        const members = await Graph.graphAll(`directoryRoles/${ga.id}/members`, { apiVersion: 'v1.0' });
        const userCount = members.filter(m => m['@odata.type'] === '#microsoft.graph.user').length;
        const ok = userCount >= 2 && userCount <= 4;
        return row('CIS.M365.1.1.3', ok ? 'Passed' : 'Failed', `Found ${userCount} Global Administrator user(s).`, { actual: userCount, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.1.1.3', e, start); }
    },
  });

  // ---------- 1.2.1 365PublicGroup ----------
  tests.push({
    id: 'CIS.M365.1.2.1', title: '(L2) Ensure that only organizationally managed/approved public groups exist',
    severity: 'Medium',
    description: 'Public Microsoft 365 groups are visible to everyone and content is accessible to all members of the org. CIS recommends none unless reviewed.',
    async run() {
      const start = performance.now();
      try {
        // 'visibility' is not filterable on /groups, so pull M365 groups (which are
        // the only ones that can be Public) and filter client-side.
        const all = await Graph.graphAll(`groups?$filter=groupTypes/any(c:c eq 'Unified')&$select=id,displayName,visibility`, { apiVersion: 'v1.0' });
        const pub = all.filter(g => (g.visibility || '').toLowerCase() === 'public');
        if (pub.length === 0) return row('CIS.M365.1.2.1', 'Passed', `No public M365 groups found (checked ${all.length}).`, { actual: 0, durationMs: ms(start) });
        return row('CIS.M365.1.2.1', 'Failed', `${pub.length} public group(s): ${pub.slice(0, 5).map(g => g.displayName).join('; ')}`, { actual: pub.length, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.1.2.1', e, start); }
    },
  });

  // ---------- 1.2.2 SharedMailboxSignIn ----------
  tests.push({
    id: 'CIS.M365.1.2.2', title: '(L1) Ensure sign-in to shared mailboxes is blocked',
    severity: 'High',
    description: 'Shared mailboxes have a backing user account that should be disabled to prevent direct sign-in.',
    async run() { return exoStub('CIS.M365.1.2.2', 'Get-Mailbox -RecipientTypeDetails SharedMailbox'); },
  });

  // ---------- 1.3.1 PasswordExpiry ----------
  tests.push({
    id: 'CIS.M365.1.3.1', title: '(L1) Ensure the "Password expiration policy" is set to "Set passwords to never expire"',
    severity: 'Medium',
    description: 'Periodic password expiry encourages reuse of weak patterns. NIST and Microsoft recommend disabling expiry on all managed domains.',
    async run() {
      const start = performance.now();
      try {
        const doms = await Graph.graphAll('domains', { apiVersion: 'v1.0' });
        const bad = doms.filter(d => d.authenticationType === 'Managed' && d.passwordValidityPeriodInDays !== 2147483647);
        if (bad.length === 0) return row('CIS.M365.1.3.1', 'Passed', 'All managed domains have passwords set to never expire.', { actual: 0, durationMs: ms(start) });
        return row('CIS.M365.1.3.1', 'Failed', `${bad.length} managed domain(s) still expire passwords: ${bad.map(d => d.id).join('; ')}`, { actual: bad.length, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.1.3.1', e, start); }
    },
  });

  // ---------- 1.3.3 CalendarSharing ----------
  tests.push({
    id: 'CIS.M365.1.3.3', title: '(L2) Ensure "External sharing" of calendars is not available',
    severity: 'Medium',
    description: 'External calendar sharing exposes free/busy and meeting subjects to anyone with the URL. CIS recommends disabling.',
    async run() { return exoStub('CIS.M365.1.3.3', 'Get-SharingPolicy'); },
  });

  // ---------- 1.3.5 FormsPhishingProtectionEnabled ----------
  tests.push({
    id: 'CIS.M365.1.3.5', title: '(L2) Ensure internal phishing protection for Forms is enabled',
    severity: 'Medium',
    description: 'Microsoft Forms can detect common phishing question patterns and warn users. Enable in admin/forms/settings.',
    requiredScopes: Auth.SCOPES.graphForms,
    async run() {
      const start = performance.now();
      try {
        const forms = await Graph.graph('admin/forms', { apiVersion: 'beta', tokenScopes: Auth.SCOPES.graphForms });
        const settings = forms?.settings || forms?.value?.settings || forms?.value?.[0]?.settings || {};
        const ok = settings?.isInOrgFormsPhishingScanEnabled === true;
        return row('CIS.M365.1.3.5', ok ? 'Passed' : 'Failed', ok ? 'Forms phishing scan is enabled.' : 'Forms phishing scan is disabled.', { actual: !!settings?.isInOrgFormsPhishingScanEnabled, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.1.3.5', e, start); }
    },
  });

  // ---------- 1.3.6 CustomerLockBox ----------
  tests.push({
    id: 'CIS.M365.1.3.6', title: '(L2) Ensure the Customer Lockbox feature is enabled',
    severity: 'Medium',
    description: 'Customer Lockbox requires explicit approval before Microsoft engineers can access your data. Available with E5/Compliance add-on.',
    async run() { return exoStub('CIS.M365.1.3.6', 'Get-OrganizationConfig | Select CustomerLockBoxEnabled'); },
  });

  // ---------- 2.1.1 SafeLink ----------
  tests.push({
    id: 'CIS.M365.2.1.1', title: '(L2) Ensure Safe Links for Office Applications is Enabled',
    severity: 'High',
    description: 'Safe Links rewrites URLs in mail and Office docs and time-of-click checks them against MDO threat intel.',
    async run() { return exoStub('CIS.M365.2.1.1', 'Get-SafeLinksRule, Get-SafeLinksPolicy'); },
  });

  // ---------- 2.1.2 SafeAttachment ----------
  tests.push({
    id: 'CIS.M365.2.1.2', title: '(L2) Ensure the Common Attachment Types Filter is enabled (Safe Attachments)',
    severity: 'High',
    description: 'Safe Attachments policy must have Action=Block and QuarantineTag=AdminOnlyAccessPolicy.',
    async run() { return exoStub('CIS.M365.2.1.2', 'Get-SafeAttachmentPolicy'); },
  });

  // ---------- 2.1.3 SafeAttachmentsAtpPolicy ----------
  tests.push({
    id: 'CIS.M365.2.1.3', title: '(L2) Ensure Safe Attachments for SharePoint, OneDrive and Microsoft Teams is enabled',
    severity: 'High',
    description: 'AtpPolicyForO365: EnableATPForSPOTeamsODB=True, EnableSafeDocs=True, AllowSafeDocsOpen=False.',
    async run() { return exoStub('CIS.M365.2.1.3', 'Get-AtpPolicyForO365'); },
  });

  // ---------- 2.1.4 AttachmentFilter (L1) ----------
  tests.push({
    id: 'CIS.M365.2.1.4', title: '(L1) Ensure the Common Attachment Types Filter is enabled',
    severity: 'High',
    description: 'Default malware filter policy must have EnableFileFilter=True.',
    async run() { return exoStub('CIS.M365.2.1.4', 'Get-MalwareFilterPolicy'); },
  });

  // ---------- 2.1.5 InternalMalwareNotification ----------
  tests.push({
    id: 'CIS.M365.2.1.5', title: '(L1) Ensure notifications for internal users sending malware is Enabled',
    severity: 'Medium',
    description: 'Default malware policy must have EnableInternalSenderAdminNotifications=True and InternalSenderAdminAddress set.',
    async run() { return exoStub('CIS.M365.2.1.5', 'Get-MalwareFilterPolicy'); },
  });

  // ---------- 2.1.6 SafeAntiPhishingPolicy ----------
  tests.push({
    id: 'CIS.M365.2.1.6', title: '(L1) Ensure Exchange Online anti-phishing policy is configured',
    severity: 'High',
    description: 'Default anti-phish policy: Enabled, EnableMailboxIntelligence(+Protection), EnableSpoofIntelligence, PhishThresholdLevel >= 2.',
    async run() { return exoStub('CIS.M365.2.1.6', 'Get-AntiPhishPolicy'); },
  });

  // ---------- 2.1.7 ConnectionFilterSafeList ----------
  tests.push({
    id: 'CIS.M365.2.1.7', title: '(L1) Ensure that an anti-phishing policy has been created (SafeList not enabled)',
    severity: 'Medium',
    description: 'EnableSafeList on the default Hosted Connection Filter Policy bypasses spam scanning. Should be False.',
    async run() { return exoStub('CIS.M365.2.1.7', 'Get-HostedConnectionFilterPolicy'); },
  });

  // ---------- 2.1.8 HostedConnectionFilterPolicy (IPAllowList empty) ----------
  tests.push({
    id: 'CIS.M365.2.1.8', title: '(L1) Ensure IP allow list is not used in Hosted Connection Filter Policy',
    severity: 'Medium',
    description: 'IPAllowList bypasses spam filtering for those IPs. CIS recommends an empty list.',
    async run() { return exoStub('CIS.M365.2.1.8', 'Get-HostedConnectionFilterPolicy'); },
  });

  // ---------- 2.1.9 OutboundSpamFilterPolicy ----------
  tests.push({
    id: 'CIS.M365.2.1.9', title: '(L1) Ensure Exchange Online Spam policies are set to notify administrators',
    severity: 'Medium',
    description: 'Default outbound spam policy: BccSuspiciousOutboundMail=True and NotifyOutboundSpam=True.',
    async run() { return exoStub('CIS.M365.2.1.9', 'Get-HostedOutboundSpamFilterPolicy'); },
  });

  // ---------- 2.1.10 AttachmentFilterComprehensive (L2) ----------
  tests.push({
    id: 'CIS.M365.2.1.10', title: '(L2) Ensure comprehensive attachment filtering is configured',
    severity: 'Medium',
    description: 'L2 expects 200+ blocked file extensions in the default malware policy FileTypes list.',
    async run() { return exoStub('CIS.M365.2.1.10', 'Get-MalwareFilterPolicy'); },
  });

  // ---------- 2.1.11 EnsureGuestAccessRestricted ----------
  tests.push({
    id: 'CIS.M365.2.1.11', title: '(L1) Ensure guest user access is restricted',
    severity: 'High',
    description: 'guestUserRoleId in authorizationPolicy must be either "Restricted Guest User" (10dae51f...) or "Guest User" (2af84b1e...).',
    async run() {
      const start = performance.now();
      try {
        const settings = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const id = settings?.guestUserRoleId;
        const ok = id === '10dae51f-b6af-4016-8d66-8c2a99b929b3' || id === '2af84b1e-32c8-42b7-82bc-daa82404023b';
        return row('CIS.M365.2.1.11', ok ? 'Passed' : 'Failed', `guestUserRoleId = ${id}`, { actual: id, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.2.1.11', e, start); }
    },
  });

  // ---------- 2.1.12 ZAP (Teams) ----------
  tests.push({
    id: 'CIS.M365.2.1.12', title: '(L1) Ensure Zero-hour auto purge for Teams is on',
    severity: 'Medium',
    description: 'ZapEnabled on TeamsProtectionPolicy must be True.',
    async run() {
      const start = performance.now();
      try {
        const p = await getFirstAvailablePolicy([
          'admin/teams/protectionPolicy',
          'admin/teams/protectionPolicies/global',
          'admin/teams/settings',
        ]);
        const zap = pickAny(p, ['zapEnabled', 'ZapEnabled']);
        if (typeof zap !== 'boolean') {
          return row('CIS.M365.2.1.12', 'Skipped', 'Teams policy endpoint reachable, but ZapEnabled is not exposed in Graph for this tenant.', { durationMs: ms(start) });
        }
        const ok = zap === true;
        return row('CIS.M365.2.1.12', ok ? 'Passed' : 'Failed', ok ? 'Zero-hour auto purge (ZAP) is enabled.' : 'Zero-hour auto purge (ZAP) is disabled.', { actual: zap, durationMs: ms(start) });
      } catch (e) { return teamsSkip('CIS.M365.2.1.12', e, start); }
    },
  });

  // ---------- 2.1.14 Dkim ----------
  tests.push({
    id: 'CIS.M365.2.1.14', title: '(L1) Ensure that DKIM is enabled for all Exchange Online Domains',
    severity: 'High',
    description: 'Every accepted (sending) domain should have DKIM enabled in Exchange. There is no Graph API for DKIM signing config.',
    async run() { return exoStub('CIS.M365.2.1.14', 'Get-DkimSigningConfig, Get-AcceptedDomain'); },
  });

  // ---------- 3.1.1 AuditLogSearch ----------
  tests.push({
    id: 'CIS.M365.3.1.1', title: '(L1) Ensure Microsoft 365 audit log search is enabled',
    severity: 'High',
    description: 'UnifiedAuditLogIngestionEnabled in admin audit log config must be True. There is no public Graph API for this setting.',
    async run() { return exoStub('CIS.M365.3.1.1', 'Get-AdminAuditLogConfig | Select UnifiedAuditLogIngestionEnabled'); },
  });

  // ---------- 4.1 DevicesWithoutCompliancePolicyMarked ----------
  tests.push({
    id: 'CIS.M365.4.1', title: '(L1) Ensure devices without a compliance policy are marked "not compliant"',
    severity: 'Medium',
    description: 'deviceManagement/settings.secureByDefault must be True.',
    async run() {
      const start = performance.now();
      try {
        const settings = await Graph.graph('deviceManagement/settings', { apiVersion: 'v1.0' });
        const ok = settings?.secureByDefault === true;
        return row('CIS.M365.4.1', ok ? 'Passed' : 'Failed', `secureByDefault = ${settings?.secureByDefault}`, { actual: !!settings?.secureByDefault, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.4.1', e, start); }
    },
  });

  // ---------- 5.1.2.3 CreateTenantDisallowed ----------
  tests.push({
    id: 'CIS.M365.5.1.2.3', title: '(L1) Ensure non-admin users cannot create tenants',
    severity: 'Medium',
    description: 'authorizationPolicy.defaultUserRolePermissions.allowedToCreateTenants must be false.',
    async run() {
      const start = performance.now();
      try {
        const p = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const v = p?.defaultUserRolePermissions?.allowedToCreateTenants;
        return row('CIS.M365.5.1.2.3', v === false ? 'Passed' : 'Failed', `allowedToCreateTenants = ${v}`, { actual: v, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.5.1.2.3', e, start); }
    },
  });

  // ---------- 5.1.2.5 ThirdPartyApplicationsDisallowed ----------
  tests.push({
    id: 'CIS.M365.5.1.2.5', title: '(L1) Ensure third-party application registration is restricted',
    severity: 'High',
    description: 'authorizationPolicy.defaultUserRolePermissions.allowedToCreateApps must be false.',
    async run() {
      const start = performance.now();
      try {
        const p = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const v = p?.defaultUserRolePermissions?.allowedToCreateApps;
        return row('CIS.M365.5.1.2.5', v === false ? 'Passed' : 'Failed', `allowedToCreateApps = ${v}`, { actual: v, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.5.1.2.5', e, start); }
    },
  });

  // ---------- 5.1.5.2 AdminConsentWorkflowEnabled ----------
  tests.push({
    id: 'CIS.M365.5.1.5.2', title: '(L1) Ensure the admin consent workflow is enabled',
    severity: 'Medium',
    description: 'adminConsentRequestPolicy.isEnabled must be true so users can request admin consent for apps.',
    async run() {
      const start = performance.now();
      try {
        const p = await Graph.graph('policies/adminConsentRequestPolicy', { apiVersion: 'v1.0' });
        const v = p?.isEnabled === true;
        return row('CIS.M365.5.1.5.2', v ? 'Passed' : 'Failed', `isEnabled = ${p?.isEnabled}`, { actual: !!p?.isEnabled, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.5.1.5.2', e, start); }
    },
  });

  // ---------- 5.1.6.1 EnsureUserConsentToAppsDisallowed ----------
  tests.push({
    id: 'CIS.M365.5.1.6.1', title: '(L2) Ensure user consent to apps is disallowed',
    severity: 'High',
    description: 'permissionGrantPoliciesAssigned must NOT contain ManagePermissionGrantsForSelf.microsoft-user-default-low or -legacy.',
    async run() {
      const start = performance.now();
      try {
        const p = await Graph.graph('policies/authorizationPolicy', { apiVersion: 'v1.0' });
        const list = p?.defaultUserRolePermissions?.permissionGrantPoliciesAssigned || [];
        const bad = list.filter(x => x === 'ManagePermissionGrantsForSelf.microsoft-user-default-low' || x === 'ManagePermissionGrantsForSelf.microsoft-user-default-legacy');
        const ok = bad.length === 0;
        return row('CIS.M365.5.1.6.1', ok ? 'Passed' : 'Failed', ok ? 'Users cannot consent to apps.' : `Bad assignments: ${bad.join(', ')}`, { actual: list, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.5.1.6.1', e, start); }
    },
  });

  // ---------- 5.1.8.1 ThirdPartyAndCustomApps (Teams) ----------
  tests.push({
    id: 'CIS.M365.5.1.8.1', title: '(L2) Ensure third-party and custom Teams apps are blocked',
    severity: 'Medium',
    description: 'Teams global app permission policy should block third-party and custom apps by default.',
    async run() {
      const start = performance.now();
      try {
        const p = await getFirstAvailablePolicy([
          { path: 'teamwork/teamsAppSettings', apiVersion: 'v1.0', tokenScopes: Auth.SCOPES.graphFull },
          { path: 'teamwork/teamsAppSettings', apiVersion: 'beta', tokenScopes: Auth.SCOPES.graphFull },
        ]);

        const customAppSettings = p?.customAppSettings || {};
        const globalType = String(
          pickAny(customAppSettings, ['globalCatalogAppsType', 'GlobalCatalogAppsType']) ||
          pickAny(p, ['globalCatalogAppsType', 'GlobalCatalogAppsType']) ||
          ''
        ).toLowerCase();
        const privateType = String(
          pickAny(customAppSettings, ['privateCatalogAppsType', 'PrivateCatalogAppsType']) ||
          pickAny(p, ['privateCatalogAppsType', 'PrivateCatalogAppsType']) ||
          ''
        ).toLowerCase();
        const customDisabled = pickAny(customAppSettings, ['isSideloadingEnabled', 'isSideLoadingEnabled']) ??
          pickAny(p, ['isSideloadingEnabled', 'isSideLoadingEnabled', 'allowUserRequestsForAppAccess']);

        // Best-effort mapping to CIS intent:
        // - third-party should not be broadly allowed
        // - custom apps should not be broadly allowed
        const thirdPartyBlocked = globalType.includes('allowedapplist') || globalType.includes('blockall') || globalType.includes('blockedapplist');
        const customBlocked = privateType.includes('allowedapplist') || privateType.includes('blockall') || privateType.includes('blockedapplist') || customDisabled === false;

        if (!globalType && !privateType && typeof customDisabled !== 'boolean') {
          return row('CIS.M365.5.1.8.1', 'Skipped', 'Teams app policy endpoint reachable, but required policy fields are not exposed for browser Graph.', { durationMs: ms(start) });
        }

        const ok = thirdPartyBlocked && customBlocked;
        return row(
          'CIS.M365.5.1.8.1',
          ok ? 'Passed' : 'Failed',
          ok ? 'Third-party and custom Teams apps are restricted.' : 'Third-party and/or custom Teams apps are not restricted enough.',
          { actual: { globalCatalogAppsType: globalType || null, privateCatalogAppsType: privateType || null, customToggle: customDisabled ?? null }, durationMs: ms(start) }
        );
      } catch (e) { return teamsSkip('CIS.M365.5.1.8.1', e, start); }
    },
  });

  // ---------- 5.1.8.2 EnsureGuestUserDynamicGroup ----------
  tests.push({
    id: 'CIS.M365.5.1.8.2', title: '(L2) Ensure a dynamic group exists targeting guest users',
    severity: 'Low',
    description: 'There should be at least one dynamic group with a membership rule like (user.userType -eq "Guest").',
    async run() {
      const start = performance.now();
      try {
        const groups = await Graph.graphAll(`groups?$filter=groupTypes/any(c:c eq 'DynamicMembership')&$select=id,displayName,membershipRule`, { apiVersion: 'v1.0', tokenScopes: Auth.SCOPES.graphFull });
        const guestGroups = groups.filter(g => /user\.userType\s*-eq\s*"?Guest"?/i.test(g.membershipRule || ''));
        const ok = guestGroups.length >= 1;
        return row('CIS.M365.5.1.8.2', ok ? 'Passed' : 'Failed', ok ? `Found ${guestGroups.length} guest dynamic group(s).` : 'No dynamic group with a Guest membership rule found.', { actual: guestGroups.length, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.5.1.8.2', e, start); }
    },
  });

  // ---------- 5.x UserOwnedAppsRestricted ----------
  tests.push({
    id: 'CIS.M365.5.5.1', title: '(L1) Restrict user-owned add-ins and trials',
    severity: 'Medium',
    description: 'admin/appsAndServices/settings: isOfficeStoreEnabled=False AND isAppAndServicesTrialEnabled=False.',
    requiredScopes: Auth.SCOPES.graphAppsAndServices,
    async run() {
      const start = performance.now();
      try {
        const appsAndServices = await Graph.graph('admin/appsAndServices', { apiVersion: 'beta', tokenScopes: Auth.SCOPES.graphAppsAndServices });
        const s = appsAndServices?.settings || appsAndServices?.value?.settings || appsAndServices?.value?.[0]?.settings || {};
        const failed = [];
        if (s?.isOfficeStoreEnabled !== false) failed.push('isOfficeStoreEnabled');
        if (s?.isAppAndServicesTrialEnabled !== false) failed.push('isAppAndServicesTrialEnabled');
        const ok = failed.length === 0;
        return row('CIS.M365.5.5.1', ok ? 'Passed' : 'Failed', ok ? 'Office Store and trials disabled.' : `Failed: ${failed.join(', ')}`, { actual: { isOfficeStoreEnabled: s?.isOfficeStoreEnabled, isAppAndServicesTrialEnabled: s?.isAppAndServicesTrialEnabled }, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.5.5.1', e, start); }
    },
  });

  // ---------- 5.2.2.1 WeakAuthenticationMethodsDisabled ----------
  tests.push({
    id: 'CIS.M365.5.2.2.1', title: '(L1) Ensure weak authentication methods (SMS, Voice, Email OTP) are disabled',
    severity: 'High',
    description: 'authenticationMethodsPolicy: Sms, Voice and Email configurations must all be state=disabled.',
    async run() {
      const start = performance.now();
      try {
        const p = await Graph.graph('policies/authenticationMethodsPolicy', { apiVersion: 'v1.0' });
        const cfgs = p?.authenticationMethodConfigurations || [];
        const get = id => cfgs.find(c => c.id === id);
        const failed = [];
        ['Sms', 'Voice', 'Email'].forEach(id => {
          const c = get(id);
          if (c && c.state !== 'disabled') failed.push(id);
        });
        const ok = failed.length === 0;
        return row('CIS.M365.5.2.2.1', ok ? 'Passed' : 'Failed', ok ? 'SMS, Voice and Email OTP disabled.' : `Enabled methods: ${failed.join(', ')}`, { actual: failed, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.5.2.2.1', e, start); }
    },
  });

  // ---------- 6.5.3 ExoAdditionalStorageProvider ----------
  tests.push({
    id: 'CIS.M365.6.5.3', title: '(L2) Ensure additional storage providers are restricted in OWA',
    severity: 'Low',
    description: 'Default OWA mailbox policy: AdditionalStorageProvidersAvailable must be False.',
    async run() { return exoStub('CIS.M365.6.5.3', 'Get-OwaMailboxPolicy'); },
  });

  // ---------- 7.x ThirdPartyStorageServicesRestricted ----------
  tests.push({
    id: 'CIS.M365.7.2.4', title: '(L2) Ensure third-party storage services are restricted in M365 on the web',
    severity: 'Medium',
    description: 'Service principal with appId c1f33bc0-bdb4-4248-ba9b-096807ddb43e (Office 365 third-party storage) should not be enabled.',
    async run() {
      const start = performance.now();
      try {
        const sps = await Graph.graphAll(`servicePrincipals?$filter=appId eq 'c1f33bc0-bdb4-4248-ba9b-096807ddb43e'`, { apiVersion: 'v1.0' });
        if (sps.length === 0) return row('CIS.M365.7.2.4', 'Skipped', 'Third-party storage service principal not present in this tenant; cannot evaluate setting.', { durationMs: ms(start) });
        const enabled = sps.find(s => s.accountEnabled === true);
        const ok = !enabled;
        return row('CIS.M365.7.2.4', ok ? 'Passed' : 'Failed', ok ? 'Third-party storage SP is disabled.' : 'Third-party storage SP is enabled.', { actual: !!enabled, durationMs: ms(start) });
      } catch (e) { return errRow('CIS.M365.7.2.4', e, start); }
    },
  });

  // ---------- 7.x ThirdPartyFileSharing (Teams) ----------
  tests.push({
    id: 'CIS.M365.7.2.5', title: '(L2) Ensure third-party file-sharing cloud services in Teams are disabled',
    severity: 'Medium',
    description: 'Get-CsTeamsClientConfiguration: AllowDropbox/Box/GoogleDrive/ShareFile/Egnyte all False.',
    async run() {
      const start = performance.now();
      try {
        const p = await getFirstAvailablePolicy([
          { path: 'teamwork/teamsAppSettings', apiVersion: 'v1.0', tokenScopes: Auth.SCOPES.graphFull },
          { path: 'teamwork/teamsAppSettings', apiVersion: 'beta', tokenScopes: Auth.SCOPES.graphFull },
        ]);
        const customAppSettings = p?.customAppSettings || {};
        const providers = {
          allowDropbox: pickAny(customAppSettings, ['allowDropbox', 'AllowDropbox']) ?? pickAny(p, ['allowDropbox', 'AllowDropbox']),
          allowBox: pickAny(customAppSettings, ['allowBox', 'AllowBox']) ?? pickAny(p, ['allowBox', 'AllowBox']),
          allowGoogleDrive: pickAny(customAppSettings, ['allowGoogleDrive', 'AllowGoogleDrive']) ?? pickAny(p, ['allowGoogleDrive', 'AllowGoogleDrive']),
          allowShareFile: pickAny(customAppSettings, ['allowShareFile', 'AllowShareFile']) ?? pickAny(p, ['allowShareFile', 'AllowShareFile']),
          allowEgnyte: pickAny(customAppSettings, ['allowEgnyte', 'AllowEgnyte']) ?? pickAny(p, ['allowEgnyte', 'AllowEgnyte']),
        };
        const keys = Object.keys(providers).filter(k => typeof providers[k] === 'boolean');
        if (!keys.length) {
          return row('CIS.M365.7.2.5', 'Skipped', 'Teams client configuration endpoint reachable, but third-party storage provider flags are not exposed.', { durationMs: ms(start) });
        }
        const enabled = keys.filter(k => providers[k] === true);
        const ok = enabled.length === 0;
        return row('CIS.M365.7.2.5', ok ? 'Passed' : 'Failed', ok ? 'All third-party file-sharing providers are disabled in Teams.' : `Enabled providers: ${enabled.join(', ')}`, { actual: providers, durationMs: ms(start) });
      } catch (e) { return teamsSkip('CIS.M365.7.2.5', e, start); }
    },
  });

  // ---------- 8.5.1 TeamsLobbyBypass (Teams) ----------
  tests.push({
    id: 'CIS.M365.8.5.1', title: '(L1) Ensure only people in my org can bypass the lobby',
    severity: 'Medium',
    description: 'Get-CsTeamsMeetingPolicy global: AutoAdmittedUsers in (InvitedUsers, EveryoneInCompanyExcludingGuests, OrganizerOnly).',
    async run() {
      const start = performance.now();
      try {
        const p = await getFirstAvailablePolicy([
          'admin/teams/meetingPolicies/global',
          'admin/teams/settings',
        ]);
        const v = String(pickAny(p, ['autoAdmittedUsers', 'AutoAdmittedUsers']) || '');
        if (!v) {
          return row('CIS.M365.8.5.1', 'Skipped', 'Teams meeting policy endpoint reachable, but AutoAdmittedUsers is not exposed.', { durationMs: ms(start) });
        }
        const allowed = new Set(['InvitedUsers', 'EveryoneInCompanyExcludingGuests', 'OrganizerOnly']);
        const ok = allowed.has(v);
        return row('CIS.M365.8.5.1', ok ? 'Passed' : 'Failed', ok ? `Lobby bypass is restricted (${v}).` : `Lobby bypass is too broad (${v}).`, { actual: v, durationMs: ms(start) });
      } catch (e) { return teamsSkip('CIS.M365.8.5.1', e, start); }
    },
  });

  // ---------- 8.x TeamsReportSecurityConcerns (Teams) ----------
  tests.push({
    id: 'CIS.M365.8.6.1', title: '(L1) Ensure users can report security concerns in Teams',
    severity: 'Low',
    description: 'Get-CsTeamsMessagingPolicy + Get-ReportSubmissionPolicy combination. Requires Teams admin REST.',
    async run() {
      const start = performance.now();
      try {
        const p = await getFirstAvailablePolicy([
          'admin/teams/messagingPolicies/global',
          'admin/teams/settings',
        ]);
        const allow = pickAny(p, ['allowSecurityEndUserReporting', 'AllowSecurityEndUserReporting']);
        if (typeof allow !== 'boolean') {
          return row('CIS.M365.8.6.1', 'Skipped', 'Teams messaging policy endpoint reachable, but AllowSecurityEndUserReporting is not exposed.', { durationMs: ms(start) });
        }
        const ok = allow === true;
        return row('CIS.M365.8.6.1', ok ? 'Passed' : 'Failed', ok ? 'Users can report security concerns in Teams.' : 'Users cannot report security concerns in Teams.', { actual: allow, durationMs: ms(start) });
      } catch (e) { return teamsSkip('CIS.M365.8.6.1', e, start); }
    },
  });

  // ---------- 8.x CommunicateInitiateExternalTeamsUsers (Teams) ----------
  tests.push({
    id: 'CIS.M365.8.2.2', title: '(L1) Ensure communication with unmanaged Teams users is restricted',
    severity: 'Medium',
    description: 'Get-CsTenantFederationConfiguration. Requires Teams admin REST.',
    async run() {
      const start = performance.now();
      try {
        const p = await getFirstAvailablePolicy([
          'admin/teams/federationConfiguration',
          'admin/teams/externalAccessPolicy/global',
          'admin/teams/settings',
        ]);
        const allowConsumer = pickAny(p, ['allowTeamsConsumer', 'AllowTeamsConsumer', 'enableTeamsConsumerAccess', 'EnableTeamsConsumerAccess']);
        if (typeof allowConsumer !== 'boolean') {
          return row('CIS.M365.8.2.2', 'Skipped', 'Teams federation/external access endpoint reachable, but unmanaged-consumer flags are not exposed.', { durationMs: ms(start) });
        }
        const ok = allowConsumer === false;
        return row('CIS.M365.8.2.2', ok ? 'Passed' : 'Failed', ok ? 'Communication with unmanaged Teams users is restricted.' : 'Communication with unmanaged Teams users is enabled.', { actual: allowConsumer, durationMs: ms(start) });
      } catch (e) { return teamsSkip('CIS.M365.8.2.2', e, start); }
    },
  });

  // ---------- 8.x CommunicateWithUnmanagedTeamsUsers (Teams) ----------
  tests.push({
    id: 'CIS.M365.8.2.1', title: '(L1) Ensure external access in Teams is configured per CIS',
    severity: 'Medium',
    description: 'Get-CsExternalAccessPolicy global. Requires Teams admin REST.',
    async run() {
      const start = performance.now();
      try {
        const p = await getFirstAvailablePolicy([
          'admin/teams/federationConfiguration',
          'admin/teams/externalAccessPolicy/global',
          'admin/teams/settings',
        ]);
        const allowInbound = pickAny(p, ['allowTeamsConsumerInbound', 'AllowTeamsConsumerInbound', 'enableTeamsConsumerInbound', 'EnableTeamsConsumerInbound']);
        if (typeof allowInbound !== 'boolean') {
          return row('CIS.M365.8.2.1', 'Skipped', 'Teams federation/external access endpoint reachable, but inbound-consumer flags are not exposed.', { durationMs: ms(start) });
        }
        const ok = allowInbound === false;
        return row('CIS.M365.8.2.1', ok ? 'Passed' : 'Failed', ok ? 'External unmanaged Teams users cannot initiate conversations.' : 'External unmanaged Teams users can initiate conversations.', { actual: allowInbound, durationMs: ms(start) });
      } catch (e) { return teamsSkip('CIS.M365.8.2.1', e, start); }
    },
  });

  // -------------------- catalog wrap --------------------
  function buildCatalog() {
    return tests.map(t => ({
      id: t.id, title: t.title, severity: t.severity, tag: 'CIS', category: 'CIS',
      docUrl: `https://maester.dev/docs/tests/${t.id}`,
      runCategory: 'CIS',
      description: t.description,
      detailMd: t.detailMd,
      requiredScopes: t.requiredScopes || Auth.SCOPES.graphFull,
      implemented: true,
      async run(ctx) {
        const r = await t.run(ctx);
        return r;
      },
    }));
  }

  window.TestsCIS = { buildCatalog };
})();
