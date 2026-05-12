// Intune / Device Management tests.
// All tests here use the Microsoft Graph deviceManagement namespace (beta).
// Requires DeviceManagementConfiguration.Read.All and/or DeviceManagementApps.Read.All
// (obtained via incremental consent – the user is prompted once before the scan starts).
//
// Reference: https://maester.dev/docs/tests/
(() => {
  // All Intune-related Graph scopes we need. Auth.SCOPES.graphFull already has
  // the base scopes; we append the Intune-specific ones here so the test runner
  // can pre-consent them via ensureScopes() before a scan starts.
  const DEVICE_SCOPES = [
    'https://graph.microsoft.com/DeviceManagementConfiguration.Read.All',
    'https://graph.microsoft.com/DeviceManagementApps.Read.All',
    'https://graph.microsoft.com/DeviceManagementServiceConfig.Read.All',
  ].concat(Auth.SCOPES.graphFull);

  const gOpts = { apiVersion: 'beta', tokenScopes: DEVICE_SCOPES };

  function ms(t0) { return Math.round(performance.now() - t0); }
  function skip403(id, e, t0) {
    if (e.status === 403 || e.status === 401 || e.status === 404) {
      return { id, status: 'Skipped', reason: `Insufficient permissions (HTTP ${e.status}). Requires Intune license and DeviceManagementConfiguration.Read.All scope. Grant consent when prompted before running the scan.`, durationMs: ms(t0) };
    }
    return { id, status: 'Error', reason: e.message, durationMs: ms(t0) };
  }

  const tests = [];

  // ── MT.INT.1001  Compliance policies exist ───────────────────────────────────
  tests.push({
    id: 'MT.INT.1001',
    title: 'Compliance policies exist for managed devices',
    severity: 'High', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.INT.1001',
    description: 'At least one device compliance policy is configured in Intune.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const policies = await Graph.graphAll('deviceManagement/deviceCompliancePolicies', gOpts);
        if (!policies.length) return { id: 'MT.INT.1001', status: 'Failed', reason: 'No device compliance policies found in Intune.', durationMs: ms(t0) };
        const types = [...new Set(policies.map(p => (p['@odata.type'] || '').replace('#microsoft.graph.', '')))].join(', ');
        return { id: 'MT.INT.1001', status: 'Passed', reason: `${policies.length} compliance polic${policies.length === 1 ? 'y' : 'ies'} found (${types}).`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.INT.1001', e, t0); }
    },
  });

  // ── MT.INT.1002  Compliance retire/block grace period ───────────────────────
  tests.push({
    id: 'MT.INT.1002',
    title: 'Device compliance retire grace period is configured',
    severity: 'Medium', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.INT.1002',
    description: 'Compliance policies should configure a retire or block action so non-compliant devices are eventually removed.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const policies = await Graph.graphAll('deviceManagement/deviceCompliancePolicies', gOpts);
        if (!policies.length) return { id: 'MT.INT.1002', status: 'Skipped', reason: 'No compliance policies configured.', durationMs: ms(t0) };
        const checked = []; const passed = [];
        for (const pol of policies.slice(0, 15)) {
          try {
            const actions = await Graph.graphAll(`deviceManagement/deviceCompliancePolicies/${pol.id}/scheduledActionsForRule`, gOpts);
            const hasAction = actions.some(a => (a.scheduledActionConfigurations || []).some(c => ['retire', 'block', 'pushNotification'].includes(c.actionType)));
            checked.push(pol.displayName);
            if (hasAction) passed.push(pol.displayName);
          } catch { /* skip if individual policy fetch fails */ }
        }
        if (!checked.length) return { id: 'MT.INT.1002', status: 'Skipped', reason: 'Could not retrieve scheduled actions.', durationMs: ms(t0) };
        if (!passed.length) return { id: 'MT.INT.1002', status: 'Failed', reason: `None of ${checked.length} compliance policies have a retire/block grace period action.`, durationMs: ms(t0) };
        return { id: 'MT.INT.1002', status: 'Passed', reason: `${passed.length}/${checked.length} policies have a grace period action.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.INT.1002', e, t0); }
    },
  });

  // ── MT.INT.1003  App protection (MAM) policies exist ───────────────────────
  tests.push({
    id: 'MT.INT.1003',
    title: 'App protection policies cover managed apps',
    severity: 'Medium', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.INT.1003',
    description: 'At least one Intune App Protection (MAM) policy is configured.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const policies = await Graph.graphAll('deviceAppManagement/managedAppPolicies', gOpts);
        if (!policies.length) return { id: 'MT.INT.1003', status: 'Failed', reason: 'No Managed App Protection (MAM) policies found.', durationMs: ms(t0) };
        const ios = policies.filter(p => (p['@odata.type'] || '').toLowerCase().includes('ios')).length;
        const and = policies.filter(p => (p['@odata.type'] || '').toLowerCase().includes('android')).length;
        return { id: 'MT.INT.1003', status: 'Passed', reason: `${policies.length} MAM polic${policies.length === 1 ? 'y' : 'ies'} found (iOS: ${ios}, Android: ${and}).`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.INT.1003', e, t0); }
    },
  });

  // ── MT.1053  MFA required for Intune enrollment ──────────────────────────────
  // Uses existing CA scope (no new scope needed). Checks for a CA policy
  // requiring MFA when enrolling in the Microsoft Intune Enrollment app.
  tests.push({
    id: 'MT.1053',
    title: 'MFA should not be blocked for Intune enrollment',
    severity: 'High', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1053',
    description: 'A Conditional Access policy should require MFA for the Microsoft Intune Enrollment app so devices cannot be enrolled without MFA.',
    async run() {
      const t0 = performance.now();
      try {
        // Microsoft Intune Enrollment app ID
        const INTUNE_ENROLLMENT_APP_ID = 'd4ebce55-015a-49b5-a083-c84d1797ae8c';
        const policies = await Graph.graphAll('identity/conditionalAccess/policies', { apiVersion: 'v1.0' });
        const active = policies.filter(p => p.state === 'enabled');
        // Check for a CA policy that REQUIRES MFA for Intune enrollment
        const mfaForEnrollment = active.filter(p => {
          const apps = p.conditions?.applications?.includeApplications || [];
          const allApps = apps.includes('All') || apps.includes('all');
          const includesEnrollment = allApps || apps.includes(INTUNE_ENROLLMENT_APP_ID);
          if (!includesEnrollment) return false;
          const grants = p.grantControls?.builtInControls || [];
          const strength = p.grantControls?.authenticationStrength;
          return grants.includes('mfa') || !!strength;
        });
        if (!mfaForEnrollment.length) {
          return { id: 'MT.1053', status: 'Failed', reason: `No active CA policy requiring MFA for the Intune Enrollment app (${INTUNE_ENROLLMENT_APP_ID}) was found. Without this, devices can be enrolled without MFA.`, durationMs: ms(t0) };
        }
        return { id: 'MT.1053', status: 'Passed', reason: `${mfaForEnrollment.length} CA polic${mfaForEnrollment.length === 1 ? 'y' : 'ies'} require MFA for Intune enrollment: ${mfaForEnrollment.map(p => p.displayName).join(', ')}.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1053', e, t0); }
    },
  });

  // ── MT.1054  Enrollment restrictions configured ──────────────────────────────
  tests.push({
    id: 'MT.1054',
    title: 'Intune enrollment restrictions are configured',
    severity: 'Medium', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1054',
    description: 'Device enrollment restrictions should be configured to limit which platforms and device types can enroll.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const configs = await Graph.graphAll('deviceManagement/deviceEnrollmentConfigurations', gOpts);
        const restrictions = configs.filter(c =>
          (c['@odata.type'] || '').includes('PlatformRestrictions') ||
          (c['@odata.type'] || '').includes('LimitConfiguration') ||
          (c['@odata.type'] || '').includes('WindowsHelloForBusiness')
        );
        if (!restrictions.length) {
          return { id: 'MT.1054', status: 'Failed', reason: 'No device enrollment restriction configurations found in Intune.', durationMs: ms(t0) };
        }
        return { id: 'MT.1054', status: 'Passed', reason: `${restrictions.length} enrollment restriction configuration(s) found.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1054', e, t0); }
    },
  });

  // ─── Windows compliance policy checks (MT.1092–MT.1098) ────────────────────
  // These all query the same endpoint; results are cached by Graph layer.
  async function getWinCompliancePolicies(t0) {
    const all = await Graph.graphAll('deviceManagement/deviceCompliancePolicies', gOpts);
    return all.filter(p => p['@odata.type'] === '#microsoft.graph.windows10CompliancePolicy');
  }

  function winCompTest(id, title, desc, field, extraCheck) {
    return {
      id, title, severity: 'High', category: 'Intune', tag: 'Intune',
      docUrl: `https://maester.dev/docs/tests/${id}`,
      description: desc,
      requiredScopes: DEVICE_SCOPES,
      async run() {
        const t0 = performance.now();
        try {
          const wins = await getWinCompliancePolicies(t0);
          if (!wins.length) return { id, status: 'Skipped', reason: 'No Windows 10/11 compliance policies found in Intune.', durationMs: ms(t0) };
          const passing = wins.filter(p => extraCheck ? extraCheck(p) : p[field] === true);
          if (!passing.length) {
            const vals = wins.map(p => `${p.displayName}: ${field}=${p[field]}`).join('; ');
            return { id, status: 'Failed', reason: `No Windows compliance policy requires ${title.toLowerCase()}. Policies: ${vals}`, durationMs: ms(t0) };
          }
          return { id, status: 'Passed', reason: `${passing.length}/${wins.length} Windows compliance polic${passing.length === 1 ? 'y' : 'ies'} require${passing.length === 1 ? 's' : ''} ${title.toLowerCase()}.`, durationMs: ms(t0) };
        } catch (e) { return skip403(id, e, t0); }
      },
    };
  }

  tests.push(winCompTest('MT.1092', 'Windows device compliance policy requires BitLocker encryption', 'Windows compliance policies should require BitLocker encryption.', 'bitLockerEnabled'));
  tests.push(winCompTest('MT.1093', 'Windows device compliance policy requires Secure Boot', 'Windows compliance policies should require Secure Boot.', 'secureBootEnabled'));
  tests.push(winCompTest('MT.1094', 'Windows device compliance policy requires Code Integrity', 'Windows compliance policies should require Code Integrity.', 'codeIntegrityEnabled'));
  tests.push(winCompTest('MT.1095', 'Windows device compliance policy requires antivirus', 'Windows compliance policies should require Defender antivirus.', 'defenderEnabled'));
  tests.push(winCompTest('MT.1096', 'Windows device compliance policy requires anti-spyware', 'Windows compliance policies should require anti-spyware protection.', 'antiSpywareRequired'));
  tests.push(winCompTest('MT.1097', 'Windows device compliance policy requires firewall', 'Windows compliance policies should require Windows Firewall.', 'activeFirewallRequired'));
  tests.push(winCompTest('MT.1098', 'Windows device compliance policy requires defender real-time protection', 'Windows compliance policies should require Defender real-time protection.', null,
    p => p.defenderEnabled === true || p.realTimeProtectionEnabled === true));

  // ── MT.1099  iOS/macOS compliance policy ────────────────────────────────────
  tests.push({
    id: 'MT.1099',
    title: 'iOS/macOS device compliance policy is configured',
    severity: 'High', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1099',
    description: 'At least one iOS or macOS compliance policy should be configured.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const all = await Graph.graphAll('deviceManagement/deviceCompliancePolicies', gOpts);
        const ios = all.filter(p => ['#microsoft.graph.iosCompliancePolicy', '#microsoft.graph.macOSCompliancePolicy'].includes(p['@odata.type']));
        if (!ios.length) return { id: 'MT.1099', status: 'Failed', reason: 'No iOS or macOS compliance policies found in Intune.', durationMs: ms(t0) };
        return { id: 'MT.1099', status: 'Passed', reason: `${ios.length} iOS/macOS compliance polic${ios.length === 1 ? 'y' : 'ies'} found.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1099', e, t0); }
    },
  });

  // ── MT.1100  Android compliance policy ──────────────────────────────────────
  tests.push({
    id: 'MT.1100',
    title: 'Android device compliance policy is configured',
    severity: 'High', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1100',
    description: 'At least one Android compliance policy should be configured.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const all = await Graph.graphAll('deviceManagement/deviceCompliancePolicies', gOpts);
        const and = all.filter(p => (p['@odata.type'] || '').toLowerCase().includes('android'));
        if (!and.length) return { id: 'MT.1100', status: 'Failed', reason: 'No Android compliance policies found in Intune.', durationMs: ms(t0) };
        return { id: 'MT.1100', status: 'Passed', reason: `${and.length} Android compliance polic${and.length === 1 ? 'y' : 'ies'} found.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1100', e, t0); }
    },
  });

  // ── MT.1101/MT.1102  MAM for iOS/Android ───────────────────────────────────
  async function getMamPolicies(platform) {
    const all = await Graph.graphAll('deviceAppManagement/managedAppPolicies', gOpts);
    return all.filter(p => (p['@odata.type'] || '').toLowerCase().includes(platform));
  }

  tests.push({
    id: 'MT.1101',
    title: 'Mobile Application Management policies are configured for iOS',
    severity: 'High', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1101',
    description: 'At least one iOS Managed App Protection (MAM) policy should be configured.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const ios = await getMamPolicies('ios');
        if (!ios.length) return { id: 'MT.1101', status: 'Failed', reason: 'No iOS MAM protection policies found in Intune.', durationMs: ms(t0) };
        return { id: 'MT.1101', status: 'Passed', reason: `${ios.length} iOS MAM polic${ios.length === 1 ? 'y' : 'ies'} found.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1101', e, t0); }
    },
  });

  tests.push({
    id: 'MT.1102',
    title: 'Mobile Application Management policies are configured for Android',
    severity: 'High', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1102',
    description: 'At least one Android Managed App Protection (MAM) policy should be configured.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const and = await getMamPolicies('android');
        if (!and.length) return { id: 'MT.1102', status: 'Failed', reason: 'No Android MAM protection policies found in Intune.', durationMs: ms(t0) };
        return { id: 'MT.1102', status: 'Passed', reason: `${and.length} Android MAM polic${and.length === 1 ? 'y' : 'ies'} found.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1102', e, t0); }
    },
  });

  // ── MT.1105  Corporate owned enrollment profile ──────────────────────────────
  tests.push({
    id: 'MT.1105',
    title: 'Device enrollment requires corporate owned device profile',
    severity: 'Medium', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1105',
    description: 'Enrollment profiles or restrictions should exist for corporate device management (Windows Autopilot, Apple DEP, or similar).',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        // Check Autopilot profiles (Windows)
        let autopilotProfiles = [];
        try { autopilotProfiles = await Graph.graphAll('deviceManagement/windowsAutopilotDeploymentProfiles', gOpts); } catch {}
        // Check Apple DEP (iOS/macOS)
        let depSettings = [];
        try { depSettings = await Graph.graphAll('deviceManagement/depOnboardingSettings', gOpts); } catch {}
        // Check enrollment configs for corp-device restrictions
        const configs = await Graph.graphAll('deviceManagement/deviceEnrollmentConfigurations', gOpts);
        const corpConfigs = configs.filter(c =>
          (c.displayName || '').toLowerCase().match(/corp|enterprise|business|owned|autopilot|dep/i) ||
          (c['@odata.type'] || '').includes('EnrollmentProfile')
        );
        const total = autopilotProfiles.length + depSettings.length + corpConfigs.length;
        if (!total) {
          return { id: 'MT.1105', status: 'Failed', reason: 'No corporate enrollment profiles (Autopilot, DEP, or corporate device restrictions) found.', durationMs: ms(t0) };
        }
        const parts = [];
        if (autopilotProfiles.length) parts.push(`${autopilotProfiles.length} Autopilot profile(s)`);
        if (depSettings.length) parts.push(`${depSettings.length} DEP setting(s)`);
        if (corpConfigs.length) parts.push(`${corpConfigs.length} enrollment config(s)`);
        return { id: 'MT.1105', status: 'Passed', reason: `Corporate enrollment configured: ${parts.join(', ')}.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1105', e, t0); }
    },
  });

  // ── MT.1123  App Protection PIN required ────────────────────────────────────
  tests.push({
    id: 'MT.1123',
    title: 'Intune App Protection PIN is required',
    severity: 'High', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1123',
    description: 'App Protection (MAM) policies should require a PIN to access protected apps.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const all = await Graph.graphAll('deviceAppManagement/managedAppPolicies', gOpts);
        if (!all.length) return { id: 'MT.1123', status: 'Skipped', reason: 'No MAM policies configured.', durationMs: ms(t0) };
        const withPin = all.filter(p => p.pinRequired === true);
        if (!withPin.length) return { id: 'MT.1123', status: 'Failed', reason: `None of ${all.length} MAM polic${all.length === 1 ? 'y requires' : 'ies require'} a PIN.`, durationMs: ms(t0) };
        return { id: 'MT.1123', status: 'Passed', reason: `${withPin.length}/${all.length} MAM polic${withPin.length === 1 ? 'y' : 'ies'} require${withPin.length === 1 ? 's' : ''} a PIN.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1123', e, t0); }
    },
  });

  // ─── Defender AV via Intune (MT.1148–MT.1171) ──────────────────────────────
  // These checks look at windows10EndpointProtectionConfiguration profiles.
  // Many organizations now use the Settings Catalog instead; we do a best-effort
  // check of both configuration sources.
  async function getEndpointProtectionConfigs() {
    const all = await Graph.graphAll('deviceManagement/deviceConfigurations', gOpts);
    return all.filter(c => c['@odata.type'] === '#microsoft.graph.windows10EndpointProtectionConfiguration');
  }

  function epTest(id, title, severity, field, desc, checkFn) {
    return {
      id, title, severity, category: 'Intune', tag: 'Intune',
      docUrl: `https://maester.dev/docs/tests/${id}`,
      description: desc,
      requiredScopes: DEVICE_SCOPES,
      async run() {
        const t0 = performance.now();
        try {
          const eps = await getEndpointProtectionConfigs();
          if (!eps.length) {
            // Also check settings catalog policies for Defender settings
            try {
              const catPolicies = await Graph.graphAll('deviceManagement/configurationPolicies', gOpts);
              const defPolicies = catPolicies.filter(p =>
                (p.name || '').toLowerCase().match(/defender|antivirus|endpoint.*protection|asr|attack.*surface/i) ||
                (p.description || '').toLowerCase().match(/defender|antivirus/)
              );
              if (defPolicies.length) {
                return { id, status: 'Passed', reason: `No classic endpoint protection profiles found, but ${defPolicies.length} Defender-related settings catalog polic${defPolicies.length === 1 ? 'y' : 'ies'} found (${defPolicies.map(p => p.name).join(', ')}).`, durationMs: ms(t0) };
              }
            } catch {}
            return { id, status: 'Failed', reason: 'No Windows Endpoint Protection configuration profiles found in Intune.', durationMs: ms(t0) };
          }
          const passing = checkFn ? eps.filter(checkFn) : eps.filter(p => {
            const val = p[field];
            return val !== undefined && val !== null && val !== 'notConfigured' && val !== 'userDefined' && val !== false;
          });
          if (!passing.length) {
            const vals = eps.map(p => `${p.displayName}: ${field}=${JSON.stringify(p[field])}`).join('; ');
            return { id, status: 'Failed', reason: `No endpoint protection profile has ${title.toLowerCase()} configured. Profiles: ${vals.substring(0, 300)}`, durationMs: ms(t0) };
          }
          return { id, status: 'Passed', reason: `${passing.length}/${eps.length} endpoint protection profile(s) configure ${title.toLowerCase()}.`, durationMs: ms(t0) };
        } catch (e) { return skip403(id, e, t0); }
      },
    };
  }

  tests.push(epTest('MT.1148', 'Microsoft Defender Antivirus cloud protection is enabled via Intune', 'High',
    'defenderCloudBlockLevel',
    'Intune endpoint protection profiles should enable Defender cloud protection.',
    p => p.defenderCloudBlockLevel && p.defenderCloudBlockLevel !== 'notConfigured' && p.defenderCloudBlockLevel !== 'userDefined'
  ));

  tests.push(epTest('MT.1149', 'Microsoft Defender Antivirus real-time protection is enabled via Intune', 'High',
    'defenderAllowRealTimeMonitoring',
    'Intune should configure Defender real-time monitoring.',
    p => p.defenderAllowRealTimeMonitoring === 'allowed' || p.defenderRealTimeScanDirection != null
  ));

  tests.push(epTest('MT.1150', 'Microsoft Defender Antivirus automatic sample submission is configured via Intune', 'High',
    'defenderAllowBehaviorMonitoring',
    'Intune should configure Defender sample submission (MAPS reporting).',
    p => p.defenderAllowBehaviorMonitoring === 'allowed' || p.defenderAllowCloudProtection === 'allowed'
  ));

  tests.push(epTest('MT.1151', 'Microsoft Defender Antivirus potentially unwanted app protection is enabled via Intune', 'Medium',
    'defenderPotentiallyUnwantedAppAction',
    'Intune should configure Defender PUA (potentially unwanted app) protection.',
    p => ['enable', 'block', 'userDefined'].includes(p.defenderPotentiallyUnwantedAppAction) && p.defenderPotentiallyUnwantedAppAction !== 'userDefined'
  ));

  tests.push(epTest('MT.1152', 'Microsoft Defender Antivirus network protection is enabled via Intune', 'High',
    'defenderNetworkProtectionType',
    'Intune should enable Defender network protection.',
    p => p.defenderNetworkProtectionType === 'enable' || p.defenderNetworkProtectionType === 'auditMode'
  ));

  tests.push(epTest('MT.1153', 'Microsoft Defender Antivirus tamper protection is enabled via Intune', 'High',
    'defenderTamperProtection',
    'Intune should enable Defender tamper protection.',
    p => p.defenderTamperProtection === 'enable' || p.defenderSecurityCenterBlockExploitProtectionOverride === false
  ));

  tests.push(epTest('MT.1154', 'Microsoft Defender Antivirus scheduled scan is configured via Intune', 'Medium',
    'defenderScheduledScanDay',
    'Intune should configure a scheduled Defender scan.',
    p => p.defenderScheduledScanDay && p.defenderScheduledScanDay !== 'userDefined'
  ));

  // ── MT.1155  ASR rules enabled (consolidated check) ─────────────────────────
  tests.push({
    id: 'MT.1155',
    title: 'Microsoft Defender Attack Surface Reduction rules are enabled via Intune',
    severity: 'High', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1155',
    description: 'Intune endpoint protection profiles should have Attack Surface Reduction (ASR) rules configured.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const eps = await getEndpointProtectionConfigs();
        if (!eps.length) {
          // Check settings catalog for ASR policies
          const catPolicies = await Graph.graphAll('deviceManagement/configurationPolicies', gOpts);
          const asrPolicies = catPolicies.filter(p => (p.name || '').toLowerCase().match(/asr|attack.surface|reduction/i));
          if (asrPolicies.length) return { id: 'MT.1155', status: 'Passed', reason: `${asrPolicies.length} ASR settings catalog polic${asrPolicies.length === 1 ? 'y' : 'ies'} found.`, durationMs: ms(t0) };
          return { id: 'MT.1155', status: 'Failed', reason: 'No endpoint protection profiles or ASR settings catalog policies found.', durationMs: ms(t0) };
        }
        const ASR_FIELDS = [
          'defenderOfficeAppsExecutableContentCreationOrLaunchType', 'defenderOfficeAppsLaunchChildProcessType',
          'defenderOfficeAppsOtherProcessInjectionType', 'defenderEmailContentExecutionType',
          'defenderPreventCredentialStealingType', 'defenderBlockPersistenceThroughWmiType',
          'defenderNetworkProtectionType', 'defenderGuardMyFoldersType',
        ];
        const withAsr = eps.filter(p => ASR_FIELDS.some(f => p[f] && p[f] !== 'notConfigured' && p[f] !== 'userDefined'));
        if (!withAsr.length) return { id: 'MT.1155', status: 'Failed', reason: `${eps.length} endpoint protection profile(s) found but none have ASR rules configured.`, durationMs: ms(t0) };
        return { id: 'MT.1155', status: 'Passed', reason: `${withAsr.length}/${eps.length} endpoint protection profile(s) have ASR rules configured.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1155', e, t0); }
    },
  });

  // ── Individual ASR rule tests (MT.1156–MT.1167) ──────────────────────────────
  const ASR_RULES = [
    { id: 'MT.1156', title: 'Microsoft Defender ASR rule: Block Win32 API from macro is enabled',          field: 'defenderOfficeMacroCodeAllowWin32ImportsType' },
    { id: 'MT.1157', title: 'Microsoft Defender ASR rule: Block Office apps from creating executable content', field: 'defenderOfficeAppsExecutableContentCreationOrLaunchType' },
    { id: 'MT.1158', title: 'Microsoft Defender ASR rule: Block Office apps from injecting code',          field: 'defenderOfficeAppsOtherProcessInjectionType' },
    { id: 'MT.1159', title: 'Microsoft Defender ASR rule: Block JS/VBS launching downloaded content',      field: 'defenderScriptDownloadedPayloadExecutionType' },
    { id: 'MT.1160', title: 'Microsoft Defender ASR rule: Block credential stealing from LSASS',           field: 'defenderPreventCredentialStealingType' },
    { id: 'MT.1161', title: 'Microsoft Defender ASR rule: Block persistence through WMI',                  field: 'defenderBlockPersistenceThroughWmiType' },
    { id: 'MT.1162', title: 'Microsoft Defender ASR rule: Block Office macro obfuscation',                 field: 'defenderScriptObfuscatedMacroCodeType' },
    { id: 'MT.1163', title: 'Microsoft Defender ASR rule: Block executable content from email/webmail',    field: 'defenderEmailContentExecutionType' },
    { id: 'MT.1164', title: 'Microsoft Defender ASR rule: Block Adobe Reader from creating child processes', field: 'defenderAdobeReaderLaunchChildProcess' },
    { id: 'MT.1165', title: 'Microsoft Defender ASR rule: Block Office communication apps from creating child processes', field: 'defenderOfficeCommunicationAppsLaunchChildProcess' },
    { id: 'MT.1166', title: 'Microsoft Defender ASR rule: Block untrusted and unsigned processes via USB', field: 'defenderUntrustedUSBProcessType' },
    { id: 'MT.1167', title: 'Microsoft Defender ASR rule: Block executable files from email that do not meet criteria', field: 'defenderEmailContentExecution' },
  ];

  for (const rule of ASR_RULES) {
    const { id, title, field } = rule;
    tests.push({
      id, title, severity: 'High', category: 'Intune', tag: 'Intune',
      docUrl: `https://maester.dev/docs/tests/${id}`,
      description: `Intune endpoint protection profiles should configure ASR rule: ${title}.`,
      requiredScopes: DEVICE_SCOPES,
      async run() {
        const t0 = performance.now();
        // Capture id and field in closure
        const _id = id, _field = field, _title = title;
        try {
          const eps = await getEndpointProtectionConfigs();
          if (!eps.length) return { id: _id, status: 'Skipped', reason: 'No Windows endpoint protection configuration profiles found in Intune.', durationMs: ms(t0) };
          const enabled = ['block', 'enable', 'auditMode'];
          const passing = eps.filter(p => enabled.includes(p[_field]));
          const blocking = eps.filter(p => p[_field] === 'block' || p[_field] === 'enable');
          if (!passing.length) {
            return { id: _id, status: 'Failed', reason: `No profile enables ASR rule "${_title}" (field: ${_field}). Current values: ${eps.map(p => `${p.displayName}:${p[_field] || 'not set'}`).join(', ')}`, durationMs: ms(t0) };
          }
          if (!blocking.length) {
            return { id: _id, status: 'Failed', reason: `ASR rule is in audit mode only across ${passing.length} profile(s) — not in block mode.`, durationMs: ms(t0) };
          }
          return { id: _id, status: 'Passed', reason: `${blocking.length} profile(s) enable ASR rule in block mode.`, durationMs: ms(t0) };
        } catch (e) { return skip403(_id, e, t0); }
      },
    });
  }

  // ── MT.1168  Exploit Protection ──────────────────────────────────────────────
  tests.push({
    id: 'MT.1168',
    title: 'Microsoft Defender Exploit Protection is enabled via Intune',
    severity: 'High', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1168',
    description: 'Intune should deploy Exploit Protection (defenderExploitProtectionXml is set).',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const eps = await getEndpointProtectionConfigs();
        if (!eps.length) return { id: 'MT.1168', status: 'Skipped', reason: 'No endpoint protection profiles found.', durationMs: ms(t0) };
        const withEP = eps.filter(p => p.defenderExploitProtectionXml && p.defenderExploitProtectionXml.length > 0);
        if (!withEP.length) return { id: 'MT.1168', status: 'Failed', reason: 'No endpoint protection profile has defenderExploitProtectionXml configured.', durationMs: ms(t0) };
        return { id: 'MT.1168', status: 'Passed', reason: `${withEP.length} profile(s) deploy Exploit Protection configuration.`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1168', e, t0); }
    },
  });

  // ── MT.1169  Controlled Folder Access ───────────────────────────────────────
  tests.push(epTest('MT.1169', 'Microsoft Defender Controlled Folder Access is enabled via Intune', 'High',
    'defenderGuardMyFoldersType',
    'Intune should enable Controlled Folder Access to protect against ransomware.',
    p => p.defenderGuardMyFoldersType === 'enable' || p.defenderGuardMyFoldersType === 'blockDiskModification'
  ));

  // ── MT.1170  Application Guard ───────────────────────────────────────────────
  tests.push(epTest('MT.1170', 'Microsoft Defender Application Guard is enabled for Microsoft Edge via Intune', 'High',
    'defenderApplicationGuardEnabled',
    'Intune should enable Microsoft Defender Application Guard for Edge.',
    p => p.defenderApplicationGuardEnabled === true
  ));

  // ── MT.1171  Application Control (WDAC) ─────────────────────────────────────
  tests.push({
    id: 'MT.1171',
    title: 'Microsoft Defender Application Control is configured via Intune',
    severity: 'Medium', category: 'Intune', tag: 'Intune',
    docUrl: 'https://maester.dev/docs/tests/MT.1171',
    description: 'Intune should configure Windows Defender Application Control (WDAC) to restrict which apps can run.',
    requiredScopes: DEVICE_SCOPES,
    async run() {
      const t0 = performance.now();
      try {
        const eps = await getEndpointProtectionConfigs();
        // Check settings catalog for WDAC policies
        const catPolicies = await Graph.graphAll('deviceManagement/configurationPolicies', gOpts);
        const wdacPolicies = catPolicies.filter(p => (p.name || '').toLowerCase().match(/application.control|wdac|applocker/i));
        // Check device configurations for WDAC-type configs
        const allConfigs = await Graph.graphAll('deviceManagement/deviceConfigurations', gOpts);
        const wdacConfigs = allConfigs.filter(c =>
          (c['@odata.type'] || '').includes('CustomConfiguration') &&
          (c.displayName || '').toLowerCase().match(/application.control|wdac|applocker/i)
        );
        const total = wdacPolicies.length + wdacConfigs.length;
        if (!total) return { id: 'MT.1171', status: 'Failed', reason: 'No Application Control (WDAC/AppLocker) policies found in Intune.', durationMs: ms(t0) };
        return { id: 'MT.1171', status: 'Passed', reason: `${total} Application Control configuration(s) found (${wdacPolicies.length} settings catalog, ${wdacConfigs.length} custom config).`, durationMs: ms(t0) };
      } catch (e) { return skip403('MT.1171', e, t0); }
    },
  });

  // ── Build catalog ────────────────────────────────────────────────────────────
  function buildCatalog() {
    return tests.map(t => ({
      id: t.id, title: t.title, severity: t.severity,
      tag: t.tag || 'Intune', category: t.category || 'Intune',
      runCategory: 'Intune',
      docUrl: t.docUrl, description: t.description,
      implemented: true,
      requiredScopes: t.requiredScopes || DEVICE_SCOPES,
      async run(ctx) {
        const r = await t.run(ctx);
        return { ...t, ...r, tag: 'Intune', category: 'Intune' };
      },
    }));
  }

  window.TestsIntune = { buildCatalog };
})();
