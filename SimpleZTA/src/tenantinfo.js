(() => {
  // Browser port of the module's Invoke-ZtTenantInfo stage: collects the aggregate datasets
  // that feed the report graphics (tenant KPIs, sankey diagrams, Intune config tables).
  // Output shape matches the module's TenantInfo keys exactly so both the official report
  // template and the native dashboard can consume it. Every collector degrades to null with
  // an explanatory note instead of failing the scan.

  const SIGNIN_DEFAULT_DAYS = 7;
  const SIGNIN_MAX_MS = 120000;
  const SIGNIN_MAX_ITEMS = 50000;
  const PAGE_MAX_MS = 180000;
  const PAGE_MAX_ITEMS = 100000;

  function percentLabel(value, total) {
    if (!(total > 0) || !(value > 0)) return '0%';
    const percent = (value / total) * 100;
    if (percent < 0 || percent > 100) return '0%';
    if (percent > 0 && percent < 1) return 'less than 1%';
    return `${Math.round(percent * 10) / 10}%`;
  }

  function durationLabel(minDate, maxDate) {
    if (!minDate || !maxDate) return '0 duration';
    const minutes = Math.floor((maxDate - minDate) / 60000);
    const hours = Math.floor(minutes / 60);
    const days = Math.floor(hours / 24);
    let value; let label;
    if (days > 0) { value = days; label = 'day'; }
    else if (hours > 0) { value = hours; label = 'hour'; }
    else if (minutes > 0) { value = minutes; label = 'minute'; }
    else { return '0 duration'; }
    return `${value} ${label}${value > 1 ? 's' : ''}`;
  }

  function node(source, target, value) {
    return { source, target, value };
  }

  async function graphCount(path, filter) {
    const value = await Api.graph(`${path}/$count`, {
      query: filter ? { $filter: filter } : {},
      headers: { ConsistencyLevel: 'eventual' },
    });
    const n = Number(value);
    return Number.isFinite(n) ? n : 0;
  }

  // Light port of Get-ZtLicenseInformation: detects Entra premium and Intune entitlements.
  async function getLicenses() {
    try {
      const skus = await Api.graphAll('subscribedSkus');
      const plans = [];
      for (const sku of skus || []) {
        if ((sku?.prepaidUnits?.enabled || 0) <= 0) continue;
        for (const plan of sku.servicePlans || []) plans.push(`${plan.servicePlanName || ''}`.toUpperCase());
      }
      return {
        entraPremium: plans.some(p => p.startsWith('AAD_PREMIUM')),
        intune: plans.some(p => p.startsWith('INTUNE_A') || p === 'INTUNE_EDU' || p === 'INTUNE_SMB'),
      };
    } catch {
      // If SKU lookup fails, attempt collectors anyway; they degrade individually.
      return { entraPremium: true, intune: true };
    }
  }

  async function collectTenantOverview() {
    const [userCount, guestCount, groupCount, applicationCount, deviceCount] = await Promise.all([
      graphCount('users', "userType ne 'Guest'"),
      graphCount('users', "userType eq 'Guest'"),
      graphCount('groups'),
      graphCount('applications'),
      graphCount('devices'),
    ]);

    let managedDeviceCount = 0;
    try {
      const managed = await Api.graph('deviceManagement/managedDeviceOverview', { beta: true });
      managedDeviceCount = Number(managed?.enrolledDeviceCount) || 0;
    } catch {
      managedDeviceCount = 0;
    }

    return {
      UserCount: userCount,
      GuestCount: guestCount,
      GroupCount: groupCount,
      ApplicationCount: applicationCount,
      DeviceCount: deviceCount,
      ManagedDeviceCount: managedDeviceCount,
    };
  }

  // One pass over userRegistrationDetails feeds both the all-users and privileged-users
  // sankeys. The module resolves privileged users via role assignments (vwRole); the browser
  // port uses the isAdmin flag Graph computes on the same report, which covers directory roles.
  async function collectAuthMethods(onProgress) {
    const { items, truncated } = await Api.graphPaged('reports/authenticationMethods/userRegistrationDetails', {
      query: { $top: 999 },
      maxItems: PAGE_MAX_ITEMS,
      maxMs: PAGE_MAX_MS,
      onPage: n => onProgress?.(`Auth methods: ${n} users processed`),
    });

    function bucket(users) {
      const has = (u, ...methods) => (u.methodsRegistered || []).some(m => methods.includes(m));
      const singleFactor = users.filter(u => !(u.methodsRegistered || []).length).length;
      const phone = users.filter(u => has(u, 'mobilePhone')).length;
      const authenticator = users.filter(u => has(u, 'microsoftAuthenticatorPush', 'softwareOneTimePasscode', 'microsoftAuthenticatorPasswordless')).length;
      const passkey = users.filter(u => has(u, 'passKeyDeviceBound', 'passKeyDeviceBoundAuthenticator')).length;
      const whfb = users.filter(u => has(u, 'windowsHelloForBusiness')).length;
      return [
        node('Users', 'Single factor', singleFactor),
        node('Users', 'Phishable', phone + authenticator),
        node('Phishable', 'Phone', phone),
        node('Phishable', 'Authenticator', authenticator),
        node('Users', 'Phish resistant', passkey + whfb),
        node('Phish resistant', 'Passkey', passkey),
        node('Phish resistant', 'WHfB', whfb),
      ];
    }

    return {
      allUsers: {
        description: 'Strongest authentication method registered by all users.',
        nodes: bucket(items),
      },
      privilegedUsers: {
        description: 'Strongest authentication method registered by privileged users.',
        nodes: bucket(items.filter(u => u.isAdmin === true)),
      },
      truncated,
    };
  }

  // One pass over interactive sign-ins feeds both the CA/MFA and CA/device sankeys.
  // v1.0 auditLogs/signIns returns interactive sign-ins only, matching the module's
  // `isInteractive == true` SQL filter. Failed sign-ins are filtered client-side.
  async function collectSignInSummaries(days, onProgress) {
    const since = new Date(Date.now() - days * 86400000).toISOString();
    const { items, truncated } = await Api.graphPaged('auditLogs/signIns', {
      query: {
        $filter: `createdDateTime ge ${since}`,
        $select: 'createdDateTime,conditionalAccessStatus,authenticationRequirement,status,deviceDetail',
        $top: 999,
      },
      maxItems: SIGNIN_MAX_ITEMS,
      maxMs: SIGNIN_MAX_MS,
      onPage: n => onProgress?.(`Sign-in logs: ${n} entries fetched`),
    });

    const ok = items.filter(s => (s?.status?.errorCode ?? 0) === 0);
    if (!ok.length) return { caMfa: null, caDevices: null, truncated };

    let minDate = null; let maxDate = null;
    for (const s of ok) {
      const d = new Date(s.createdDateTime);
      if (!Number.isNaN(+d)) {
        if (!minDate || d < minDate) minDate = d;
        if (!maxDate || d > maxDate) maxDate = d;
      }
    }
    const duration = durationLabel(minDate, maxDate);

    const caMfaCount = ok.filter(s => s.conditionalAccessStatus === 'success' && s.authenticationRequirement === 'multiFactorAuthentication').length;
    const caNoMfa = ok.filter(s => s.conditionalAccessStatus === 'success' && s.authenticationRequirement === 'singleFactorAuthentication').length;
    const noCaMfa = ok.filter(s => s.conditionalAccessStatus === 'notApplied' && s.authenticationRequirement === 'multiFactorAuthentication').length;
    const noCaNoMfa = ok.filter(s => s.conditionalAccessStatus === 'notApplied' && s.authenticationRequirement === 'singleFactorAuthentication').length;
    const mfaTotal = caMfaCount + caNoMfa + noCaMfa + noCaNoMfa;

    const caMfa = {
      description: `Over the past ${duration}, ${percentLabel(caMfaCount, mfaTotal)} of sign-ins were protected by conditional access policies enforcing multifactor.`,
      nodes: [
        node('User sign in', 'No CA applied', noCaMfa + noCaNoMfa),
        node('User sign in', 'CA applied', caMfaCount + caNoMfa),
        node('CA applied', 'No MFA', caNoMfa),
        node('CA applied', 'MFA', caMfaCount),
      ],
    };

    const managed = ok.filter(s => s?.deviceDetail?.isManaged === true).length;
    const unmanaged = ok.filter(s => s?.deviceDetail?.isManaged === false).length;
    const compliant = ok.filter(s => s?.deviceDetail?.isManaged === true && s?.deviceDetail?.isCompliant === true).length;
    const nonCompliant = ok.filter(s => s?.deviceDetail?.isManaged === true && s?.deviceDetail?.isCompliant !== true).length;

    const caDevices = {
      description: `Over the past ${duration}, ${percentLabel(compliant, managed + unmanaged)} of sign-ins were from compliant devices.`,
      nodes: [
        node('User sign in', 'Unmanaged', unmanaged),
        node('User sign in', 'Managed', managed),
        node('Managed', 'Non-compliant', nonCompliant),
        node('Managed', 'Compliant', compliant),
      ],
    };

    return { caMfa, caDevices, truncated };
  }

  async function collectDeviceOverview(intuneLicensed, onProgress) {
    const { items: devices } = await Api.graphPaged('devices', {
      query: {
        $select: 'operatingSystem,trustType,isCompliant,deviceOwnership,accountEnabled,isManaged',
        $top: 999,
      },
      maxItems: PAGE_MAX_ITEMS,
      maxMs: PAGE_MAX_MS,
      onPage: n => onProgress?.(`Devices: ${n} processed`),
    });

    const managedEnabled = devices.filter(d => d.accountEnabled && d.isManaged);
    const isMac = d => d.operatingSystem === 'MacMDM' || d.operatingSystem === 'macOS';
    const isWindows = d => d.operatingSystem === 'Windows';
    const isAndroid = d => `${d.operatingSystem || ''}`.startsWith('Android');
    const isIos = d => ['iOS', 'IPhone', 'iPad', 'iPadOS'].includes(d.operatingSystem);

    // Device ownership (managed + enabled only, mirroring the module's SQL)
    const deviceOwnership = {
      corporateCount: managedEnabled.filter(d => d.deviceOwnership === 'Company').length,
      personalCount: managedEnabled.filter(d => d.deviceOwnership === 'Personal').length,
    };

    // Desktop sankey (Windows/macOS with a join type)
    const desktop = devices.filter(d => (isWindows(d) || isMac(d)) && d.trustType != null);
    const win = desktop.filter(isWindows);
    const mac = desktop.filter(isMac);
    const joinOf = { AzureAd: 'Entra joined', ServerAd: 'Entra hybrid joined', Workplace: 'Entra registered' };
    const winByJoin = {};
    for (const trust of Object.keys(joinOf)) {
      const set = win.filter(d => d.trustType === trust);
      winByJoin[joinOf[trust]] = {
        total: set.length,
        compliant: set.filter(d => d.isCompliant === true).length,
        nonCompliant: set.filter(d => d.isCompliant === false).length,
        unmanaged: set.filter(d => d.isCompliant == null).length,
      };
    }
    const macCompliant = mac.filter(d => d.isCompliant === true).length;
    const macNonCompliant = mac.filter(d => d.isCompliant === false).length;
    const macUnmanaged = mac.length - macCompliant - macNonCompliant;

    const desktopNodes = [
      node('Desktop devices', 'Windows', win.length),
      node('Desktop devices', 'macOS', mac.length),
    ];
    for (const [join, stats] of Object.entries(winByJoin)) {
      desktopNodes.push(node('Windows', join, stats.total));
    }
    for (const [join, stats] of Object.entries(winByJoin)) {
      desktopNodes.push(node(join, 'Compliant', stats.compliant));
      desktopNodes.push(node(join, 'Non-compliant', stats.nonCompliant));
      desktopNodes.push(node(join, 'Unmanaged', stats.unmanaged));
    }
    desktopNodes.push(node('macOS', 'Compliant', macCompliant));
    desktopNodes.push(node('macOS', 'Non-compliant', macNonCompliant));
    desktopNodes.push(node('macOS', 'Unmanaged', macUnmanaged));

    const desktopDevicesSummary = {
      description: 'Desktop devices (Windows and macOS) by join type and compliance status.',
      nodes: desktopNodes,
      totalDevices: win.length + mac.length,
      entrajoined: winByJoin['Entra joined'].total,
      entrahybridjoined: winByJoin['Entra hybrid joined'].total,
      // Key name preserved from the module (consumed by the report template as-is).
      entrareigstered: winByJoin['Entra registered'].total,
    };

    // Mobile sankey (non-Windows with a known compliance state)
    const mobile = devices.filter(d => !isWindows(d) && !isMac(d) && d.isCompliant != null && (isAndroid(d) || isIos(d)));
    function mobileGroup(platform, filter) {
      const company = mobile.filter(d => filter(d) && d.deviceOwnership === 'Company');
      const personal = mobile.filter(d => filter(d) && d.deviceOwnership === 'Personal');
      return {
        company: company.length,
        personal: personal.length,
        companyCompliant: company.filter(d => d.isCompliant === true).length,
        companyNonCompliant: company.filter(d => d.isCompliant === false).length,
        personalCompliant: personal.filter(d => d.isCompliant === true).length,
        personalNonCompliant: personal.filter(d => d.isCompliant === false).length,
        total: company.length + personal.length,
      };
    }
    const android = mobileGroup('Android', isAndroid);
    const ios = mobileGroup('iOS', isIos);

    const mobileSummary = {
      description: 'Mobile devices by compliance status.',
      nodes: [
        node('Mobile devices', 'Android', android.total),
        node('Mobile devices', 'iOS', ios.total),
        node('Android', 'Android (Company)', android.company),
        node('Android', 'Android (Personal)', android.personal),
        node('iOS', 'iOS (Company)', ios.company),
        node('iOS', 'iOS (Personal)', ios.personal),
        node('Android (Company)', 'Compliant', android.companyCompliant),
        node('Android (Company)', 'Non-compliant', android.companyNonCompliant),
        node('Android (Personal)', 'Compliant', android.personalCompliant),
        node('Android (Personal)', 'Non-compliant', android.personalNonCompliant),
        node('iOS (Company)', 'Compliant', ios.companyCompliant),
        node('iOS (Company)', 'Non-compliant', ios.companyNonCompliant),
        node('iOS (Personal)', 'Compliant', ios.personalCompliant),
        node('iOS (Personal)', 'Non-compliant', ios.personalNonCompliant),
      ],
      totalDevices: mobile.length,
    };

    // ManagedDevices + DeviceCompliance: prefer Intune APIs, fall back to Entra aggregates.
    let managedDevices = null;
    let deviceCompliance = null;
    if (intuneLicensed) {
      try {
        const overview = await Api.graph('deviceManagement/managedDeviceOverview', { beta: true });
        const os = overview?.deviceOperatingSystemSummary || {};
        const desktopCount = (os.windowsCount || 0) + (os.macOSCount || 0);
        const mobileCount = (os.iosCount || os.iOSCount || 0) + (os.androidCount || 0);
        if (desktopCount + mobileCount > 0) {
          managedDevices = { ...overview, desktopCount, mobileCount, totalCount: desktopCount + mobileCount };
        }
        deviceCompliance = await Api.graph('deviceManagement/deviceCompliancePolicyDeviceStateSummary', { beta: true });
      } catch {
        managedDevices = null;
        deviceCompliance = null;
      }
    }
    if (!managedDevices) {
      const windowsCount = managedEnabled.filter(isWindows).length;
      const macOSCount = managedEnabled.filter(isMac).length;
      const iOSCount = managedEnabled.filter(isIos).length;
      const androidCount = managedEnabled.filter(isAndroid).length;
      const linuxCount = managedEnabled.filter(d => d.operatingSystem === 'Linux').length;
      const desktopCount = windowsCount + macOSCount;
      const mobileCount = iOSCount + androidCount;
      if (desktopCount + mobileCount > 0) {
        managedDevices = {
          deviceOperatingSystemSummary: { windowsCount, macOSCount, iosCount: iOSCount, androidCount, linuxCount },
          enrolledDeviceCount: desktopCount + mobileCount,
          desktopCount,
          mobileCount,
          totalCount: desktopCount + mobileCount,
        };
      }
    }
    if (!deviceCompliance) {
      deviceCompliance = {
        compliantDeviceCount: managedEnabled.filter(d => d.isCompliant === true).length,
        nonCompliantDeviceCount: managedEnabled.filter(d => d.isCompliant === false).length,
        inGracePeriodCount: 0,
        configManagerCount: 0,
        unknownDeviceCount: 0,
        notApplicableDeviceCount: 0,
        remediatedDeviceCount: 0,
        errorDeviceCount: 0,
        conflictDeviceCount: 0,
      };
    }

    return {
      DesktopDevicesSummary: desktopDevicesSummary,
      ManagedDevices: managedDevices,
      MobileSummary: mobileSummary,
      DeviceCompliance: deviceCompliance,
      DeviceOwnership: deviceOwnership,
    };
  }

  function assignmentText(assignments) {
    const kinds = new Set();
    for (const a of assignments || []) {
      const type = `${a?.target?.['@odata.type'] || ''}`;
      if (type.includes('allLicensedUsers')) kinds.add('All users');
      else if (type.includes('allDevices')) kinds.add('All devices');
      else if (type.includes('exclusionGroup')) kinds.add('Excluded groups');
      else if (type.includes('group')) kinds.add('Selected groups');
    }
    return Array.from(kinds).join(', ');
  }

  function scopeText(roleScopeTagIds) {
    const tags = (roleScopeTagIds || []).map(t => (`${t}` === '0' ? 'Default' : `${t}`));
    return tags.join(', ') || 'Default';
  }

  function blockAllow(value) {
    if (value === true || value === 'true') return 'Blocked';
    if (value === false || value === 'false') return 'Allowed';
    return '';
  }

  async function collectWindowsEnrollment() {
    const policies = await Api.graphAll('policies/mobileDeviceManagementPolicies', {
      beta: true,
      query: { $expand: 'includedGroups' },
    });
    const appliesToName = { all: 'All', selected: 'Selected', none: 'None' };
    return (policies || [])
      .sort((a, b) => `${b.appliesTo}`.localeCompare(`${a.appliesTo}`) || `${a.displayName}`.localeCompare(`${b.displayName}`))
      .map(p => ({
        Type: 'MDM',
        PolicyName: p.displayName,
        AppliesTo: appliesToName[p.appliesTo] || p.appliesTo,
        Groups: p.appliesTo === 'selected' && (p.includedGroups || []).length
          ? p.includedGroups.map(g => g.displayName).join(', ')
          : 'Not Applicable',
      }));
  }

  const PLATFORM_NAMES = {
    android: 'Android device administrator',
    androidForWork: 'Android Enterprise (work profile)',
    ios: 'iOS/iPadOS',
    mac: 'macOS',
    windows: 'Windows',
    windowsPhone: 'Windows Phone',
  };

  async function collectEnrollmentRestrictions() {
    const configs = await Api.graphAll('deviceManagement/deviceEnrollmentConfigurations', {
      beta: true,
      query: { $expand: 'assignments' },
    });

    const rows = [];
    const single = (configs || [])
      .filter(c => c.deviceEnrollmentConfigurationType === 'singlePlatformRestriction')
      .sort((a, b) => (b.priority || 0) - (a.priority || 0) || `${a.displayName}`.localeCompare(`${b.displayName}`));
    for (const c of single) {
      const r = c.platformRestriction || {};
      rows.push({
        Platform: PLATFORM_NAMES[c.platformType] || c.platformType || '',
        Priority: c.priority,
        Name: c.displayName,
        MDM: blockAllow(r.platformBlocked),
        MinVer: r.osMinimumVersion || '',
        MaxVer: r.osMaximumVersion || '',
        PersonallyOwned: blockAllow(r.personalDeviceEnrollmentBlocked),
        BlockedManufacturers: (r.blockedManufacturers || []).join(', '),
        Scope: scopeText(c.roleScopeTagIds),
        AssignedTo: assignmentText(c.assignments),
      });
    }

    const defaults = (configs || []).find(c => `${c['@odata.type']}` === '#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration');
    if (defaults) {
      const defaultPlatforms = {
        iosRestriction: 'iOS/iPadOS',
        windowsRestriction: 'Windows',
        androidRestriction: 'Android device administrator',
        macOSRestriction: 'macOS',
        androidForWorkRestriction: 'Android Enterprise (work profile)',
      };
      for (const [prop, display] of Object.entries(defaultPlatforms)) {
        const r = defaults[prop];
        if (!r) continue;
        rows.push({
          Platform: display,
          Priority: 'Default',
          Name: defaults.displayName || 'All users and all devices',
          MDM: blockAllow(r.platformBlocked),
          MinVer: r.osMinimumVersion || '',
          MaxVer: r.osMaximumVersion || '',
          PersonallyOwned: blockAllow(r.personalDeviceEnrollmentBlocked),
          BlockedManufacturers: (r.blockedManufacturers || []).join(', '),
          Scope: scopeText(defaults.roleScopeTagIds),
          AssignedTo: 'All users',
        });
      }
    }

    return rows;
  }

  const COMPLIANCE_PLATFORMS = {
    '#microsoft.graph.androidCompliancePolicy': 'Android device administrator',
    '#microsoft.graph.androidDeviceOwnerCompliancePolicy': 'Android Enterprise (Corp)',
    '#microsoft.graph.androidWorkProfileCompliancePolicy': 'Android Enterprise (Personal)',
    '#microsoft.graph.aospDeviceOwnerCompliancePolicy': 'Android (AOSP)',
    '#microsoft.graph.iosCompliancePolicy': 'iOS/iPadOS',
    '#microsoft.graph.macOSCompliancePolicy': 'macOS',
    '#microsoft.graph.windows10CompliancePolicy': 'Windows 10 and later',
    '#microsoft.graph.windows81CompliancePolicy': 'Windows 8.1 and later',
  };

  function passwordTypeLabel(value) {
    const map = {
      deviceDefault: 'Device default',
      alphanumeric: 'Alphanumeric',
      numeric: 'Numeric',
      numericComplex: 'Numeric complex',
      required: 'Password required',
      alphanumericWithSymbols: 'Alphanumeric with symbols',
      lowSecurityBiometric: 'Low security biometric',
      any: 'Any',
      Any: 'Any',
    };
    return value == null ? '' : (map[value] || `${value}`);
  }

  function threatLevelLabel(value) {
    const map = { unavailable: '', secured: 'Secured', low: 'Low', medium: 'Medium', high: 'High', notSet: '' };
    return value == null ? '' : (map[value] ?? `${value}`);
  }

  function graceDays(policy, actionType) {
    const configs = (policy.scheduledActionsForRule || []).flatMap(r => r.scheduledActionConfigurations || []);
    const action = configs.find(c => c.actionType === actionType);
    if (!action) return '';
    const hours = Number(action.gracePeriodHours);
    if (!Number.isFinite(hours)) return '';
    if (hours === 0) return actionType === 'block' ? 'Immediately' : 0;
    return Math.round(hours / 24);
  }

  async function collectCompliancePolicies() {
    const policies = await Api.graphAll('deviceManagement/deviceCompliancePolicies', {
      beta: true,
      query: { $expand: 'assignments,scheduledActionsForRule($expand=scheduledActionConfigurations)' },
    });

    return (policies || []).map(p => {
      const type = `${p['@odata.type'] || ''}`;
      const platform = COMPLIANCE_PLATFORMS[type] || type.replace('#microsoft.graph.', '');
      const view = {
        Platform: platform,
        PolicyName: p.displayName,
        DefenderForEndPoint: threatLevelLabel(p.advancedThreatProtectionRequiredSecurityLevel),
        MinOsVersion: p.osMinimumVersion || '',
        MaxOsVersion: p.osMaximumVersion || '',
        RequirePswd: p.passwordRequired === true ? 'Yes' : '',
        MinPswdLength: p.passwordMinimumLength ?? '',
        PasswordType: passwordTypeLabel(p.passwordRequiredType ?? p.passwordType),
        PswdExpiryDays: p.passwordExpirationDays ?? '',
        CountOfPreviousPswdToBlock: p.passwordPreviousPasswordBlockCount ?? '',
        RequireEncryption: p.storageRequireEncryption === true ? 'Yes' : '',
        RootedJailbrokenDevices: p.securityBlockJailbrokenDevices === true ? 'Blocked' : '',
        MaxDeviceThreatLevel: threatLevelLabel(p.deviceThreatProtectionRequiredSecurityLevel),
        RequireFirewall: '',
        MaxInactivityMin: p.passwordMinutesOfInactivityBeforeLock ?? '',
        ActionForNoncomplianceDaysPushNotification: graceDays(p, 'pushNotification'),
        ActionForNoncomplianceDaysSendEmail: graceDays(p, 'notification'),
        ActionForNoncomplianceDaysRemoteLock: graceDays(p, 'remoteLock'),
        ActionForNoncomplianceDaysBlock: graceDays(p, 'block'),
        ActionForNoncomplianceDaysRetire: graceDays(p, 'retire'),
        Scope: scopeText(p.roleScopeTagIds),
        IncludedGroups: '',
        ExcludedGroups: '',
      };

      switch (type) {
        case '#microsoft.graph.androidCompliancePolicy':
        case '#microsoft.graph.androidDeviceOwnerCompliancePolicy':
        case '#microsoft.graph.androidWorkProfileCompliancePolicy':
          view.RequireFirewall = 'Not Applicable';
          break;
        case '#microsoft.graph.aospDeviceOwnerCompliancePolicy':
          view.DefenderForEndPoint = 'Not Applicable';
          view.PswdExpiryDays = 'Not Applicable';
          view.CountOfPreviousPswdToBlock = 'Not Applicable';
          view.MaxDeviceThreatLevel = 'Not Applicable';
          view.RequireFirewall = 'Not Applicable';
          break;
        case '#microsoft.graph.iosCompliancePolicy':
          view.RequireFirewall = 'Not Applicable';
          view.RequirePswd = p.passcodeRequired === true ? 'Yes' : '';
          view.MinPswdLength = p.passcodeMinimumLength ?? '';
          view.PswdExpiryDays = p.passcodeExpirationDays ?? '';
          view.MaxInactivityMin = p.passcodeMinutesOfInactivityBeforeLock ?? '';
          view.CountOfPreviousPswdToBlock = p.passcodePreviousPasscodeBlockCount ?? '';
          view.PasswordType = passwordTypeLabel(p.passcodeRequiredType);
          view.RequireEncryption = 'Not Applicable';
          break;
        case '#microsoft.graph.macOSCompliancePolicy':
          view.RootedJailbrokenDevices = 'Not Applicable';
          view.RequireFirewall = p.firewallEnabled === true ? 'Yes' : '';
          break;
        case '#microsoft.graph.windows10CompliancePolicy':
          view.RequireFirewall = p.activeFirewallRequired === true ? 'Yes' : '';
          view.MaxDeviceThreatLevel = 'Not Applicable';
          view.RootedJailbrokenDevices = 'Not Applicable';
          break;
        case '#microsoft.graph.windows81CompliancePolicy':
          view.DefenderForEndPoint = 'Not Applicable';
          view.MaxDeviceThreatLevel = 'Not Applicable';
          view.RootedJailbrokenDevices = 'Not Applicable';
          view.RequireFirewall = 'Not Applicable';
          break;
        default:
          break;
      }
      return view;
    });
  }

  function appGroupTypeLabel(value) {
    const map = {
      allApps: 'All apps',
      allMicrosoftApps: 'All Microsoft apps',
      allCoreMicrosoftApps: 'Core Microsoft apps',
      selectedPublicApps: 'Selected apps',
    };
    return value == null ? '' : (map[value] || `${value}`);
  }

  function labelAllowBlockBlank(blocked) {
    if (blocked === true) return 'Block';
    if (blocked === false) return 'Allow';
    return '';
  }

  function transferLevelLabel(value) {
    const map = { allApps: 'All apps', managedApps: 'Policy managed apps', none: 'None' };
    return value == null ? '' : (map[value] || `${value}`);
  }

  function storageLocationsLabel(locations) {
    const map = {
      oneDriveForBusiness: 'OneDrive for Business',
      sharePoint: 'SharePoint',
      box: 'Box',
      localStorage: 'Local storage',
      photoLibrary: 'Photo library',
    };
    return (locations || []).map(l => map[l] || `${l}`).join(', ');
  }

  function dialerLabel(value) {
    const map = {
      allApps: 'Any dialer app',
      managedApps: 'Policy managed apps',
      customApp: 'A specific dialer app',
      blocked: 'None, do not transfer this data between apps',
    };
    return value == null ? '' : (map[value] || `${value}`);
  }

  function complianceActionLabel(value) {
    const map = { block: 'Block access', wipe: 'Wipe data', warn: 'Warn' };
    return value == null ? '' : (map[value] || `${value}`);
  }

  async function collectAppProtectionPolicies() {
    const [android, ios, windows] = await Promise.all([
      Api.graphAll('deviceAppManagement/androidManagedAppProtections', { beta: true, query: { $expand: 'assignments,apps' } }).catch(() => []),
      Api.graphAll('deviceAppManagement/iosManagedAppProtections', { beta: true, query: { $expand: 'assignments,apps' } }).catch(() => []),
      Api.graphAll('deviceAppManagement/mdmWindowsInformationProtectionPolicies', { beta: true, query: { $expand: 'assignments' } }).catch(() => []),
    ]);

    const all = [
      ...(android || []).map(p => ({ policy: p, platform: 'Android' })),
      ...(ios || []).map(p => ({ policy: p, platform: 'iOS/iPadOS' })),
      ...(windows || []).map(p => ({ policy: p, platform: 'Windows' })),
    ];

    return all.map(({ policy: p, platform }) => ({
      Platform: platform,
      Name: p.displayName,
      AppsPublic: appGroupTypeLabel(p.appGroupType),
      AppsCustom: '',
      BackupOrgDataToICloudOrGoogle: labelAllowBlockBlank(p.dataBackupBlocked),
      SendOrgDataToOtherApps: transferLevelLabel(p.allowedOutboundDataTransferDestinations),
      AppsToExempt: (p.exemptedAppPackages || p.exemptedAppProtocols || []).map(x => x.name || x.value || '').filter(Boolean).join(', '),
      SaveCopiesOfOrgData: labelAllowBlockBlank(p.saveAsBlocked),
      AllowUserToSaveCopiesToSelectedServices: storageLocationsLabel(p.allowedDataStorageLocations),
      DataProtectionTransferTelecommunicationDataTo: dialerLabel(p.dialerRestrictionLevel),
      DataProtectionReceiveDataFromOtherApps: transferLevelLabel(p.allowedInboundDataTransferSources),
      DataProtectionOpenDataIntoOrgDocuments: '',
      DataProtectionAllowUsersToOpenDataFromSelectedServices: '',
      DataProtectionRestrictCutCopyBetweenOtherApps: transferLevelLabel(p.allowedOutboundClipboardSharingLevel),
      DataProtectionCutCopyCharacterLimitForAnyApp: '',
      DataProtectionEncryptOrgData: p.encryptAppData === true ? 'Yes' : '',
      DataProtectionSyncPolicyManagedAppDataWithNativeApps: '',
      DataProtectionPrintingOrgData: labelAllowBlockBlank(p.printBlocked),
      DataProtectionRestrictWebContentTransferWithOtherApps: '',
      DataProtectionOrgDataNotifications: '',
      ConditionalLaunchAppMaxPinAttempts: p.maximumPinRetries ?? '',
      ConditionalLaunchAppOfflineGracePeriodBlockAccess: p.periodOfflineBeforeAccessCheck || '',
      ConditionalLaunchAppOfflineGracePeriodWipeData: p.periodOfflineBeforeWipeIsEnforced || '',
      ConditionalLaunchAppDisabedAccount: '',
      ConditionalLaunchAppMinAppVersion: p.minimumRequiredAppVersion || '',
      ConditionalLaunchDeviceRootedJailbrokenDevices: complianceActionLabel(p.appActionIfDeviceComplianceRequired),
      ConditionalLaunchDevicePrimaryMtdService: '',
      ConditionalLaunchDeviceMaxAllowedDeviceThreatLevel: '',
      ConditionalLaunchDeviceMinOsVersion: p.minimumRequiredOsVersion || '',
      ConditionalLaunchDeviceMaxOsVersion: '',
      Scope: scopeText(p.roleScopeTagIds),
      IncludedGroups: '',
      ExcludedGroups: '',
    }));
  }

  // Runs all collectors sequentially (throttle-friendly). Returns { tenantInfo, notes }:
  // tenantInfo uses the module's exact key names; notes explains every null/absent key.
  async function collect(options) {
    const opts = options || {};
    const log = opts.log || (() => {});
    const progress = opts.progress || (() => {});
    const days = Math.min(30, Math.max(1, Number(opts.days) || SIGNIN_DEFAULT_DAYS));

    const tenantInfo = {};
    const notes = {};

    async function step(key, label, fn) {
      progress(`Tenant insights: ${label}`);
      try {
        const value = await fn();
        tenantInfo[key] = value ?? null;
        if (value == null) notes[key] = notes[key] || 'No data returned.';
      } catch (e) {
        tenantInfo[key] = null;
        notes[key] = e?.status === 403
          ? 'Permission or license missing for this dataset.'
          : `Collection failed: ${e?.message || e}`;
        log(`Tenant insights: ${label} failed: ${e?.message || e}`);
      }
    }

    const licenses = await getLicenses();

    await step('TenantOverview', 'tenant overview', collectTenantOverview);

    if (!licenses.entraPremium) {
      const reason = 'Requires Entra ID P1 or higher.';
      tenantInfo.OverviewAuthMethodsAllUsers = null;
      tenantInfo.OverviewAuthMethodsPrivilegedUsers = null;
      tenantInfo.OverviewCaMfaAllUsers = null;
      tenantInfo.OverviewCaDevicesAllUsers = null;
      notes.OverviewAuthMethodsAllUsers = reason;
      notes.OverviewAuthMethodsPrivilegedUsers = reason;
      notes.OverviewCaMfaAllUsers = reason;
      notes.OverviewCaDevicesAllUsers = reason;
    } else {
      let authTruncated = false;
      await step('OverviewAuthMethodsAllUsers', 'authentication methods', async () => {
        const res = await collectAuthMethods(progress);
        tenantInfo.OverviewAuthMethodsPrivilegedUsers = res.privilegedUsers;
        authTruncated = res.truncated;
        return res.allUsers;
      });
      if (authTruncated) notes.OverviewAuthMethodsAllUsers = 'Sampled: user registration paging hit the browser collection limit.';

      await step('OverviewCaMfaAllUsers', `sign-in logs (${days} days)`, async () => {
        const res = await collectSignInSummaries(days, progress);
        tenantInfo.OverviewCaDevicesAllUsers = res.caDevices;
        if (res.truncated) {
          notes.OverviewCaMfaAllUsers = 'Sampled: sign-in log paging hit the browser collection limit; percentages reflect the sampled window.';
          notes.OverviewCaDevicesAllUsers = notes.OverviewCaMfaAllUsers;
        }
        return res.caMfa;
      });
    }

    if (!licenses.intune) {
      const reason = 'Requires a Microsoft Intune license.';
      for (const key of ['DeviceOverview', 'ConfigWindowsEnrollment', 'ConfigDeviceEnrollmentRestriction', 'ConfigDeviceCompliancePolicies', 'ConfigDeviceAppProtectionPolicies']) {
        tenantInfo[key] = null;
        notes[key] = reason;
      }
    } else {
      await step('DeviceOverview', 'device overview', () => collectDeviceOverview(licenses.intune, progress));
      await step('ConfigWindowsEnrollment', 'Windows enrollment policies', collectWindowsEnrollment);
      await step('ConfigDeviceEnrollmentRestriction', 'enrollment restrictions', collectEnrollmentRestrictions);
      await step('ConfigDeviceCompliancePolicies', 'compliance policies', collectCompliancePolicies);
      await step('ConfigDeviceAppProtectionPolicies', 'app protection policies', collectAppProtectionPolicies);
    }

    return { tenantInfo, notes, days };
  }

  window.TenantInfo = { collect };
})();
