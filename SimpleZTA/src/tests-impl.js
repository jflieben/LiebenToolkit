(() => {
  function result(status, test, message, evidence) {
    return { id: test.id, title: test.title, status, message, evidence: evidence || null };
  }

  function pass(test, message, evidence) {
    return result('Passed', test, message, evidence);
  }

  function fail(test, message, evidence) {
    return result('Failed', test, message, evidence);
  }

  function skip(test, message, evidence) {
    return result('Skipped', test, message, evidence);
  }

  function daysBetween(a, b) {
    return Math.floor((new Date(a).getTime() - new Date(b).getTime()) / 86400000);
  }

  function toArray(value) {
    if (Array.isArray(value)) return value;
    if (value == null) return [];
    return [value];
  }

  function lower(value) {
    return `${value || ''}`.toLowerCase();
  }

  function uniqueBy(items, keySelector) {
    const seen = new Set();
    const out = [];
    for (const item of items || []) {
      const key = keySelector(item);
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(item);
    }
    return out;
  }

  function cloneWithoutKeys(obj, keysToDrop) {
    const src = obj || {};
    const out = {};
    const drop = new Set(toArray(keysToDrop));
    for (const [key, value] of Object.entries(src)) {
      if (!drop.has(key)) out[key] = value;
    }
    return out;
  }

  function isRetriableBadRequest(err) {
    return err?.status === 400;
  }

  async function graphAllWithFallback(path, attempts) {
    let lastError = null;
    for (let i = 0; i < attempts.length; i++) {
      const attempt = attempts[i] || {};
      try {
        return await Api.graphAll(path, attempt);
      } catch (err) {
        lastError = err;
        if (!isRetriableBadRequest(err) || i === attempts.length - 1) throw err;
      }
    }
    throw lastError || new Error('Unknown Graph failure');
  }

  async function graphWithFallback(path, attempts) {
    let lastError = null;
    for (let i = 0; i < attempts.length; i++) {
      const attempt = attempts[i] || {};
      try {
        return await Api.graph(path, attempt);
      } catch (err) {
        lastError = err;
        if (!isRetriableBadRequest(err) || i === attempts.length - 1) throw err;
      }
    }
    throw lastError || new Error('Unknown Graph failure');
  }

  function intersects(left, right) {
    const rightSet = new Set(toArray(right).map(x => `${x}`));
    return toArray(left).some(x => rightSet.has(`${x}`));
  }

  function isTimeoutError(err) {
    const status = Number(err?.status || 0);
    if ([408, 504, 524].includes(status)) return true;
    const msg = `${err?.message || ''}`.toLowerCase();
    return msg.includes('timeout') || msg.includes('timed out') || msg.includes('aborted');
  }

  function includesAllUsers(policy) {
    const users = toArray(policy?.conditions?.users?.includeUsers);
    return users.includes('All');
  }

  function includesRoles(policy, roleIds) {
    return intersects(policy?.conditions?.users?.includeRoles, roleIds);
  }

  function includesUserAction(policy, action) {
    return toArray(policy?.conditions?.applications?.includeUserActions).includes(action);
  }

  function hasGrantControl(policy, control) {
    return toArray(policy?.grantControls?.builtInControls).includes(control);
  }

  function getAuthStrengthId(policy) {
    return policy?.grantControls?.authenticationStrength?.id || null;
  }

  function hasAnyEnabledAuthMethods(authPolicy) {
    return toArray(authPolicy?.authenticationMethodConfigurations).some(method => lower(method?.state) === 'enabled');
  }

  function featureSettingEnabledForAllUsers(setting) {
    return lower(setting?.state) === 'enabled' && lower(setting?.includeTarget?.id) === 'all_users';
  }

  function includeTargetsAllUsers(targets) {
    return toArray(targets).some(target => lower(target?.id || target?.target) === 'all_users');
  }

  function summarizeTargets(targets) {
    const values = toArray(targets).map(target => target?.displayName || target?.id || target?.target || '').filter(Boolean);
    return values.length ? values.join(', ') : 'None';
  }

  function isPhishingResistantCombination(combo) {
    const allowed = new Set(['windowshelloforbusiness', 'fido2', 'x509certificatemultifactor']);
    const parts = `${combo || ''}`.split(',').map(x => x.trim().toLowerCase()).filter(Boolean);
    return parts.length > 0 && parts.every(part => allowed.has(part));
  }

  async function getAuthorizationPolicy() {
    const payload = await Api.graph('policies/authorizationPolicy', { beta: true });
    if (Array.isArray(payload?.value)) return payload.value[0] || null;
    return payload || null;
  }

  async function getEntraDiagnosticSettings() {
    const armScope = ['https://management.azure.com/user_impersonation'];
    const url = 'https://management.azure.com/providers/microsoft.aadiam/diagnosticsettings?api-version=2017-04-01-preview';

    try {
      const token = await Auth.getTokenSilent(armScope);
      if (!token) return null;

      const response = await fetch(url, {
        method: 'GET',
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      const text = response.status === 204 ? '' : await response.text();
      const parsed = text ? JSON.parse(text) : {};
      if (!response.ok) {
        const err = new Error(`ARM ${response.status}: ${text}`);
        err.status = response.status;
        throw err;
      }

      return toArray(parsed?.value);
    } catch (err) {
      if (err.status === 400 || err.status === 401 || err.status === 403 || err.status === 404) return null;
      if ((err?.message || '').includes('Interaction required')) return null;
      throw err;
    }
  }

  async function getAuthenticationMethodsPolicy() {
    return Api.graph('policies/authenticationMethodsPolicy', { beta: true });
  }

  async function getAuthenticatorConfig() {
    return Api.graph('policies/authenticationMethodsPolicy/authenticationMethodConfigurations/MicrosoftAuthenticator', { beta: true });
  }

  async function getFido2Config() {
    return Api.graph('policies/authenticationMethodsPolicy/authenticationMethodConfigurations/Fido2', { beta: true });
  }

  async function getTapConfig() {
    return Api.graph('policies/authenticationMethodsPolicy/authenticationMethodConfigurations/temporaryAccessPass', { beta: true });
  }

  async function getConditionalAccessPolicies() {
    return Api.graphAll('identity/conditionalAccess/policies', { beta: true, query: { '$top': '500' } });
  }

  async function getAuthStrengthPolicies() {
    return Api.graphAll('policies/authenticationStrengthPolicies', { beta: true });
  }

  async function getPrivilegedBuiltInRoles() {
    const roles = await Api.graphAll('roleManagement/directory/roleDefinitions', {
      beta: true,
      query: { '$select': 'id,displayName,isBuiltIn,isPrivileged' },
    });
    return roles.filter(role => role.isBuiltIn && role.isPrivileged);
  }

  async function getRoleAssignments() {
    return Api.graphAll('roleManagement/directory/roleAssignments', {
      beta: true,
      query: { '$select': 'id,principalId,roleDefinitionId,directoryScopeId', '$top': '999' },
    });
  }

  async function getServicePrincipals() {
    return Api.graphAll('servicePrincipals', {
      beta: true,
      query: {
        '$select': 'id,appId,displayName,servicePrincipalType,signInActivity,appOwnerOrganizationId,passwordCredentials,keyCredentials',
        '$top': '999',
      },
    });
  }

  async function getUsersByIds(ids) {
    const uniqueIds = uniqueBy(toArray(ids).filter(Boolean), id => `${id}`).map(id => `${id}`);
    if (!uniqueIds.length) return [];
    const payload = await Api.graph('directoryObjects/getByIds', {
      method: 'POST',
      body: { ids: uniqueIds, types: ['user'] },
    });
    return toArray(payload?.value).filter(item => lower(item?.['@odata.type']).includes('user'));
  }

  async function getUserAuthMethods(userId) {
    return Api.graphAll(`users/${userId}/authentication/methods`, { beta: true, query: { '$top': '50' } });
  }

  async function getCrossTenantDefaultPolicy() {
    return Api.graph('policies/crossTenantAccessPolicy/default', { beta: true });
  }

  async function getLegacyPolicies() {
    try {
      const payload = await Api.graph('legacy/policies', { beta: true });
      return Array.isArray(payload?.value) ? payload.value : toArray(payload);
    } catch (err) {
      if (err.status === 403 || err.status === 404) return [];
      throw err;
    }
  }

  async function getNetworkCrossTenantAccessSettings() {
    try {
      return await Api.graph('networkAccess/settings/crossTenantAccess', { beta: true });
    } catch (err) {
      if (err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getDeviceRegistrationPolicy() {
    return Api.graph('policies/deviceRegistrationPolicy', { beta: false });
  }

  async function getRiskyUsers(filter) {
    try {
      return await Api.graphAll('identityProtection/riskyUsers', {
        query: { '$filter': filter, '$top': '200' },
      });
    } catch (err) {
      if (err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getRiskyServicePrincipals() {
    try {
      return await Api.graphAll('identityProtection/riskyServicePrincipals', {
        query: { '$filter': "riskState eq 'atRisk'", '$top': '200' },
      });
    } catch (err) {
      if (err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getRiskDetections(filter) {
    try {
      return await Api.graphAll('identityProtection/riskDetections', {
        beta: true,
        query: { '$filter': filter, '$top': '200' },
      });
    } catch (err) {
      if (err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getServicePrincipalRiskDetections() {
    try {
      return await Api.graphAll('identityProtection/servicePrincipalRiskDetections', {
        beta: true,
        query: { '$top': '200' },
      });
    } catch (err) {
      if (err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getAuthContextClassReferences() {
    return Api.graphAll('identity/conditionalAccess/authenticationContextClassReferences', {
      beta: true,
      query: {},
    });
  }

  async function getNamedLocations() {
    return Api.graphAll('identity/conditionalAccess/namedLocations', {
      beta: true,
      query: { '$top': '200' },
    });
  }

  async function getOrganization() {
    const payload = await Api.graph('organization', { query: { '$select': 'id,displayName,onPremisesSyncEnabled', '$top': '1' } });
    return toArray(payload?.value)[0] || null;
  }

  async function getAdminConsentRequestPolicy() {
    return Api.graph('policies/adminConsentRequestPolicy');
  }

  async function getDirectorySettings() {
    try {
      return await Api.graphAll('settings', { beta: true, query: { '$top': '100' } });
    } catch (err) {
      if (err.status === 403 || err.status === 404) return [];
      throw err;
    }
  }

  async function getGuestUsers() {
    return Api.graphAll('users', {
      beta: true,
      query: {
        '$filter': "userType eq 'Guest'",
        '$select': 'id,displayName,userPrincipalName,signInActivity,externalUserState,accountEnabled',
        '$top': '500',
      },
    });
  }

  async function getRoleAssignmentScheduleInstances(filter) {
    try {
      const q = { '$select': 'id,principalId,roleDefinitionId,assignmentType,memberType', '$top': '200' };
      if (filter) q['$filter'] = filter;
      return await Api.graphAll('roleManagement/directory/roleAssignmentScheduleInstances', { beta: true, query: q });
    } catch (err) {
      if (err.status === 403 || err.status === 404 || err.status === 429) return null;
      throw err;
    }
  }

  async function getRoleEligibilitySchedules() {
    try {
      return await Api.graphAll('roleManagement/directory/roleEligibilitySchedules', {
        beta: true,
        query: { '$select': 'id,principalId,roleDefinitionId,status,scheduleInfo', '$top': '200' },
      });
    } catch (err) {
      if (err.status === 403 || err.status === 404 || err.status === 429) return null;
      throw err;
    }
  }

  async function getEntraRecommendations() {
    try {
      return await Api.graphAll('directory/recommendations', {
        beta: true,
        query: { '$select': 'id,displayName,priority,status,recommendationType', '$top': '50' },
      });
    } catch (err) {
      if (err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getAccessReviewDefinitions() {
    try {
      return await Api.graphAll('identityGovernance/accessReviews/definitions', {
        beta: true,
        query: { '$select': 'id,displayName,status,scope,instanceEnumerationScope', '$top': '100' },
      });
    } catch (err) {
      if (err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getDeviceCompliancePolicies() {
    try {
      return await graphAllWithFallback('deviceManagement/deviceCompliancePolicies', [
        { query: { '$top': '100' } },
        { beta: true, query: { '$top': '100' } },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getManagedAppPolicies() {
    try {
      return await graphAllWithFallback('deviceAppManagement/managedAppPolicies', [
        { query: { '$top': '100' } },
        { beta: true, query: { '$top': '100' } },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getConfigurationPolicies() {
    try {
      return await graphAllWithFallback('deviceManagement/configurationPolicies', [
        { beta: true, query: { '$select': 'id,name,platforms,technologies', '$top': '100' } },
        { beta: true, query: { '$top': '100' } },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getDeviceConfigurations() {
    try {
      return await graphAllWithFallback('deviceManagement/deviceConfigurations', [
        { query: { '$top': '100' } },
        { beta: true, query: { '$top': '100' } },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getServicePrincipalOwners(spId) {
    try {
      return await Api.graphAll(`servicePrincipals/${spId}/owners`, {
        query: { '$select': 'id,displayName,userType', '$top': '10' },
      });
    } catch {
      return [];
    }
  }

  async function getSignInLogs(filter, top, options) {
    const opts = options || {};
    try {
      const q = {
        '$top': `${top || 25}`,
        '$select': 'id,createdDateTime,userPrincipalName,appDisplayName,clientAppUsed,authenticationRequirement,authenticationLibrary,status',
      };
      if (filter) q['$filter'] = filter;
      const payload = await graphWithFallback('auditLogs/signIns', [
        { beta: true, query: q, timeoutMs: opts.timeoutMs },
        { beta: true, query: cloneWithoutKeys(q, ['$select']), timeoutMs: opts.timeoutMs },
        { beta: true, query: cloneWithoutKeys(q, ['$filter']), timeoutMs: opts.timeoutMs },
      ]);
      return toArray(payload?.value);
    } catch (err) {
      if (isTimeoutError(err)) throw err;
      if (err.status === 400 || err.status === 403 || err.status === 404 || err.status === 429) return null;
      throw err;
    }
  }

  async function getArmRootRoleAssignments() {
    const armScope = ['https://management.azure.com/user_impersonation'];
    let token = null;

    try {
      token = await Auth.getTokenSilent(armScope);
      if (!token) {
        // If scope is missing, attempt interactive consent request.
        token = await Auth.getToken(armScope);
      }
      if (!token) return { assignments: null, scopePrompted: true };

      const response = await fetch('https://management.azure.com/providers/Microsoft.Authorization/roleAssignments?$filter=atScope()&api-version=2022-04-01', {
        method: 'GET',
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      const text = response.status === 204 ? '' : await response.text();
      let parsed = {};
      if (text) {
        try {
          parsed = JSON.parse(text);
        } catch {
          parsed = {};
        }
      }

      if (!response.ok) {
        const err = new Error(`ARM ${response.status}: ${text}`);
        err.status = response.status;
        throw err;
      }

      return { assignments: toArray(parsed?.value), scopePrompted: false };
    } catch (err) {
      if (err.status === 400 || err.status === 401 || err.status === 403 || err.status === 404) {
        return { assignments: null, scopePrompted: false };
      }
      if ((err?.message || '').includes('Interaction required')) {
        return { assignments: null, scopePrompted: true };
      }
      throw err;
    }
  }

  async function getDirectoryAuditLogs(filter, top) {
    try {
      const q = { '$top': `${top || 25}`, '$select': 'id,activityDisplayName,activityDateTime,initiatedBy,targetResources,result' };
      if (filter) q['$filter'] = filter;
      return await graphAllWithFallback('auditLogs/directoryAudits', [
        { query: q },
        { query: cloneWithoutKeys(q, ['$select']) },
        { query: cloneWithoutKeys(q, ['$select', '$filter']) },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404 || err.status === 429) return null;
      throw err;
    }
  }

  async function getSelfServicePasswordResetPolicy() {
    try {
      const [authMethodsPolicy, authorizationPolicy] = await Promise.all([
        Api.graph('policies/authenticationMethodsPolicy', { beta: true }),
        getAuthorizationPolicy(),
      ]);
      return {
        isEnabled: authorizationPolicy?.allowedToUseSspr !== false,
        allowedToUseSspr: authorizationPolicy?.allowedToUseSspr,
        registrationEnforcement: authMethodsPolicy?.registrationEnforcement || null,
      };
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getUserRegistrationDetails(top) {
    try {
      const q = { '$select': 'id,userPrincipalName,isMfaRegistered,isMfaCapable,methodsRegistered,isAdmin', '$top': `${top || 200}` };
      return await graphAllWithFallback('reports/authenticationMethods/userRegistrationDetails', [
        { beta: true, query: q },
        { beta: true, query: cloneWithoutKeys(q, ['$select']) },
        { query: cloneWithoutKeys(q, ['$select']) },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getIdentityProtectionMfaPolicy() {
    // Graph does not expose identityProtection/policies/mfaRegistration.
    // Keep behavior aligned with the existing security-defaults fallback in test 21893.
    return null;
  }

  async function getIdentityProtectionNotificationPolicy() {
    // Graph does not expose Identity Protection notification recipient/settings APIs.
    return null;
  }

  async function getRoleManagementPolicyAssignments(filter, top) {
    try {
      const q = { '$top': `${top || 20}` };
      if (filter) q['$filter'] = filter;
      return await Api.graphAll('policies/roleManagementPolicyAssignments', {
        beta: true,
        query: q,
      });
    } catch (err) {
      if (isTimeoutError(err)) throw err;
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getRoleManagementPolicyRule(policyId, ruleId) {
    try {
      return await Api.graph(`policies/roleManagementPolicies/${policyId}/rules/${ruleId}`, {
        beta: true,
      });
    } catch (err) {
      if (isTimeoutError(err)) throw err;
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getRoleManagementAlerts() {
    try {
      const q = { '$select': 'id,alertDefinitionId,scopeId,isActive,alertDefinition', '$expand': 'alertDefinition', '$top': '50' };
      return await graphAllWithFallback('identityGovernance/roleManagementAlerts/alerts', [
        { beta: true, query: q },
        { beta: true, query: cloneWithoutKeys(q, ['$select']) },
        { beta: true, query: { '$top': '50' } },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404 || err.status === 429) return null;
      throw err;
    }
  }

  async function getEntitlementConnectedOrgs() {
    try {
      const q = { '$select': 'id,displayName,state,identitySources', '$top': '100' };
      return await graphAllWithFallback('identityGovernance/entitlementManagement/connectedOrganizations', [
        { query: q },
        { beta: true, query: q },
        { beta: true, query: cloneWithoutKeys(q, ['$select']) },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getEntitlementAssignmentPolicies() {
    try {
      const q = { '$select': 'id,displayName,allowedTargetScope,expiration,requestApprovalSettings,requestorSettings', '$top': '100' };
      return await graphAllWithFallback('identityGovernance/entitlementManagement/assignmentPolicies', [
        { query: q },
        { beta: true, query: q },
        { beta: true, query: cloneWithoutKeys(q, ['$select']) },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getEntitlementAccessPackages() {
    try {
      const q = { '$select': 'id,displayName,isHidden', '$top': '100' };
      return await graphAllWithFallback('identityGovernance/entitlementManagement/accessPackages', [
        { query: q },
        { beta: true, query: q },
        { beta: true, query: cloneWithoutKeys(q, ['$select']) },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getDeviceTermsAndConditions() {
    try {
      const q = { '$select': 'id,displayName', '$top': '50' };
      return await graphAllWithFallback('deviceManagement/termsAndConditions', [
        { query: q },
        { beta: true, query: q },
        { beta: true, query: cloneWithoutKeys(q, ['$select']) },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getTermsAndConditionsWithAssignments() {
    const policies = await getDeviceTermsAndConditions();
    if (policies === null) return null;
    const out = [];
    for (const policy of toArray(policies)) {
      let assignments = [];
      try {
        assignments = await graphAllWithFallback(`deviceManagement/termsAndConditions/${policy.id}/assignments`, [
          { beta: true, query: { '$top': '100' } },
          { beta: true },
          { query: { '$top': '100' } },
        ]);
      } catch (err) {
        if (err.status !== 400 && err.status !== 403 && err.status !== 404) throw err;
      }
      out.push({ ...policy, assignments });
    }
    return out;
  }

  async function getEnrollmentNotificationConfigurations() {
    try {
      return await graphAllWithFallback('deviceManagement/deviceEnrollmentConfigurations', [
        { beta: true, query: { '$expand': 'assignments', '$filter': "deviceEnrollmentConfigurationType eq 'EnrollmentNotificationsConfiguration'" } },
        { beta: true, query: { '$expand': 'assignments', '$top': '100' } },
        { beta: true, query: { '$top': '100' } },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  let graphPermissionRiskRowsPromise = null;
  async function getGraphPermissionRiskRows() {
    if (!graphPermissionRiskRowsPromise) {
      graphPermissionRiskRowsPromise = fetch('./powershell/assets/aadconsentgrantpermissiontable.csv', { cache: 'no-store' })
        .then(r => r.ok ? r.text() : '')
        .then(text => {
          const lines = `${text || ''}`.split(/\r?\n/).filter(Boolean);
          if (lines.length <= 1) return [];
          const rows = [];
          for (let i = 1; i < lines.length; i++) {
            const parts = lines[i].split(',');
            if (parts.length < 3) continue;
            rows.push({
              type: `${parts[0] || ''}`.trim(),
              permission: `${parts[1] || ''}`.trim(),
              privilege: `${parts[2] || ''}`.trim(),
            });
          }
          return rows;
        })
        .catch(() => []);
    }
    return graphPermissionRiskRowsPromise;
  }

  async function getGraphPermissionMaps() {
    const graphSp = await Api.graphAll('servicePrincipals', {
      query: {
        '$filter': "appId eq '00000003-0000-0000-c000-000000000000'",
        '$select': 'id,appRoles,oauth2PermissionScopes',
        '$top': '1',
      },
    });
    const first = toArray(graphSp)[0] || null;
    const delegatedById = new Map();
    const appById = new Map();
    for (const scope of toArray(first?.oauth2PermissionScopes)) {
      if (scope?.id && scope?.value) delegatedById.set(lower(scope.id), `${scope.value}`);
    }
    for (const role of toArray(first?.appRoles)) {
      if (role?.id && role?.value) appById.set(lower(role.id), `${role.value}`);
    }
    return { delegatedById, appById };
  }

  function permissionRisk(rows, type, permission) {
    const normalizedType = `${type || ''}`.toLowerCase();
    const normalizedPermission = `${permission || ''}`.trim();
    const exact = rows.find(r => lower(r.type) === normalizedType && `${r.permission}` === normalizedPermission);
    if (exact) return exact.privilege || 'Unranked';
    const root = normalizedPermission.includes('.') ? normalizedPermission.split('.')[0] : normalizedPermission;
    const rootRow = rows.find(r => lower(r.type) === normalizedType && `${r.permission}` === root);
    if (rootRow) return rootRow.privilege || 'Unranked';
    if (normalizedType === 'application') {
      if (normalizedPermission.includes('Write')) return 'High';
      return 'Medium';
    }
    return 'Unranked';
  }

  function graphRisk(delegatePermissions, applicationPermissions, riskRows) {
    let finalRisk = 'Unranked';
    for (const permission of toArray(applicationPermissions)) {
      const risk = permissionRisk(riskRows, 'Application', permission);
      if (risk === 'High') return 'High';
      if (risk === 'Medium') finalRisk = 'Medium';
      if (risk === 'Low' && finalRisk === 'Unranked') finalRisk = 'Low';
    }
    for (const permission of toArray(delegatePermissions)) {
      const risk = permissionRisk(riskRows, 'Delegated', permission);
      if (risk === 'High') return 'High';
      if (risk === 'Medium') finalRisk = 'Medium';
      if (risk === 'Low' && finalRisk === 'Unranked') finalRisk = 'Low';
    }
    return finalRisk;
  }

  async function getApplicationsWithPermissionsFor21770() {
    const [apps, sps, permMaps, riskRows] = await Promise.all([
      Api.graphAll('applications', {
        query: {
          '$select': 'id,appId,displayName,publisherDomain,requiredResourceAccess',
          '$top': '999',
        },
      }),
      Api.graphAll('servicePrincipals', {
        beta: true,
        query: {
          '$filter': "servicePrincipalType eq 'Application'",
          '$select': 'id,appId,displayName,appOwnerOrganizationId,signInActivity',
          '$top': '999',
        },
      }),
      getGraphPermissionMaps(),
      getGraphPermissionRiskRows(),
    ]);

    const spByAppId = new Map(toArray(sps).map(sp => [lower(sp.appId), sp]));
    const out = [];
    for (const app of toArray(apps)) {
      const graphReq = toArray(app.requiredResourceAccess).find(r => lower(r.resourceAppId) === '00000003-0000-0000-c000-000000000000');
      if (!graphReq) continue;

      const delegatePermissions = [];
      const appPermissions = [];
      for (const ra of toArray(graphReq.resourceAccess)) {
        const id = lower(ra?.id);
        if (!id) continue;
        if (lower(ra?.type) === 'scope') {
          const name = permMaps.delegatedById.get(id);
          if (name) delegatePermissions.push(name);
        } else {
          const name = permMaps.appById.get(id);
          if (name) appPermissions.push(name);
        }
      }

      // Match original intent: only applications with actual Graph permission entries.
      if (!delegatePermissions.length && !appPermissions.length) continue;

      const risk = graphRisk(delegatePermissions, appPermissions, riskRows);
      const sp = spByAppId.get(lower(app.appId));
      out.push({
        id: sp?.id || app.id,
        appId: app.appId,
        displayName: app.displayName,
        publisherName: app.publisherDomain || sp?.appOwnerOrganizationId || '',
        lastSignInDateTime: sp?.signInActivity?.lastSignInDateTime || null,
        DelegatePermissions: delegatePermissions,
        AppPermissions: appPermissions,
        Risk: risk,
        IsRisky: risk === 'High',
      });
    }
    return out;
  }

  async function getDeviceEnrollmentConfigurations() {
    try {
      return await graphAllWithFallback('deviceManagement/deviceEnrollmentConfigurations', [
        { query: { '$top': '100' } },
        { beta: true, query: { '$top': '100' } },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getDeviceRoleScopeTags() {
    try {
      return await graphAllWithFallback('deviceManagement/roleScopeTags', [
        { beta: true, query: { '$select': 'id,displayName', '$top': '50' } },
        { beta: true, query: { '$top': '50' } },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getDeviceCleanupRules() {
    try {
      const q = { '$select': 'id,displayName,deviceInactivityBeforeRetirementInDays', '$top': '20' };
      return await graphAllWithFallback('deviceManagement/managedDeviceCleanupRules', [
        { beta: true, query: q },
        { beta: true, query: cloneWithoutKeys(q, ['$select']) },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  async function getDeviceIntents() {
    try {
      const q = { '$select': 'id,displayName,templateId,description', '$top': '100' };
      return await graphAllWithFallback('deviceManagement/intents', [
        { beta: true, query: q },
        { beta: true, query: cloneWithoutKeys(q, ['$select']) },
      ]);
    } catch (err) {
      if (err.status === 400 || err.status === 403 || err.status === 404) return null;
      throw err;
    }
  }

  function hasPhishingResistantRegisteredMethod(methods) {
    return toArray(methods).some(method => {
      const type = lower(method?.['@odata.type']);
      return type.includes('fido2') || type.includes('windowshelloforbusiness') || type.includes('x509certificate');
    });
  }

  const impl = {
    '21770': async (test) => {
      const apps = await getApplicationsWithPermissionsFor21770();
      if (!apps.length) return pass(test, 'No applications with Graph permissions were found in the current query scope.');

      const filtered = apps.filter(a => ['High', 'Unranked'].includes(`${a.Risk}`));
      const inactiveRisky = [];
      const other = [];
      for (const item of filtered) {
        if (!item.lastSignInDateTime && item.IsRisky) inactiveRisky.push(item);
        else other.push(item);
      }

      if (!inactiveRisky.length) return pass(test, 'No inactive applications with high-risk Graph permissions were found.', {
        checked: filtered.length,
        highRiskInactive: 0,
      });

      return fail(test, `Found ${inactiveRisky.length} inactive applications with high-risk Graph permissions.`, {
        highRiskInactive: inactiveRisky.slice(0, 40).map(a => ({
          id: a.id,
          appId: a.appId,
          displayName: a.displayName,
          risk: a.Risk,
          delegatePermissions: a.DelegatePermissions,
          applicationPermissions: a.AppPermissions,
          lastSignInDateTime: a.lastSignInDateTime,
        })),
        other: other.slice(0, 40).map(a => ({
          id: a.id,
          appId: a.appId,
          displayName: a.displayName,
          risk: a.Risk,
          lastSignInDateTime: a.lastSignInDateTime,
        })),
      });
    },

    '21771': async (test) => {
      const [roles, assignments, servicePrincipals] = await Promise.all([
        getPrivilegedBuiltInRoles(),
        getRoleAssignments(),
        getServicePrincipals(),
      ]);

      const roleMap = new Map(roles.map(role => [role.id, role.displayName]));
      const spMap = new Map(servicePrincipals.map(sp => [sp.id, sp]));
      const matches = assignments
        .filter(assignment => roleMap.has(assignment.roleDefinitionId) && spMap.has(assignment.principalId))
        .map(assignment => {
          const sp = spMap.get(assignment.principalId);
          return {
            principalId: assignment.principalId,
            displayName: sp.displayName,
            appId: sp.appId,
            roleName: roleMap.get(assignment.roleDefinitionId),
            lastSignInDateTime: sp.signInActivity?.lastSignInDateTime || null,
          };
        });

      const inactive = matches.filter(item => !item.lastSignInDateTime);
      if (!inactive.length) return pass(test, 'No inactive service principals were found with privileged built-in roles.', { assignmentsChecked: matches.length });
      return fail(test, `Found ${inactive.length} inactive service principals with privileged built-in roles.`, inactive.slice(0, 30));
    },

    '21772': async (test) => {
      const apps = await Api.graphAll('applications', {
        query: { '$select': 'id,appId,displayName,passwordCredentials', '$top': '999' },
      });
      const sps = await Api.graphAll('servicePrincipals', {
        query: { '$select': 'id,appId,displayName,passwordCredentials,appOwnerOrganizationId', '$top': '999' },
      });

      const appSecrets = apps.filter(app => Array.isArray(app.passwordCredentials) && app.passwordCredentials.length > 0);
      const spSecrets = sps.filter(sp => Array.isArray(sp.passwordCredentials) && sp.passwordCredentials.length > 0);
      const total = appSecrets.length + spSecrets.length;
      if (total === 0) return pass(test, 'No client secrets found on applications or service principals.');
      return fail(test, `Found ${appSecrets.length} applications and ${spSecrets.length} service principals with client secrets.`, {
        applications: appSecrets.slice(0, 20).map(app => ({ displayName: app.displayName, appId: app.appId })),
        servicePrincipals: spSecrets.slice(0, 20).map(sp => ({ displayName: sp.displayName, appId: sp.appId, ownerTenantId: sp.appOwnerOrganizationId })),
      });
    },

    '21773': async (test) => {
      const apps = await Api.graphAll('applications', {
        query: { '$select': 'id,appId,displayName,keyCredentials', '$top': '999' },
      });
      const tooLong = [];
      for (const app of apps) {
        for (const credential of toArray(app.keyCredentials)) {
          if (!credential.startDateTime || !credential.endDateTime) continue;
          if (daysBetween(credential.endDateTime, credential.startDateTime) > 180) {
            tooLong.push({
              displayName: app.displayName,
              appId: app.appId,
              startDateTime: credential.startDateTime,
              endDateTime: credential.endDateTime,
            });
          }
        }
      }
      if (!tooLong.length) return pass(test, 'No certificates with validity over 180 days found.');
      return fail(test, `Found ${tooLong.length} certificate credentials with validity longer than 180 days.`, tooLong.slice(0, 30));
    },

    '21774': async (test, ctx) => {
      const org = await Api.graph('organization', { query: { '$select': 'id,displayName,verifiedDomains' } });
      const tenantId = org?.value?.[0]?.id || ctx.tenantId;
      const servicePrincipals = await getServicePrincipals();
      const externalWithCreds = servicePrincipals.filter(sp =>
        sp.appOwnerOrganizationId && tenantId && sp.appOwnerOrganizationId !== tenantId &&
        (toArray(sp.passwordCredentials).length > 0 || toArray(sp.keyCredentials).length > 0)
      );
      if (!externalWithCreds.length) return pass(test, 'No external Microsoft service applications with credentials configured were found.');
      return fail(test, `Found ${externalWithCreds.length} external service principals with credentials configured.`, externalWithCreds.slice(0, 20).map(sp => ({ displayName: sp.displayName, appId: sp.appId, ownerTenantId: sp.appOwnerOrganizationId })));
    },

    '21775': async (test) => {
      const policy = await Api.graph('policies/defaultAppManagementPolicy', { beta: true });
      if (policy?.id) return pass(test, 'Default app management policy is configured.', { id: policy.id, displayName: policy.displayName || '' });
      return fail(test, 'Default app management policy endpoint returned no policy.');
    },

    '21776': async (test) => {
      const policy = await getAuthorizationPolicy();
      const grants = toArray(policy?.defaultUserRolePermissions?.permissionGrantPoliciesAssigned);
      const hasBroadSelfGrant = grants.some(grant => /^managePermissionGrantsForSelf/i.test(grant));
      const hasLowImpactGrant = grants.includes('managePermissionGrantsForSelf.microsoft-user-default-low');
      if (!hasBroadSelfGrant || hasLowImpactGrant) return pass(test, 'User consent settings are restricted.', { grants });
      return fail(test, 'User consent allows broad self-grant behavior.', { grants });
    },

    '21777': async (test) => {
      const apps = await Api.graphAll('applications', {
        beta: true,
        query: { '$select': 'id,displayName,signInAudience,servicePrincipalLockConfiguration', '$top': '999' },
      });
      const multiTenant = apps.filter(app => app.signInAudience && app.signInAudience !== 'AzureADMyOrg');
      const unlocked = multiTenant.filter(app => {
        const config = app.servicePrincipalLockConfiguration || {};
        return !(config.isEnabled || config.allProperties || config.credentialsWithUsageVerify || config.credentialsWithUsageSign);
      });
      if (!unlocked.length) return pass(test, 'Multitenant applications appear to have app instance property lock configured.');
      return fail(test, `Found ${unlocked.length} multitenant applications without visible property lock configuration.`, unlocked.slice(0, 25));
    },

    '21781': async (test) => {
      const roles = await Api.graphAll('directoryRoles', { query: { '$select': 'id,displayName' } });
      const targetRoles = roles.filter(role => /global administrator|privileged role administrator|security administrator/i.test(role.displayName));
      if (!targetRoles.length) return skip(test, 'No target privileged directory roles are active.');

      const members = [];
      for (const role of targetRoles) {
        const roleMembers = await Api.graphAll(`directoryRoles/${role.id}/members`, { query: { '$select': 'id,displayName,userPrincipalName' } });
        members.push(...roleMembers.filter(member => member.userPrincipalName));
      }

      const uniqueMembers = uniqueBy(members, member => member.id);
      let resistantCount = 0;
      const evidence = [];
      for (const member of uniqueMembers) {
        const methods = await getUserAuthMethods(member.id);
        const resistant = hasPhishingResistantRegisteredMethod(methods);
        if (resistant) resistantCount += 1;
        evidence.push({ displayName: member.displayName, userPrincipalName: member.userPrincipalName, resistant });
      }

      if (!uniqueMembers.length) return skip(test, 'Could not inspect authentication methods for privileged users.');
      if (resistantCount === uniqueMembers.length) return pass(test, `All sampled privileged users (${uniqueMembers.length}) have phishing-resistant methods registered.`, evidence);
      return fail(test, `${resistantCount}/${uniqueMembers.length} sampled privileged users have phishing-resistant methods registered.`, evidence);
    },

    '21782': async (test) => {
      const [roles, assignments] = await Promise.all([getPrivilegedBuiltInRoles(), getRoleAssignments()]);
      const roleMap = new Map(roles.map(role => [role.id, role.displayName]));
      const userAssignments = assignments.filter(assignment => roleMap.has(assignment.roleDefinitionId));
      const users = await getUsersByIds(userAssignments.map(assignment => assignment.principalId));
      const usersById = new Map(users.map(user => [user.id, user]));
      const rolesByUserId = new Map();
      for (const assignment of userAssignments) {
        if (!usersById.has(assignment.principalId)) continue;
        const list = rolesByUserId.get(assignment.principalId) || [];
        list.push(roleMap.get(assignment.roleDefinitionId));
        rolesByUserId.set(assignment.principalId, list);
      }

      const evidence = [];
      let resistantCount = 0;
      for (const [userId, roleNames] of rolesByUserId.entries()) {
        const user = usersById.get(userId);
        const methods = await getUserAuthMethods(userId);
        const resistant = hasPhishingResistantRegisteredMethod(methods);
        if (resistant) resistantCount += 1;
        evidence.push({ displayName: user.displayName, userPrincipalName: user.userPrincipalName, roles: uniqueBy(roleNames, name => name), resistant });
      }

      if (!evidence.length) return skip(test, 'No privileged user assignments could be evaluated.');
      if (resistantCount === evidence.length) return pass(test, 'All privileged accounts have phishing-resistant methods registered.', evidence);
      return fail(test, `Found ${evidence.length - resistantCount} privileged accounts without phishing-resistant methods registered.`, evidence);
    },

    '21783': async (test) => {
      const [roles, policies, authStrengthPolicies] = await Promise.all([
        getPrivilegedBuiltInRoles(),
        getConditionalAccessPolicies(),
        getAuthStrengthPolicies(),
      ]);

      const phishResistantStrengthIds = new Set(
        authStrengthPolicies
          .filter(policy => toArray(policy.allowedCombinations).every(isPhishingResistantCombination))
          .map(policy => policy.id)
      );

      const enabledPolicies = policies.filter(policy => lower(policy.state) === 'enabled');
      const protectingPolicies = enabledPolicies.filter(policy => includesRoles(policy, roles.map(role => role.id)) && phishResistantStrengthIds.has(getAuthStrengthId(policy)));
      const coveredRoleIds = new Set(protectingPolicies.flatMap(policy => toArray(policy?.conditions?.users?.includeRoles)));
      const uncoveredRoles = roles.filter(role => !coveredRoleIds.has(role.id));

      if (!uncoveredRoles.length) return pass(test, 'All privileged built-in roles are targeted by phishing-resistant Conditional Access policies.', protectingPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })));
      return fail(test, `Found ${uncoveredRoles.length} privileged built-in roles not covered by phishing-resistant Conditional Access policies.`, {
        uncoveredRoles: uncoveredRoles.map(role => ({ id: role.id, displayName: role.displayName })),
        matchingPolicies: protectingPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
      });
    },

    '21786': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const requiredApps = new Set([
        '00000002-0000-0ff1-ce00-000000000000',
        '00000003-0000-0ff1-ce00-000000000000',
      ]);
      const matching = policies.filter(policy => {
        const clientApps = toArray(policy?.conditions?.clientAppTypes);
        const platforms = toArray(policy?.conditions?.platforms?.includePlatforms);
        const apps = toArray(policy?.conditions?.applications?.includeApplications);
        return lower(policy.state) === 'enabled' &&
          clientApps.length === 1 && clientApps[0] === 'mobileAppsAndDesktopClients' &&
          platforms.length === 1 && lower(platforms[0]) === 'windows' &&
          [...requiredApps].every(appId => apps.includes(appId)) &&
          policy?.sessionControls?.secureSignInSession?.isEnabled === true;
      });
      if (matching.length) return pass(test, 'Token protection is enforced for Windows desktop and mobile sign-in activity.', matching.map(policy => ({ id: policy.id, displayName: policy.displayName })));
      return fail(test, 'No enabled Conditional Access policy fully matches the token protection requirements for user sign-in activity.');
    },

    '21787': async (test) => {
      const policy = await getAuthorizationPolicy();
      if (!policy?.defaultUserRolePermissions?.allowedToCreateTenants) return pass(test, 'Non-privileged users are restricted from creating tenants.');
      return fail(test, 'Non-privileged users are allowed to create tenants.', { allowedToCreateTenants: !!policy?.defaultUserRolePermissions?.allowedToCreateTenants });
    },

    '21790': async (test) => {
      const policy = await getCrossTenantDefaultPolicy();
      const collaborationBlocked = lower(policy?.b2bCollaborationOutbound?.usersAndGroups?.accessType) === 'blocked' && lower(policy?.b2bCollaborationOutbound?.applications?.accessType) === 'blocked';
      const directBlocked = lower(policy?.b2bDirectConnectOutbound?.usersAndGroups?.accessType) === 'blocked' && lower(policy?.b2bDirectConnectOutbound?.applications?.accessType) === 'blocked';
      if (collaborationBlocked && directBlocked) return pass(test, 'Outbound cross-tenant access defaults are configured to block access.', policy);
      return fail(test, 'Outbound cross-tenant access defaults allow unrestricted or partially unrestricted outbound access.', policy);
    },

    '21791': async (test) => {
      const policy = await getAuthorizationPolicy();
      const value = `${policy.allowInvitesFrom || ''}`;
      if (value && value !== 'everyone') return pass(test, `Guest invites are restricted: ${value}.`, { allowInvitesFrom: value });
      return fail(test, 'Guests may invite other guests.', { allowInvitesFrom: value || '(unset)' });
    },

    '21792': async (test) => {
      const policy = await getAuthorizationPolicy();
      const restrictiveRoleId = '10dae51f-b6af-4016-8d66-8c2a99b929b3';
      if (lower(policy.guestUserRoleId) === restrictiveRoleId) return pass(test, 'Guest access to directory objects is restricted.', { guestUserRoleId: policy.guestUserRoleId });
      return fail(test, 'Guest directory access is not set to the most restrictive built-in guest role.', { guestUserRoleId: policy.guestUserRoleId });
    },

    '21793': async (test) => {
      const policy = await getCrossTenantDefaultPolicy();
      if (policy?.b2bCollaborationInbound || policy?.b2bDirectConnectInbound || policy?.inboundTrust) return pass(test, 'Default cross-tenant access settings are configured.', policy);
      return fail(test, 'Cross-tenant access policy default exists but appears minimally configured.', policy);
    },

    '21796': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const matching = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        includesAllUsers(policy) &&
        hasGrantControl(policy, 'block') &&
        toArray(policy?.conditions?.clientAppTypes).includes('exchangeActiveSync') &&
        toArray(policy?.conditions?.clientAppTypes).includes('other')
      );

      if (matching.length) return pass(test, 'At least one enabled Conditional Access policy blocks legacy authentication for all users.', matching.map(policy => ({ id: policy.id, displayName: policy.displayName })));
      return fail(test, 'No enabled Conditional Access policy blocks legacy authentication for all users.');
    },

    '21797': async (test) => {
      const [authPolicy, policies] = await Promise.all([getAuthenticationMethodsPolicy(), getConditionalAccessPolicies()]);
      const passwordlessEnabled = toArray(authPolicy?.authenticationMethodConfigurations).some(method => {
        const id = lower(method.id);
        if (id === 'fido2') return lower(method.state) === 'enabled';
        if (id === 'x509certificate') return lower(method.state) === 'enabled' && lower(method.x509CertificateAuthenticationDefaultMode) === 'x509certificatemultifactor';
        return false;
      });

      const enabledPolicies = policies.filter(policy => lower(policy.state) === 'enabled');
      const passwordChangePolicies = enabledPolicies.filter(policy => toArray(policy?.conditions?.userRiskLevels).includes('high') && hasGrantControl(policy, 'passwordChange'));
      const blockPolicies = enabledPolicies.filter(policy => toArray(policy?.conditions?.userRiskLevels).includes('high') && hasGrantControl(policy, 'block'));
      const passed = (!passwordlessEnabled && (passwordChangePolicies.length + blockPolicies.length > 0)) || (passwordlessEnabled && blockPolicies.length > 0);

      if (passed) return pass(test, 'High-risk user protections are configured in line with the available passwordless posture.', {
        passwordlessEnabled,
        passwordChangePolicies: passwordChangePolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        blockPolicies: blockPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
      });

      return fail(test, 'High-risk user protections are incomplete for the current authentication posture.', {
        passwordlessEnabled,
        passwordChangePolicies: passwordChangePolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        blockPolicies: blockPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
      });
    },

    '21799': async (test) => {
      const [authPolicy, policies] = await Promise.all([getAuthenticationMethodsPolicy(), getConditionalAccessPolicies()]);
      const anyEnabledMethods = hasAnyEnabledAuthMethods(authPolicy);
      const matching = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        includesAllUsers(policy) &&
        toArray(policy?.conditions?.signInRiskLevels).includes('high') &&
        (hasGrantControl(policy, 'block') || (anyEnabledMethods && (hasGrantControl(policy, 'mfa') || !!getAuthStrengthId(policy))))
      );

      if (matching.length) return pass(test, 'Enabled Conditional Access policies mitigate high-risk sign-ins.', matching.map(policy => ({ id: policy.id, displayName: policy.displayName })));
      return fail(test, 'No enabled Conditional Access policy adequately mitigates high-risk sign-ins.');
    },

    '21802': async (test) => {
      const config = await getAuthenticatorConfig();
      const appInfoEnabled = featureSettingEnabledForAllUsers(config?.featureSettings?.displayAppInformationRequiredState);
      const locationEnabled = featureSettingEnabledForAllUsers(config?.featureSettings?.displayLocationInformationRequiredState);
      if (appInfoEnabled && locationEnabled) return pass(test, 'Microsoft Authenticator shows application and geographic location context for sign-in prompts.', config?.featureSettings);
      return fail(test, 'Microsoft Authenticator sign-in context is incomplete.', {
        displayAppInformationRequiredState: config?.featureSettings?.displayAppInformationRequiredState || null,
        displayLocationInformationRequiredState: config?.featureSettings?.displayLocationInformationRequiredState || null,
      });
    },

    '21803': async (test) => {
      const policy = await getAuthenticationMethodsPolicy();
      const state = policy?.policyMigrationState;
      if (!state || state === 'migrationComplete') return pass(test, `Policy migration state: ${state || 'not-applicable'}.`, { policyMigrationState: state || null });
      return fail(test, `Legacy MFA and SSPR migration is incomplete: ${state}.`, { policyMigrationState: state });
    },

    '21804': async (test) => {
      const policy = await getAuthenticationMethodsPolicy();
      const weakMethods = toArray(policy?.authenticationMethodConfigurations).filter(method => ['Sms', 'Voice'].includes(`${method.id}`));
      const enabledWeakMethods = weakMethods.filter(method => lower(method.state) === 'enabled');
      if (!enabledWeakMethods.length) return pass(test, 'SMS and Voice authentication methods are disabled.', weakMethods.map(method => ({ id: method.id, state: method.state })));
      return fail(test, 'Weak authentication methods are still enabled.', weakMethods.map(method => ({ id: method.id, state: method.state })));
    },

    '21806': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const matching = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        includesAllUsers(policy) &&
        includesUserAction(policy, 'urn:user:registersecurityinfo')
      );
      if (matching.length) return pass(test, 'Security information registration is protected by Conditional Access.', matching.map(policy => ({ id: policy.id, displayName: policy.displayName })));
      return fail(test, 'No enabled Conditional Access policy protects security information registration.');
    },

    '21807': async (test) => {
      const policy = await getAuthorizationPolicy();
      if (!policy?.defaultUserRolePermissions?.allowedToCreateApps) return pass(test, 'Creating new applications and service principals is restricted to privileged users.');
      return fail(test, 'Non-privileged users can register new applications and service principals.', { allowedToCreateApps: !!policy?.defaultUserRolePermissions?.allowedToCreateApps });
    },

    '21808': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const matching = policies.filter(policy => {
        const transferMethods = `${policy?.conditions?.authenticationFlows?.transferMethods || ''}`.split(',').map(item => item.trim()).filter(Boolean);
        return lower(policy.state) === 'enabled' && transferMethods.includes('deviceCodeFlow') && hasGrantControl(policy, 'block');
      });
      if (matching.length) return pass(test, 'Device code flow is restricted by enabled Conditional Access policy.', matching.map(policy => ({ id: policy.id, displayName: policy.displayName })));
      return fail(test, 'No enabled Conditional Access policy blocks device code flow.');
    },

    '21824': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const guestPolicies = policies.filter(policy =>
        ['enabled', 'enabledforreportingbutnotenforced'].includes(lower(policy.state)) &&
        policy?.conditions?.users?.includeGuestsOrExternalUsers &&
        toArray(policy?.grantControls?.termsOfUse).length === 0
      );

      const compliant = guestPolicies.filter(policy => {
        const freq = policy?.sessionControls?.signInFrequency;
        if (!freq || !freq.isEnabled) return false;
        return (lower(freq.type) === 'hours' && Number(freq.value) <= 24) ||
          (lower(freq.type) === 'days' && Number(freq.value) === 1) ||
          (lower(freq.frequencyInterval) === 'everytime');
      });

      if (guestPolicies.length && guestPolicies.length === compliant.length) return pass(test, 'Guest or external user policies enforce short-lived sign-in sessions.', compliant.map(policy => ({ id: policy.id, displayName: policy.displayName })));
      return fail(test, 'Guest or external user policies do not consistently enforce short-lived sign-in sessions.', {
        guestPolicies: guestPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        compliantPolicies: compliant.map(policy => ({ id: policy.id, displayName: policy.displayName })),
      });
    },

    '21822': async (test) => {
      const policies = await getLegacyPolicies();
      const b2bPolicy = policies.find(policy => `${policy?.type}` === 'B2BManagementPolicy');
      const definition = b2bPolicy?.definition ? JSON.parse(b2bPolicy.definition) : null;
      const policy = definition?.B2BManagementPolicy?.InvitationsAllowedAndBlockedDomainsPolicy || null;
      const allowedDomains = toArray(policy?.AllowedDomains);
      const blockedDomains = toArray(policy?.BlockedDomains);
      if (allowedDomains.length > 0) return pass(test, 'Guest access is limited to approved tenants.', { allowedDomains, blockedDomains });
      return fail(test, 'Guest access is not limited to approved tenants.', { allowedDomains, blockedDomains });
    },

    '21825': async (test) => {
      const [roles, policies] = await Promise.all([getPrivilegedBuiltInRoles(), getConditionalAccessPolicies()]);
      const recommendedMaxHours = 24;
      const enabledPolicies = policies.filter(policy => lower(policy.state) === 'enabled' && includesRoles(policy, roles.map(role => role.id)));
      const coveredRoleIds = new Set();
      for (const policy of enabledPolicies) {
        const freq = policy?.sessionControls?.signInFrequency;
        const compliant = freq && lower(freq.type) === 'hours' && Number(freq.value) <= recommendedMaxHours;
        if (!compliant) continue;
        for (const roleId of toArray(policy?.conditions?.users?.includeRoles)) coveredRoleIds.add(roleId);
      }
      const uncovered = roles.filter(role => !coveredRoleIds.has(role.id));
      if (!uncovered.length) return pass(test, 'Privileged roles are covered by sign-in frequency policies with short-lived sessions.', { recommendedMaxHours, coveredRoles: roles.length });
      return fail(test, `Found ${uncovered.length} privileged roles without compliant short-lived session policies.`, { recommendedMaxHours, uncoveredRoles: uncovered.map(role => ({ id: role.id, displayName: role.displayName })) });
    },

    '21828': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const matching = policies.filter(policy => {
        const transferMethods = `${policy?.conditions?.authenticationFlows?.transferMethods || ''}`.split(',').map(item => item.trim()).filter(Boolean);
        const includeUsers = toArray(policy?.conditions?.users?.includeUsers);
        const includeApps = toArray(policy?.conditions?.applications?.includeApplications);
        return lower(policy.state) === 'enabled' &&
          transferMethods.includes('authenticationTransfer') &&
          hasGrantControl(policy, 'block') &&
          includeUsers.includes('All') &&
          includeApps.includes('All');
      });
      if (matching.length) return pass(test, 'Authentication transfer is blocked by enabled Conditional Access policy.', matching.map(policy => ({ id: policy.id, displayName: policy.displayName })));
      return fail(test, 'No enabled Conditional Access policy blocks authentication transfer for all users and all applications.');
    },

    '21830': async (test) => {
      const [roles, policies] = await Promise.all([getPrivilegedBuiltInRoles(), getConditionalAccessPolicies()]);
      const roleIds = roles.map(role => role.id);
      const enabledPolicies = policies.filter(policy => lower(policy.state) === 'enabled' && includesRoles(policy, roleIds));
      const compliantDevicePolicies = enabledPolicies.filter(policy => hasGrantControl(policy, 'compliantDevice'));
      const deviceFilterBlockPolicies = enabledPolicies.filter(policy => lower(policy?.conditions?.devices?.deviceFilter?.mode) === 'exclude' && hasGrantControl(policy, 'block'));
      if (compliantDevicePolicies.length && deviceFilterBlockPolicies.length) return pass(test, 'Privileged role access is constrained by compliant-device and device-filter Conditional Access policies.', {
        compliantDevicePolicies: compliantDevicePolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        deviceFilterPolicies: deviceFilterBlockPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
      });
      return fail(test, 'Conditional Access policies for privileged access workstations are incomplete.', {
        compliantDevicePolicies: compliantDevicePolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        deviceFilterPolicies: deviceFilterBlockPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
      });
    },

    '21838': async (test) => {
      const config = await getFido2Config();
      if (lower(config?.state) === 'enabled') return pass(test, 'Security key authentication method is enabled.', { state: config.state, includeTargets: summarizeTargets(config.includeTargets), excludeTargets: summarizeTargets(config.excludeTargets) });
      return fail(test, 'Security key authentication method is not enabled.', { state: config?.state || 'unknown' });
    },

    '21839': async (test) => {
      const config = await getFido2Config();
      const enabled = lower(config?.state) === 'enabled';
      const hasTargets = toArray(config?.includeTargets).length > 0;
      if (enabled && hasTargets) return pass(test, 'Passkey authentication method is enabled and scoped to users.', { state: config.state, includeTargets: config.includeTargets });
      return fail(test, 'Passkey authentication method is disabled or not scoped to any users.', { state: config?.state || 'unknown', includeTargets: config?.includeTargets || [] });
    },

    '21840': async (test) => {
      const config = await getFido2Config();
      if (config?.isAttestationEnforced === true) return pass(test, 'Security key attestation is enforced.', { isAttestationEnforced: true, keyRestrictions: config?.keyRestrictions || null });
      return fail(test, 'Security key attestation is not enforced.', { isAttestationEnforced: !!config?.isAttestationEnforced, keyRestrictions: config?.keyRestrictions || null });
    },

    '21841': async (test) => {
      const policy = await getAuthenticationMethodsPolicy();
      const settings = policy?.reportSuspiciousActivitySettings;
      const passed = lower(settings?.state) === 'enabled' && lower(settings?.includeTarget?.id) === 'all_users';
      if (passed) return pass(test, 'Microsoft Authenticator suspicious activity reporting is enabled for all users.', settings);
      return fail(test, 'Microsoft Authenticator suspicious activity reporting is not enabled for all users.', settings || null);
    },

    '21842': async (test) => {
      const policy = await getAuthorizationPolicy();
      if (policy?.allowedToUseSspr === false) return pass(test, 'Administrators are blocked from using self-service password reset.', { allowedToUseSspr: false });
      return fail(test, 'Administrators can still use self-service password reset.', { allowedToUseSspr: policy?.allowedToUseSspr });
    },

    '21845': async (test) => {
      const [tapConfig, policies, authStrengthPolicies] = await Promise.all([
        getTapConfig(),
        getConditionalAccessPolicies(),
        getAuthStrengthPolicies(),
      ]);

      const securityInfoPolicies = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        includesUserAction(policy, 'urn:user:registersecurityinfo') &&
        !!getAuthStrengthId(policy)
      );
      const strengthIds = new Set(toArray(securityInfoPolicies.map(getAuthStrengthId)).filter(Boolean));
      const referencedPolicies = authStrengthPolicies.filter(policy => strengthIds.has(policy.id));
      const supportsTap = referencedPolicies.some(policy => toArray(policy.allowedCombinations).some(combo => lower(combo).includes('temporaryaccesspass')));
      const enabled = lower(tapConfig?.state) === 'enabled';
      const targetsAllUsers = includeTargetsAllUsers(tapConfig?.includeTargets);
      const passed = enabled && targetsAllUsers && securityInfoPolicies.length > 0 && supportsTap;

      if (passed) return pass(test, 'Temporary Access Pass is enabled, targets all users, and is enforced by Conditional Access.', {
        tapState: tapConfig.state,
        targetsAllUsers,
        securityInfoPolicies: securityInfoPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        authStrengthPolicies: referencedPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
      });

      return fail(test, 'Temporary Access Pass is not fully enabled and enforced.', {
        tapState: tapConfig?.state || 'unknown',
        targetsAllUsers,
        securityInfoPolicies: securityInfoPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        authStrengthPolicies: referencedPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        supportsTap,
      });
    },

    '21846': async (test) => {
      const config = await getTapConfig();
      if (config?.isUsableOnce === true) return pass(test, 'Temporary Access Pass is restricted to single use.', { isUsableOnce: true, state: config?.state || 'unknown' });
      return fail(test, 'Temporary Access Pass allows multiple uses during the validity period.', { isUsableOnce: !!config?.isUsableOnce, state: config?.state || 'unknown' });
    },

    '21872': async (test) => {
      const [policies, deviceRegistrationPolicy] = await Promise.all([getConditionalAccessPolicies(), getDeviceRegistrationPolicy()]);
      const deviceRegistrationPolicies = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        includesUserAction(policy, 'urn:user:registerdevice') &&
        (hasGrantControl(policy, 'mfa') || !!getAuthStrengthId(policy))
      );
      const mfaRequiredInDeviceSettings = lower(deviceRegistrationPolicy?.multiFactorAuthConfiguration) === 'required';

      if (!mfaRequiredInDeviceSettings && deviceRegistrationPolicies.length) {
        return pass(test, 'Device registration is protected by Conditional Access rather than device settings.', {
          multiFactorAuthConfiguration: deviceRegistrationPolicy?.multiFactorAuthConfiguration || null,
          policies: deviceRegistrationPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        });
      }

      return fail(test, 'Device registration MFA enforcement is incomplete or configured in the legacy device setting.', {
        multiFactorAuthConfiguration: deviceRegistrationPolicy?.multiFactorAuthConfiguration || null,
        policies: deviceRegistrationPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
      });
    },

    '21861': async (test) => {
      const riskyUsers = await getRiskyUsers("riskState eq 'atRisk' and riskLevel eq 'high'");
      if (riskyUsers === null) return skip(test, 'Identity Protection is not available in this tenant (P2 license required).');
      if (!riskyUsers.length) return pass(test, 'No untriaged high-risk users were found in Identity Protection.');
      return fail(test, `Found ${riskyUsers.length} untriaged high-risk users in Identity Protection.`, riskyUsers.map(user => ({ userPrincipalName: user.userPrincipalName || user.id, riskLevel: user.riskLevel, riskState: user.riskState, riskLastUpdatedDateTime: user.riskLastUpdatedDateTime })));
    },

    '21862': async (test) => {
      const [riskyPrincipals, detections] = await Promise.all([getRiskyServicePrincipals(), getServicePrincipalRiskDetections()]);
      if (riskyPrincipals === null || detections === null) return skip(test, 'Identity Protection workload identity risk data is not available in this tenant (P2 license required).');
      const atRiskDetections = detections.filter(detection => lower(detection.riskState) === 'atrisk');
      if (!riskyPrincipals.length && !atRiskDetections.length) return pass(test, 'No untriaged risky workload identities were found.');
      return fail(test, `Found ${riskyPrincipals.length} risky workload identities and ${atRiskDetections.length} risky workload identity detections requiring triage.`, {
        riskyServicePrincipals: riskyPrincipals.map(item => ({ displayName: item.displayName, appId: item.appId, riskLevel: item.riskLevel, riskLastUpdatedDateTime: item.riskLastUpdatedDateTime })),
        detections: atRiskDetections.map(item => ({ servicePrincipalDisplayName: item.servicePrincipalDisplayName, appId: item.appId, riskLevel: item.riskLevel, detectedDateTime: item.detectedDateTime })),
      });
    },

    '21863': async (test) => {
      const detections = await getRiskDetections("riskState eq 'atRisk' and riskLevel eq 'high'");
      if (detections === null) return skip(test, 'Identity Protection sign-in risk data is not available in this tenant (P2 license required).');
      if (!detections.length) return pass(test, 'No untriaged high-risk sign-ins were found.');
      return fail(test, `Found ${detections.length} untriaged high-risk sign-ins.`, detections.map(item => ({ userPrincipalName: item.userPrincipalName, riskEventType: item.riskEventType, riskLevel: item.riskLevel, detectedDateTime: item.detectedDateTime })));
    },

    '21884': async (test) => {
      const [organization, servicePrincipals, policies, namedLocations] = await Promise.all([
        getOrganization(),
        getServicePrincipals(),
        getConditionalAccessPolicies(),
        getNamedLocations(),
      ]);

      const tenantId = organization?.id || null;
      const principalsWithCreds = servicePrincipals.filter(sp =>
        lower(sp.servicePrincipalType) === 'application' &&
        (!tenantId || sp.appOwnerOrganizationId === tenantId) &&
        (toArray(sp.passwordCredentials).length > 0 || toArray(sp.keyCredentials).length > 0)
      );

      if (!principalsWithCreds.length) return pass(test, 'No workload identities with credentials were found to evaluate.');

      const locationNames = new Map(namedLocations.map(location => [location.id, location.displayName]));
      const enabledPolicies = policies.filter(policy => lower(policy.state) === 'enabled');
      const globalPolicies = enabledPolicies.filter(policy => toArray(policy?.conditions?.clientApplications?.includeServicePrincipals).includes('ServicePrincipalsInMyTenant'));

      const evidence = principalsWithCreds.slice(0, 200).map(sp => {
        const matchedPolicies = enabledPolicies.filter(policy => {
          const includePrincipals = toArray(policy?.conditions?.clientApplications?.includeServicePrincipals);
          return includePrincipals.includes('ServicePrincipalsInMyTenant') || includePrincipals.includes(sp.id);
        });

        const locationProtectedPolicies = matchedPolicies.filter(policy => {
          const locations = policy?.conditions?.locations || {};
          return toArray(locations.includeLocations).length > 0 || toArray(locations.excludeLocations).length > 0;
        });

        const readableLocations = uniqueBy(locationProtectedPolicies.flatMap(policy => [
          ...toArray(policy?.conditions?.locations?.includeLocations),
          ...toArray(policy?.conditions?.locations?.excludeLocations),
        ]), item => `${item}`).map(id => locationNames.get(id) || id);

        return {
          displayName: sp.displayName,
          appId: sp.appId,
          credentialTypes: [toArray(sp.passwordCredentials).length ? 'Password' : null, toArray(sp.keyCredentials).length ? 'Certificate' : null].filter(Boolean),
          protected: locationProtectedPolicies.length > 0,
          policies: locationProtectedPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
          namedLocations: readableLocations,
        };
      });

      const unprotected = evidence.filter(item => !item.protected);
      if (!unprotected.length) return pass(test, 'Workload identities with credentials are covered by location-based Conditional Access policies.', { globalPolicies: globalPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })), evaluatedPrincipals: evidence.length, sample: evidence.slice(0, 40) });
      return fail(test, `Found ${unprotected.length} workload identities without location-based Conditional Access protection in the evaluated sample.`, { globalPolicies: globalPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })), evaluatedPrincipals: evidence.length, unprotected: unprotected.slice(0, 40) });
    },

    '21889': async (test) => {
      const authPolicy = await getAuthenticationMethodsPolicy();
      const fido2 = toArray(authPolicy?.authenticationMethodConfigurations).find(method => `${method.id}` === 'Fido2') || null;
      const authenticator = toArray(authPolicy?.authenticationMethodConfigurations).find(method => `${method.id}` === 'MicrosoftAuthenticator') || null;
      const fido2Valid = lower(fido2?.state) === 'enabled' && toArray(fido2?.includeTargets).length > 0;
      const authMode = lower(toArray(authenticator?.includeTargets)[0]?.authenticationMode || authenticator?.includeTargets?.authenticationMode || '');
      const authValid = lower(authenticator?.state) === 'enabled' && toArray(authenticator?.includeTargets).length > 0 && ['any', 'devicebasedpush'].includes(authMode);
      if (fido2Valid && authValid) return pass(test, 'Multiple passwordless methods are configured, reducing the user-visible password surface area.', {
        fido2: { state: fido2?.state || 'unknown', includeTargets: fido2?.includeTargets || [] },
        authenticator: { state: authenticator?.state || 'unknown', includeTargets: authenticator?.includeTargets || [], authenticationMode: authMode || 'not-configured' },
      });
      return fail(test, 'Passwordless coverage is incomplete for FIDO2 and Microsoft Authenticator.', {
        fido2: { state: fido2?.state || 'unknown', includeTargets: fido2?.includeTargets || [] },
        authenticator: { state: authenticator?.state || 'unknown', includeTargets: authenticator?.includeTargets || [], authenticationMode: authMode || 'not-configured' },
      });
    },

    '21892': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const matching = policies.filter(policy => {
        const enabled = lower(policy.state) === 'enabled';
        const allUsers = includesAllUsers(policy);
        const allApps = toArray(policy?.conditions?.applications?.includeApplications).includes('All');
        const deviceControl = hasGrantControl(policy, 'compliantDevice') || hasGrantControl(policy, 'domainJoinedDevice');
        return enabled && allUsers && allApps && deviceControl;
      });
      if (matching.length) return pass(test, 'At least one enabled Conditional Access policy restricts all sign-ins to managed devices.', matching.map(policy => ({ id: policy.id, displayName: policy.displayName, controls: toArray(policy?.grantControls?.builtInControls) })));
      return fail(test, 'No enabled Conditional Access policy requires managed devices for all users and all apps.');
    },

    '21941': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const requiredAppIds = new Set([
        '00000002-0000-0ff1-ce00-000000000000',
        '00000003-0000-0ff1-ce00-000000000000',
      ]);
      const matching = policies.filter(policy => {
        const secureSessionEnabled = policy?.sessionControls?.secureSignInSession?.isEnabled === true;
        const includeUsers = toArray(policy?.conditions?.users?.includeUsers);
        const includeApps = toArray(policy?.conditions?.applications?.includeApplications);
        const coversApps = includeApps.includes('All') || [...requiredAppIds].every(appId => includeApps.includes(appId));
        return lower(policy.state) === 'enabled' && secureSessionEnabled && includeUsers.length > 0 && !includeUsers.includes('None') && coversApps;
      });
      if (matching.length) return pass(test, 'Token protection Conditional Access policies are configured and enabled.', matching.map(policy => ({ id: policy.id, displayName: policy.displayName })));
      return fail(test, 'No enabled Conditional Access policy meets the token protection requirements.');
    },

    '21964': async (test) => {
      const protectedActionIds = [
        'microsoft.directory-conditionalAccessPolicies-basic-update-patch',
        'microsoft.directory-conditionalAccessPolicies-create-post',
        'microsoft.directory-conditionalAccessPolicies-delete-delete',
        'microsoft.directory-resourceNamespaces-resourceActions-authenticationContext-update-post',
      ];

      const [actionResults, policies, authContexts, authStrengthPolicies] = await Promise.all([
        Promise.all(protectedActionIds.map(id => Api.graph(`roleManagement/directory/resourceNamespaces/microsoft.directory/resourceActions/${id}`, {
          beta: true,
          query: { '$select': 'authenticationContextId,isAuthenticationContextSettable,name,description' },
        }).catch(err => (err.status === 404 || err.status === 403) ? null : Promise.reject(err)))),
        getConditionalAccessPolicies(),
        getAuthContextClassReferences(),
        getAuthStrengthPolicies(),
      ]);

      const actions = actionResults.filter(a => a !== null);
      if (!actions.length) return skip(test, 'Protected action resource endpoints are not accessible (endpoint may not exist or require higher permissions).');

      const authContextMap = new Map(authContexts.map(item => [item.id, item]));
      const phishResistantIds = new Set(
        authStrengthPolicies
          .filter(policy => toArray(policy.allowedCombinations).some(isPhishingResistantCombination))
          .map(policy => policy.id)
      );

      const evidence = actions.map(action => {
        const authContextId = action?.authenticationContextId || null;
        const matchedPolicies = policies.filter(policy => toArray(policy?.conditions?.applications?.includeAuthenticationContextClassReferences).includes(authContextId));
        const enabledPolicies = matchedPolicies.filter(policy => lower(policy.state) === 'enabled');
        const compliantPolicies = enabledPolicies.filter(policy => {
          const hasAuthStrength = phishResistantIds.has(getAuthStrengthId(policy));
          const hasSessionControl = !!policy?.sessionControls?.signInFrequency;
          const hasDeviceFilter = !!policy?.conditions?.devices;
          return (hasAuthStrength && hasDeviceFilter) || hasSessionControl;
        });

        return {
          actionId: action?.name || null,
          description: action?.description || null,
          authenticationContextId: authContextId,
          authenticationContextName: authContextMap.get(authContextId)?.displayName || null,
          enabledPolicies: enabledPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
          compliantPolicies: compliantPolicies.map(policy => ({ id: policy.id, displayName: policy.displayName })),
        };
      });

      const passed = evidence.every(item => item.authenticationContextId && item.compliantPolicies.length > 0);
      if (passed) return pass(test, 'Protected actions for Conditional Access management are secured by compliant Conditional Access policies.', evidence);
      return fail(test, 'Protected actions for Conditional Access management are missing authentication contexts or compliant Conditional Access policies.', evidence);
    },

    '22659': async (test) => {
      const detections = await getServicePrincipalRiskDetections();
      if (detections === null) return skip(test, 'Identity Protection workload identity risk data is not available in this tenant (P2 license required).');
      const signInDetections = detections.filter(detection => lower(detection.activity) === 'signin' && lower(detection.riskState) === 'atrisk');
      if (!signInDetections.length) return pass(test, 'No risky workload identity sign-ins require triage.');
      return fail(test, `Found ${signInDetections.length} risky workload identity sign-ins that require triage.`, signInDetections.map(item => ({ servicePrincipalDisplayName: item.servicePrincipalDisplayName, appId: item.appId, riskLevel: item.riskLevel, detectedDateTime: item.detectedDateTime })));
    },

    '25377': async (test) => {
      const [settings, policy] = await Promise.all([getNetworkCrossTenantAccessSettings(), getCrossTenantDefaultPolicy()]);
      if (settings === null) return skip(test, 'Global Secure Access network cross-tenant access settings are not available in this tenant.');
      const taggingEnabled = lower(settings?.networkPacketTaggingStatus) === 'enabled';
      const usersBlocked = lower(policy?.tenantRestrictions?.usersAndGroups?.accessType) === 'blocked' && toArray(policy?.tenantRestrictions?.usersAndGroups?.targets).some(item => item?.target === 'AllUsers');
      const appsBlocked = lower(policy?.tenantRestrictions?.applications?.accessType) === 'blocked' && toArray(policy?.tenantRestrictions?.applications?.targets).some(item => item?.target === 'AllApplications');
      if (taggingEnabled && usersBlocked && appsBlocked) return pass(test, 'Universal tenant restrictions block unauthorized external tenant access.', {
        networkPacketTaggingStatus: settings?.networkPacketTaggingStatus || null,
        tenantRestrictions: policy?.tenantRestrictions || null,
      });
      return fail(test, 'Universal tenant restrictions are not fully configured to block unauthorized external tenant access.', {
        networkPacketTaggingStatus: settings?.networkPacketTaggingStatus || null,
        tenantRestrictions: policy?.tenantRestrictions || null,
      });
    },

    '21954': async (test) => {
      const policy = await getAuthorizationPolicy();
      if (policy?.defaultUserRolePermissions?.allowedToReadBitlockerKeysForOwnedDevice === false) return pass(test, 'Non-administrator users are restricted from recovering BitLocker keys for their owned devices.');
      return fail(test, 'Non-administrator users can recover BitLocker keys for their owned devices.', { allowedToReadBitlockerKeysForOwnedDevice: policy?.defaultUserRolePermissions?.allowedToReadBitlockerKeysForOwnedDevice });
    },

    '27004': async (test) => {
      const sys = await fetch('./powershell/assets/27004-system-bypass-fqdns.json').then(response => response.json());
      const systemFqdns = new Set((sys?.fqdns || []).map(item => `${item}`.toLowerCase()));

      const policies = await Api.graph('networkAccess/tlsInspectionPolicies', {
        beta: true,
        query: { '$expand': 'policyRules' },
      });
      const list = Array.isArray(policies?.value) ? policies.value : [];
      if (!list.length) return skip(test, 'TLS inspection policies are not configured in this tenant.');

      const redundant = [];
      for (const policy of list) {
        for (const rule of toArray(policy.policyRules)) {
          if (rule.action !== 'bypass') continue;
          for (const destination of toArray(rule?.matchingConditions?.destinations)) {
            for (const value of toArray(destination?.values)) {
              if (systemFqdns.has(lower(value))) {
                redundant.push({ policyName: policy.name, ruleName: rule.name, destination: value });
              }
            }
          }
        }
      }

      if (!redundant.length) return pass(test, 'No duplicate custom TLS bypass destinations were found in the system bypass list.');
      return fail(test, `Found ${redundant.length} duplicate custom TLS bypass destinations.`, redundant.slice(0, 30));
    },
    '21784': async (test) => {
      const [policies, authStrengthPolicies] = await Promise.all([getConditionalAccessPolicies(), getAuthStrengthPolicies()]);
      const phishResistantStrengthIds = new Set(
        authStrengthPolicies
          .filter(policy => toArray(policy.allowedCombinations).every(isPhishingResistantCombination))
          .map(policy => policy.id)
      );
      const matching = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        includesAllUsers(policy) &&
        phishResistantStrengthIds.has(getAuthStrengthId(policy))
      );
      if (matching.length) return pass(test, 'All users are required to use phishing-resistant authentication methods by Conditional Access policies.', matching.map(p => ({ id: p.id, displayName: p.displayName })));
      return fail(test, 'No enabled Conditional Access policy requires all users to use phishing-resistant authentication methods.');
    },

    '21809': async (test) => {
      const policy = await getAdminConsentRequestPolicy();
      if (policy?.isEnabled === true) return pass(test, 'Admin consent workflow is enabled.', { isEnabled: policy.isEnabled, version: policy.version, notifyReviewers: policy.notifyReviewers });
      return fail(test, 'Admin consent workflow is not enabled. Users who are denied consent cannot request admin approval.', { isEnabled: policy?.isEnabled || false });
    },

    '21811': async (test) => {
      const payload = await Api.graph('domains', { query: { '$select': 'id,isDefault,isInitial,passwordValidityPeriodInDays', '$top': '50' } });
      const domains = toArray(payload?.value);
      const managedDomains = domains.filter(d => !d.isInitial);
      const domainsToCheck = managedDomains.length ? managedDomains : domains;
      if (!domainsToCheck.length) return skip(test, 'No domain information available.');
      const withExpiry = domainsToCheck.filter(d => d.passwordValidityPeriodInDays && d.passwordValidityPeriodInDays < 2147483647);
      if (!withExpiry.length) return pass(test, 'Password expiration is disabled (set to never expire) on all checked domains.', domainsToCheck.map(d => ({ id: d.id, passwordValidityPeriodInDays: d.passwordValidityPeriodInDays })));
      return fail(test, `${withExpiry.length} domain(s) have password expiration enabled.`, withExpiry.map(d => ({ id: d.id, passwordValidityPeriodInDays: d.passwordValidityPeriodInDays })));
    },

    '21812': async (test) => {
      const roles = await Api.graphAll('directoryRoles', { query: { '$select': 'id,displayName' } });
      const gaRole = roles.find(role => /^global administrator$/i.test(role.displayName));
      if (!gaRole) return skip(test, 'Global Administrator role is not active in this tenant.');
      const members = await Api.graphAll(`directoryRoles/${gaRole.id}/members`, { query: { '$select': 'id,displayName,userPrincipalName' } });
      const userMembers = members.filter(m => m.userPrincipalName);
      if (userMembers.length <= 8) return pass(test, `Global Administrator has ${userMembers.length} user member(s) — within the recommended maximum of 8.`, { count: userMembers.length });
      return fail(test, `Global Administrator has ${userMembers.length} user members, exceeding the recommended maximum of 8.`, { count: userMembers.length, members: userMembers.map(m => ({ displayName: m.displayName, userPrincipalName: m.userPrincipalName })) });
    },

    '21813': async (test) => {
      const [roles, assignments] = await Promise.all([getPrivilegedBuiltInRoles(), getRoleAssignments()]);
      const roleMap = new Map(roles.map(r => [r.id, r.displayName]));
      const users = await getUsersByIds(assignments.filter(a => roleMap.has(a.roleDefinitionId)).map(a => a.principalId));
      const usersById = new Map(users.map(u => [u.id, u]));
      const gaRole = roles.find(r => /^global administrator$/i.test(r.displayName));
      if (!gaRole) return skip(test, 'Global Administrator role not found in privileged built-in roles.');
      const gaAssignees = new Set(assignments.filter(a => a.roleDefinitionId === gaRole.id && usersById.has(a.principalId)).map(a => a.principalId));
      const privAssignees = new Set(assignments.filter(a => roleMap.has(a.roleDefinitionId) && usersById.has(a.principalId)).map(a => a.principalId));
      const gaCount = gaAssignees.size;
      const privCount = privAssignees.size;
      const ratio = privCount > 0 ? gaCount / privCount : 0;
      if (ratio <= 0.5) return pass(test, `Global Administrator count (${gaCount}) is within an acceptable ratio of all privileged users (${privCount}).`, { gaCount, privCount, ratio: ratio.toFixed(2) });
      return fail(test, `Global Administrator count (${gaCount}) is high relative to total privileged users (${privCount}). Ratio: ${ratio.toFixed(2)}.`, { gaCount, privCount, ratio: ratio.toFixed(2) });
    },

    '21814': async (test) => {
      const [roles, assignments] = await Promise.all([getPrivilegedBuiltInRoles(), getRoleAssignments()]);
      const roleMap = new Map(roles.map(r => [r.id, r.displayName]));
      const users = await getUsersByIds(assignments.filter(a => roleMap.has(a.roleDefinitionId)).map(a => a.principalId));
      const usersById = new Map(users.map(u => [u.id, u]));
      const syncedPriv = [];
      for (const assignment of assignments) {
        if (!roleMap.has(assignment.roleDefinitionId)) continue;
        const user = usersById.get(assignment.principalId);
        if (!user) continue;
        if (user.onPremisesSyncEnabled) {
          syncedPriv.push({ displayName: user.displayName, userPrincipalName: user.userPrincipalName, role: roleMap.get(assignment.roleDefinitionId) });
        }
      }
      if (!syncedPriv.length) return pass(test, 'All sampled privileged accounts appear to be cloud-native identities (not synced from on-premises).');
      return fail(test, `Found ${syncedPriv.length} privileged accounts synced from on-premises Active Directory.`, syncedPriv.slice(0, 30));
    },

    '21815': async (test) => {
      const [roles, instances] = await Promise.all([
        getPrivilegedBuiltInRoles(),
        getRoleAssignmentScheduleInstances("assignmentType eq 'Assigned'"),
      ]);
      if (instances === null) return skip(test, 'Role assignment schedule instances are not accessible (P2 PIM license may be required).');
      const roleMap = new Map(roles.map(r => [r.id, r.displayName]));
      const permanentPriv = instances.filter(i => roleMap.has(i.roleDefinitionId));
      if (!permanentPriv.length) return pass(test, 'No permanent (non-PIM-activated) privileged role assignments found — all assignments appear to be JIT.');
      const users = await getUsersByIds(permanentPriv.map(i => i.principalId));
      const userMap = new Map(users.map(u => [u.id, u]));
      return fail(test, `Found ${permanentPriv.length} permanent active privileged role assignments not activated via PIM.`, permanentPriv.slice(0, 30).map(i => ({
        principalId: i.principalId,
        displayName: userMap.get(i.principalId)?.displayName || i.principalId,
        role: roleMap.get(i.roleDefinitionId),
        assignmentType: i.assignmentType,
      })));
    },

    '21816': async (test) => {
      const schedules = await getRoleEligibilitySchedules();
      if (schedules === null) return skip(test, 'PIM role eligibility schedules are not accessible (P2 license may be required).');
      const [roles, assignments] = await Promise.all([getPrivilegedBuiltInRoles(), getRoleAssignments()]);
      const roleMap = new Map(roles.map(r => [r.id, r.displayName]));
      const privAssignments = assignments.filter(a => roleMap.has(a.roleDefinitionId));
      const eligibleIds = new Set(schedules.map(s => `${s.principalId}:${s.roleDefinitionId}`));
      const notInPIM = privAssignments.filter(a => !eligibleIds.has(`${a.principalId}:${a.roleDefinitionId}`));
      if (!notInPIM.length) return pass(test, `All ${privAssignments.length} privileged role assignments appear to be managed through PIM.`, { eligibilitySchedules: schedules.length });
      return fail(test, `Found ${notInPIM.length} privileged role assignments that may not be managed through PIM.`, { total: privAssignments.length, notInPIM: notInPIM.slice(0, 20) });
    },

    '21817': async (test) => {
      const gaRoleTemplateId = '62e90394-69f5-4237-9190-012177145e10';
      try {
        const assignments = await getRoleManagementPolicyAssignments(`scopeType eq 'DirectoryRole' and scopeId eq '/' and roleDefinitionId eq '${gaRoleTemplateId}'`, 10);
        if (assignments === null) return skip(test, 'PIM role management policy assignments are not accessible (license/provider/permission limitation).');
        if (!assignments.length) return fail(test, 'No PIM policy assignment found for the Global Administrator role.');

        const policyId = assignments[0]?.policyId;
        if (!policyId) return fail(test, 'Global Administrator PIM policy assignment was found, but no policyId is present.');

        const approvalRule = await Api.graph(`policies/roleManagementPolicies/${policyId}/rules/Approval_EndUser_Assignment`);
        const setting = approvalRule?.setting || null;
        const approvalRequired = setting?.isApprovalRequired === true;
        const approverCount = toArray(setting?.approvalStages).reduce((count, stage) => count + toArray(stage?.primaryApprovers).length, 0);

        if (approvalRequired && approverCount > 0) {
          return pass(test, `Global Administrator activation approval workflow is configured with ${approverCount} primary approver(s).`, {
            policyId,
            approvalRequired,
            approverCount,
          });
        }
        if (approvalRequired && approverCount === 0) {
          return fail(test, 'Global Administrator activation requires approval, but no primary approvers are configured.', {
            policyId,
            approvalRequired,
            approverCount,
          });
        }
        return fail(test, 'Approval is not required for Global Administrator role activation.', {
          policyId,
          approvalRequired,
          approverCount,
        });
      } catch (err) {
        if (err.status === 400 || err.status === 403 || err.status === 404) return skip(test, 'PIM role management policies are not accessible in this tenant (P2 license may be required or endpoint is unavailable).');
        throw err;
      }
    },

    '21821': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const guestPolicies = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        policy?.conditions?.users?.includeGuestsOrExternalUsers
      );
      if (guestPolicies.length) return pass(test, `Found ${guestPolicies.length} enabled Conditional Access policy(s) targeting guest or external users.`, guestPolicies.map(p => ({ id: p.id, displayName: p.displayName })));
      return fail(test, 'No enabled Conditional Access policies target guest or external users.');
    },

    '21823': async (test) => {
      const policy = await getAuthorizationPolicy();
      const emailVerifiedJoin = policy?.allowEmailVerifiedUsersToJoinOrganization;
      if (emailVerifiedJoin === false) return pass(test, 'Email-verified users cannot self-service join the organization (guest self-service sign-up is restricted).', { allowEmailVerifiedUsersToJoinOrganization: false });
      return fail(test, 'Email-verified users can self-service join the organization.', { allowEmailVerifiedUsersToJoinOrganization: emailVerifiedJoin });
    },

    '21836': async (test) => {
      const [roles, assignments, sps] = await Promise.all([getPrivilegedBuiltInRoles(), getRoleAssignments(), getServicePrincipals()]);
      const roleMap = new Map(roles.map(r => [r.id, r.displayName]));
      const spMap = new Map(sps.map(sp => [sp.id, sp]));
      const matches = assignments
        .filter(a => roleMap.has(a.roleDefinitionId) && spMap.has(a.principalId))
        .map(a => ({ displayName: spMap.get(a.principalId).displayName, appId: spMap.get(a.principalId).appId, roleName: roleMap.get(a.roleDefinitionId) }));
      if (!matches.length) return pass(test, 'No workload identities found with privileged built-in role assignments.');
      return fail(test, `Found ${matches.length} workload identity assignment(s) to privileged roles.`, matches.slice(0, 30));
    },

    '21837': async (test) => {
      const policy = await getDeviceRegistrationPolicy();
      const limit = policy?.userDeviceQuota;
      if (limit && limit <= 10) return pass(test, `Device quota is set to ${limit} per user — within the recommended maximum of 10.`, { userDeviceQuota: limit });
      if (!limit) return fail(test, 'Device quota is set to unlimited (no maximum device per user limit).', { userDeviceQuota: null });
      return fail(test, `Device quota is set to ${limit} per user, exceeding the recommended maximum of 10.`, { userDeviceQuota: limit });
    },

    '21843': async (test) => {
      const policy = await getAuthorizationPolicy();
      if (policy && typeof policy.blockMsolPowerShell === 'boolean') {
        if (policy.blockMsolPowerShell) return pass(test, 'Legacy Microsoft Online PowerShell (MSOL) is blocked via authorization policy.', { blockMsolPowerShell: true });
        return fail(test, 'Legacy Microsoft Online PowerShell (MSOL) is not blocked (blockMsolPowerShell=false).', { blockMsolPowerShell: false });
      }

      const legacyAppIds = new Set(['1b730954-1685-4b74-9bfd-dac224a7b894']);
      const policies = await getConditionalAccessPolicies();
      const blocking = policies.filter(policyItem =>
        lower(policyItem.state) === 'enabled' &&
        hasGrantControl(policyItem, 'block') &&
        toArray(policyItem?.conditions?.applications?.includeApplications).some(id => legacyAppIds.has(lower(id)))
      );
      if (blocking.length) return pass(test, 'Legacy Microsoft Online PowerShell module appears blocked by Conditional Access fallback detection.', blocking.map(p => ({ id: p.id, displayName: p.displayName })));
      return skip(test, 'blockMsolPowerShell setting is not exposed in this tenant response, and no clear CA fallback signal was detected.');
    },

    '21844': async (test) => {
      const legacyAzureAdAppIds = new Set(['1b730954-1685-4b74-9bfd-dac224a7b894', '00000002-0000-0000-c000-000000000000']);
      const policies = await getConditionalAccessPolicies();
      const blocking = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        hasGrantControl(policy, 'block') &&
        toArray(policy?.conditions?.applications?.includeApplications).some(id => legacyAzureAdAppIds.has(lower(id)))
      );
      if (blocking.length) return pass(test, 'Legacy Azure AD PowerShell module or Azure AD Graph API access is blocked by Conditional Access.', blocking.map(p => ({ id: p.id, displayName: p.displayName })));
      return fail(test, 'No enabled Conditional Access policy blocks the legacy Azure AD PowerShell module or Azure AD Graph API.');
    },

    '21849': async (test) => {
      const settings = await getDirectorySettings();
      const pwdSettings = settings.find(s => lower(s.displayName) === 'passwordrulesettings' || s.templateId === '5cf42378-d67d-4f36-ba46-e8b86229381d');
      if (!pwdSettings) return skip(test, 'Smart lockout directory settings not found. Default tenant values may apply.');
      const lockoutDuration = toArray(pwdSettings.values).find(v => lower(v.name) === 'lockoutdurationinseconds')?.value;
      const durationSecs = parseInt(lockoutDuration, 10);
      if (isNaN(durationSecs)) return skip(test, 'Smart lockout duration value is not set in directory settings.');
      if (durationSecs >= 60) return pass(test, `Smart lockout duration is ${durationSecs} seconds — meets the 60-second minimum.`, { lockoutDurationInSeconds: durationSecs });
      return fail(test, `Smart lockout duration is ${durationSecs} seconds, which is below the recommended minimum of 60 seconds.`, { lockoutDurationInSeconds: durationSecs });
    },

    '21850': async (test) => {
      const settings = await getDirectorySettings();
      const pwdSettings = settings.find(s => lower(s.displayName) === 'passwordrulesettings' || s.templateId === '5cf42378-d67d-4f36-ba46-e8b86229381d');
      if (!pwdSettings) return skip(test, 'Smart lockout directory settings not found. Default tenant values may apply.');
      const threshold = toArray(pwdSettings.values).find(v => lower(v.name) === 'lockoutthreshold')?.value;
      const thresholdNum = parseInt(threshold, 10);
      if (isNaN(thresholdNum)) return skip(test, 'Smart lockout threshold value is not set in directory settings.');
      if (thresholdNum <= 10) return pass(test, `Smart lockout threshold is ${thresholdNum} attempts — within the recommended maximum of 10.`, { lockoutThreshold: thresholdNum });
      return fail(test, `Smart lockout threshold is ${thresholdNum} attempts, exceeding the recommended maximum of 10.`, { lockoutThreshold: thresholdNum });
    },

    '21851': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const guestPolicies = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        policy?.conditions?.users?.includeGuestsOrExternalUsers &&
        (hasGrantControl(policy, 'mfa') || !!getAuthStrengthId(policy))
      );
      if (guestPolicies.length) return pass(test, 'Guest and external user access is protected by strong authentication in enabled Conditional Access policies.', guestPolicies.map(p => ({ id: p.id, displayName: p.displayName })));
      return fail(test, 'No enabled Conditional Access policy requires strong authentication for guest or external users.');
    },

    '21854': async (test) => {
      const [roles, assignments] = await Promise.all([getPrivilegedBuiltInRoles(), getRoleAssignments()]);
      const roleMap = new Map(roles.map(r => [r.id, r.displayName]));
      const users = await getUsersByIds(assignments.filter(a => roleMap.has(a.roleDefinitionId)).map(a => a.principalId));
      const usersById = new Map(users.map(u => [u.id, u]));
      const thresholdDays = 90;
      const staleUsers = [];
      for (const assignment of assignments) {
        if (!roleMap.has(assignment.roleDefinitionId)) continue;
        const user = usersById.get(assignment.principalId);
        if (!user) continue;
        const lastSignIn = user.signInActivity?.lastSignInDateTime || null;
        if (!lastSignIn) {
          staleUsers.push({ displayName: user.displayName, userPrincipalName: user.userPrincipalName, role: roleMap.get(assignment.roleDefinitionId), lastSignIn: null, daysSince: null });
          continue;
        }
        const daysSince = daysBetween(new Date().toISOString(), lastSignIn);
        if (daysSince > thresholdDays) staleUsers.push({ displayName: user.displayName, userPrincipalName: user.userPrincipalName, role: roleMap.get(assignment.roleDefinitionId), lastSignIn, daysSince });
      }
      if (!staleUsers.length) return pass(test, `No privileged accounts with stale sign-in activity (>${thresholdDays} days) found.`);
      return fail(test, `Found ${staleUsers.length} privileged accounts with stale or no sign-in activity.`, staleUsers.slice(0, 30));
    },

    '21855': async (test) => {
      const reviews = await getAccessReviewDefinitions();
      if (reviews === null) return skip(test, 'Access review definitions are not accessible (P2 license may be required).');
      const roleReviews = reviews.filter(r => {
        const scopeQuery = lower(r?.scope?.query || r?.instanceEnumerationScope?.query || '');
        return scopeQuery.includes('roleassignment') || scopeQuery.includes('roleeligibility');
      });
      if (roleReviews.length) return pass(test, `Found ${roleReviews.length} access review definition(s) scoped to privileged role assignments.`, roleReviews.map(r => ({ id: r.id, displayName: r.displayName, status: r.status })));
      return fail(test, 'No access review definitions were found scoped to privileged role assignments.');
    },

    '21857': async (test) => {
      const reviews = await getAccessReviewDefinitions();
      if (reviews === null) return skip(test, 'Access review definitions are not accessible (P2 license may be required).');
      const guestReviews = reviews.filter(r => {
        const scopeQuery = lower(r?.scope?.query || r?.instanceEnumerationScope?.query || '');
        return scopeQuery.includes('guest') || scopeQuery.includes("usertype eq 'guest'");
      });
      if (guestReviews.length) return pass(test, `Found ${guestReviews.length} access review definition(s) targeting guest identities.`, guestReviews.map(r => ({ id: r.id, displayName: r.displayName, status: r.status })));
      return fail(test, 'No access review definitions were found targeting guest identities.');
    },

    '21858': async (test) => {
      const guests = await getGuestUsers();
      const thresholdDays = 90;
      const inactive = guests.filter(g => {
        const lastSignIn = g.signInActivity?.lastSignInDateTime;
        if (!lastSignIn) return true;
        return daysBetween(new Date().toISOString(), lastSignIn) > thresholdDays;
      });
      if (!inactive.length) return pass(test, `No inactive guest identities found (${guests.length} guests checked, ${thresholdDays}-day threshold).`);
      return fail(test, `Found ${inactive.length} guest identities that are inactive or have never signed in.`, {
        total: guests.length,
        inactive: inactive.slice(0, 30).map(g => ({ displayName: g.displayName, userPrincipalName: g.userPrincipalName, lastSignIn: g.signInActivity?.lastSignInDateTime || null })),
      });
    },

    '21864': async (test) => {
      const detections = await getRiskDetections("riskState eq 'atRisk'");
      if (detections === null) return skip(test, 'Identity Protection risk detections are not available in this tenant (P2 license required).');
      if (!detections.length) return pass(test, 'No untriaged risk detections were found in Identity Protection.');
      return fail(test, `Found ${detections.length} untriaged risk detections that require review.`, detections.slice(0, 20).map(d => ({ userPrincipalName: d.userPrincipalName, riskEventType: d.riskEventType, riskLevel: d.riskLevel, detectedDateTime: d.detectedDateTime })));
    },

    '21865': async (test) => {
      const locations = await getNamedLocations();
      if (locations.length) return pass(test, `${locations.length} named location(s) are configured.`, locations.map(l => ({ id: l.id, displayName: l.displayName, type: l['@odata.type'] || '' })));
      return fail(test, 'No named locations are configured. Named locations are needed for network-based Conditional Access controls.');
    },

    '21867': async (test) => {
      const org = await getOrganization();
      const tenantId = org?.id;
      const sps = await Api.graphAll('servicePrincipals', {
        query: { '$filter': "servicePrincipalType eq 'Application'", '$select': 'id,displayName,appId,appOwnerOrganizationId', '$top': '100' },
        beta: true,
      });
      const tenantSps = tenantId ? sps.filter(sp => sp.appOwnerOrganizationId === tenantId) : sps;
      const sample = tenantSps.slice(0, 30);
      if (!sample.length) return skip(test, 'No tenant-owned enterprise applications found.');
      let withoutOwners = 0;
      const evidence = [];
      for (const sp of sample) {
        const owners = await getServicePrincipalOwners(sp.id);
        if (!owners.length) { withoutOwners++; evidence.push({ displayName: sp.displayName, appId: sp.appId }); }
      }
      if (!withoutOwners) return pass(test, `All ${sample.length} sampled enterprise applications have at least one owner.`);
      return fail(test, `Found ${withoutOwners}/${sample.length} sampled enterprise applications without owners.`, evidence.slice(0, 20));
    },

    '21868': async (test) => {
      const apps = await Api.graphAll('applications', { query: { '$select': 'id,displayName,appId', '$top': '50' } });
      if (!apps.length) return skip(test, 'No applications found.');
      const guestOwners = [];
      for (const app of apps.slice(0, 30)) {
        try {
          const owners = await Api.graphAll(`applications/${app.id}/owners`, { query: { '$select': 'id,displayName,userType', '$top': '10' } });
          for (const owner of owners.filter(o => lower(o.userType) === 'guest')) {
            guestOwners.push({ appDisplayName: app.displayName, appId: app.appId, guestDisplayName: owner.displayName, guestId: owner.id });
          }
        } catch { /* skip apps with inaccessible owners */ }
      }
      if (!guestOwners.length) return pass(test, `No guest owners found in the sampled ${Math.min(apps.length, 30)} applications.`);
      return fail(test, `Found ${guestOwners.length} application(s) with guest owners.`, guestOwners.slice(0, 20));
    },

    '21869': async (test) => {
      const org = await getOrganization();
      const tenantId = org?.id;
      const sps = await Api.graphAll('servicePrincipals', {
        query: { '$filter': "servicePrincipalType eq 'Application'", '$select': 'id,displayName,appId,appOwnerOrganizationId,appRoleAssignmentRequired', '$top': '200' },
        beta: true,
      });
      const tenantSps = tenantId ? sps.filter(sp => sp.appOwnerOrganizationId === tenantId) : sps;
      const noAssignment = tenantSps.filter(sp => !sp.appRoleAssignmentRequired);
      if (!noAssignment.length) return pass(test, `All ${tenantSps.length} tenant enterprise applications require explicit user assignment.`);
      return fail(test, `Found ${noAssignment.length} tenant enterprise applications that do not require explicit user assignment.`, noAssignment.slice(0, 30).map(sp => ({ displayName: sp.displayName, appId: sp.appId })));
    },

    '21874': async (test) => {
      const crossTenantPolicy = await getCrossTenantDefaultPolicy();
      const hasInboundUserRestrictions = crossTenantPolicy?.b2bCollaborationInbound?.usersAndGroups?.accessType;
      const hasInboundAppRestrictions = crossTenantPolicy?.b2bCollaborationInbound?.applications?.accessType;
      const isRestricted = (hasInboundUserRestrictions && lower(hasInboundUserRestrictions) !== 'allowed') ||
        (hasInboundAppRestrictions && lower(hasInboundAppRestrictions) !== 'allowed');
      if (isRestricted) return pass(test, 'Inbound cross-tenant collaboration access restrictions are configured.', {
        b2bCollaborationInbound: crossTenantPolicy?.b2bCollaborationInbound,
      });
      return fail(test, 'No domain-based allow or deny restrictions are configured for inbound cross-tenant collaboration.', {
        b2bCollaborationInbound: crossTenantPolicy?.b2bCollaborationInbound || null,
      });
    },

    '21876': async (test) => {
      const schedules = await getRoleEligibilitySchedules();
      if (schedules === null) return skip(test, 'PIM role eligibility schedules are not accessible (P2 license may be required).');
      const active = schedules.filter(s => lower(s.status) === 'provisioned' || lower(s.status) === 'active');
      if (active.length) return pass(test, `PIM is in use: found ${active.length} active role eligibility schedule(s).`, { count: active.length, total: schedules.length });
      if (schedules.length) return pass(test, `PIM is configured: found ${schedules.length} role eligibility schedule(s).`, { count: schedules.length });
      return fail(test, 'No PIM role eligibility schedules were found. Privileged roles may not be managed through PIM.');
    },

    '21883': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const workloadPolicies = policies.filter(policy => {
        const clientApps = toArray(policy?.conditions?.clientApplications?.includeServicePrincipals);
        const serviceRisk = toArray(policy?.conditions?.servicePrincipalRiskLevels);
        return lower(policy.state) === 'enabled' && clientApps.length > 0 && serviceRisk.length > 0;
      });
      if (workloadPolicies.length) return pass(test, 'Workload identities are covered by risk-based Conditional Access policies.', workloadPolicies.map(p => ({ id: p.id, displayName: p.displayName })));
      return fail(test, 'No enabled Conditional Access policies apply service principal risk conditions to workload identities.');
    },

    '21885': async (test) => {
      const apps = await Api.graphAll('applications', { query: { '$select': 'id,displayName,appId,web,publicClient,spa', '$top': '200' } });
      const dangerous = [];
      for (const app of apps) {
        const allUris = [...toArray(app.web?.redirectUris), ...toArray(app.spa?.redirectUris), ...toArray(app.publicClient?.redirectUris)];
        for (const uri of allUris) {
          if (/^http:\/\//i.test(uri) && !/localhost/i.test(uri) && !/127\.0\.0\.1/.test(uri) && !/::1/.test(uri)) {
            dangerous.push({ appDisplayName: app.displayName, appId: app.appId, uri });
          }
        }
      }
      if (!dangerous.length) return pass(test, `No insecure (HTTP non-local) redirect URIs found in ${apps.length} app registrations.`);
      return fail(test, `Found ${dangerous.length} insecure redirect URI(s) in app registrations.`, dangerous.slice(0, 30));
    },

    '21888': async (test) => {
      const apps = await Api.graphAll('applications', { query: { '$select': 'id,displayName,appId,web,spa,publicClient', '$top': '200' } });
      const insecure = [];
      for (const app of apps) {
        const allUris = [...toArray(app.web?.redirectUris), ...toArray(app.spa?.redirectUris), ...toArray(app.publicClient?.redirectUris)];
        for (const uri of allUris) {
          if (/^http:\/\//i.test(uri) && !/localhost/i.test(uri) && !/127\.0\.0\.1/.test(uri) && !/::1/.test(uri)) {
            insecure.push({ appDisplayName: app.displayName, appId: app.appId, uri });
          }
        }
      }
      if (!insecure.length) return pass(test, `No insecure or potentially dangling redirect URIs found in ${apps.length} app registrations.`);
      return fail(test, `Found ${insecure.length} potentially insecure redirect URI(s) that should be reviewed for dangling domain ownership.`, insecure.slice(0, 30));
    },

    '21899': async (test) => {
      return skip(test, 'This control is not fully implemented in the source module yet (under construction). Browser engine intentionally skips to avoid noisy/invalid Graph calls.');
    },

    '21953': async (test) => {
      const policy = await getDeviceRegistrationPolicy();
      if (policy?.localAdminPassword?.isEnabled === true) return pass(test, 'Cloud LAPS (Local Administrator Password Solution) is enabled in the device registration policy.');
      return fail(test, 'Cloud LAPS is not enabled in the device registration policy.', { localAdminPassword: policy?.localAdminPassword || null });
    },

    '21955': async (test) => {
      const policy = await getDeviceRegistrationPolicy();
      if (policy?.localAdminPassword?.isEnabled === true) return pass(test, 'Local administrator management via Cloud LAPS is configured for Entra joined devices.', { localAdminPassword: policy.localAdminPassword });
      return fail(test, 'Local administrator management is not configured for Microsoft Entra joined devices.', { localAdminPassword: policy?.localAdminPassword || null });
    },

    '21983': async (test) => {
      const recs = await getEntraRecommendations();
      if (recs === null) return skip(test, 'Entra recommendations are not available in this tenant configuration (license or permissions may be required).');
      const activeMedium = recs.filter(r => lower(r.status) === 'active' && lower(r.priority) === 'medium');
      if (!activeMedium.length) return pass(test, 'No active medium priority Entra recommendations found.', { total: recs.length });
      return fail(test, `Found ${activeMedium.length} active medium priority Entra recommendation(s).`, activeMedium.map(r => ({ id: r.id, displayName: r.displayName, priority: r.priority, status: r.status })));
    },

    '21984': async (test) => {
      const recs = await getEntraRecommendations();
      if (recs === null) return skip(test, 'Entra recommendations are not available in this tenant configuration.');
      const activeLow = recs.filter(r => lower(r.status) === 'active' && lower(r.priority) === 'low');
      if (!activeLow.length) return pass(test, 'No active low priority Entra recommendations found.', { total: recs.length });
      return fail(test, `Found ${activeLow.length} active low priority Entra recommendation(s).`, activeLow.map(r => ({ id: r.id, displayName: r.displayName, priority: r.priority, status: r.status })));
    },

    '22072': async (test) => {
      const policy = await getAuthenticationMethodsPolicy();
      const securityQ = toArray(policy?.authenticationMethodConfigurations).find(m => lower(`${m.id}`) === 'securityquestions');
      if (!securityQ || lower(securityQ.state) !== 'enabled') return pass(test, 'Security questions authentication method is not enabled.', { state: securityQ?.state || 'not configured' });
      return fail(test, 'Security questions (Q&A) authentication method is enabled, which is a weak SSPR mechanism.', { id: securityQ.id, state: securityQ.state });
    },

    '22124': async (test) => {
      const recs = await getEntraRecommendations();
      if (recs === null) return skip(test, 'Entra recommendations are not available in this tenant configuration.');
      const activeHigh = recs.filter(r => lower(r.status) === 'active' && lower(r.priority) === 'high');
      if (!activeHigh.length) return pass(test, 'No active high priority Entra recommendations found.', { total: recs.length });
      return fail(test, `Found ${activeHigh.length} active high priority Entra recommendation(s).`, activeHigh.map(r => ({ id: r.id, displayName: r.displayName, priority: r.priority, status: r.status })));
    },

    '22128': async (test) => {
      const [roles, assignments] = await Promise.all([getPrivilegedBuiltInRoles(), getRoleAssignments()]);
      const roleMap = new Map(roles.map(r => [r.id, r.displayName]));
      const users = await getUsersByIds(assignments.filter(a => roleMap.has(a.roleDefinitionId)).map(a => a.principalId));
      const usersById = new Map(users.map(u => [u.id, u]));
      const guestPriv = assignments
        .filter(a => roleMap.has(a.roleDefinitionId) && usersById.has(a.principalId) && lower(usersById.get(a.principalId).userType) === 'guest')
        .map(a => ({ displayName: usersById.get(a.principalId).displayName, userPrincipalName: usersById.get(a.principalId).userPrincipalName, roleName: roleMap.get(a.roleDefinitionId) }));
      if (!guestPriv.length) return pass(test, 'No guest identities found with privileged directory role assignments.');
      return fail(test, `Found ${guestPriv.length} guest identity(s) with privileged directory role assignments.`, guestPriv.slice(0, 20));
    },

    '23183': async (test) => {
      const sps = await Api.graphAll('servicePrincipals', {
        query: { '$select': 'id,displayName,appId,replyUrls,web', '$top': '200' },
        beta: true,
      });
      const dangerous = [];
      for (const sp of sps) {
        const allUris = [...toArray(sp.replyUrls), ...toArray(sp.web?.redirectUris)];
        for (const uri of allUris) {
          if (/^http:\/\//i.test(uri) && !/localhost/i.test(uri) && !/127\.0\.0\.1/.test(uri) && !/::1/.test(uri)) {
            dangerous.push({ spDisplayName: sp.displayName, appId: sp.appId, uri });
          }
        }
      }
      if (!dangerous.length) return pass(test, `No insecure redirect URIs found in ${sps.length} sampled service principals.`);
      return fail(test, `Found ${dangerous.length} insecure redirect URI(s) in service principals.`, dangerous.slice(0, 30));
    },

    '24541': async (test) => {
      const policies = await getDeviceCompliancePolicies();
      if (policies === null) return skip(test, 'Intune device compliance policies are not accessible (license or DeviceManagementConfiguration.Read.All permission required).');
      const win = policies.filter(p => lower(p['@odata.type']).includes('windows'));
      if (win.length) return pass(test, `Found ${win.length} Windows device compliance policy(s).`, win.map(p => ({ displayName: p.displayName, type: p['@odata.type'] })));
      return fail(test, 'No Windows device compliance policy was found in Intune.');
    },

    '24542': async (test) => {
      const policies = await getDeviceCompliancePolicies();
      if (policies === null) return skip(test, 'Intune device compliance policies are not accessible.');
      const mac = policies.filter(p => lower(p['@odata.type']).includes('mac'));
      if (mac.length) return pass(test, `Found ${mac.length} macOS device compliance policy(s).`, mac.map(p => ({ displayName: p.displayName, type: p['@odata.type'] })));
      return fail(test, 'No macOS device compliance policy was found in Intune.');
    },

    '24543': async (test) => {
      const policies = await getDeviceCompliancePolicies();
      if (policies === null) return skip(test, 'Intune device compliance policies are not accessible.');
      const ios = policies.filter(p => lower(p['@odata.type']).includes('ios'));
      if (ios.length) return pass(test, `Found ${ios.length} iOS device compliance policy(s).`, ios.map(p => ({ displayName: p.displayName, type: p['@odata.type'] })));
      return fail(test, 'No iOS device compliance policy was found in Intune.');
    },

    '24545': async (test) => {
      const policies = await getDeviceCompliancePolicies();
      if (policies === null) return skip(test, 'Intune device compliance policies are not accessible.');
      const android = policies.filter(p => lower(p['@odata.type']).includes('android') && (lower(p['@odata.type']).includes('enterprisefully') || lower(p['@odata.type']).includes('aosw') || lower(p.displayName).includes('android')));
      if (android.length) return pass(test, `Found ${android.length} Android Enterprise compliance policy(s).`, android.map(p => ({ displayName: p.displayName, type: p['@odata.type'] })));
      return fail(test, 'No Android Enterprise Fully Managed compliance policy was found in Intune.');
    },

    '24547': async (test) => {
      const policies = await getDeviceCompliancePolicies();
      if (policies === null) return skip(test, 'Intune device compliance policies are not accessible.');
      const android = policies.filter(p => lower(p['@odata.type']).includes('android') && (lower(p['@odata.type']).includes('workprofile') || lower(p['@odata.type']).includes('personally')));
      if (android.length) return pass(test, `Found ${android.length} Android Personal Work Profile compliance policy(s).`, android.map(p => ({ displayName: p.displayName, type: p['@odata.type'] })));
      return fail(test, 'No Android Enterprise Personally-Owned Work Profile compliance policy was found in Intune.');
    },

    '24548': async (test) => {
      const policies = await getManagedAppPolicies();
      if (policies === null) return skip(test, 'Intune managed app policies are not accessible (license or DeviceManagementApps.Read.All permission required).');
      const ios = policies.filter(p => lower(p['@odata.type']).includes('ios'));
      if (ios.length) return pass(test, `Found ${ios.length} iOS app protection policy(s).`, ios.map(p => ({ displayName: p.displayName, type: p['@odata.type'] })));
      return fail(test, 'No iOS app protection policy was found in Intune.');
    },

    '24549': async (test) => {
      const policies = await getManagedAppPolicies();
      if (policies === null) return skip(test, 'Intune managed app policies are not accessible.');
      const android = policies.filter(p => lower(p['@odata.type']).includes('android'));
      if (android.length) return pass(test, `Found ${android.length} Android app protection policy(s).`, android.map(p => ({ displayName: p.displayName, type: p['@odata.type'] })));
      return fail(test, 'No Android app protection policy was found in Intune.');
    },

    '24550': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('bitlocker') || /bitlocker|disk encryption/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /bitlocker|disk encryption/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} BitLocker-related configuration policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName, type: 'deviceConfiguration' })), ...inPolicies.map(p => ({ displayName: p.name, type: 'configurationPolicy' }))]);
      return fail(test, 'No BitLocker device configuration policy was found in Intune.');
    },

    '24551': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('windowsidentityprotection') || /whfb|hello for business/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /hello for business|whfb|identity protection/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} Windows Hello for Business policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName, type: 'deviceConfiguration' })), ...inPolicies.map(p => ({ displayName: p.name, type: 'configurationPolicy' }))]);
      return fail(test, 'No Windows Hello for Business policy was found in Intune.');
    },

    '24553': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('windowsupdate') || /windows update|wu ring/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /windows update|update ring/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} Windows Update policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName, type: 'deviceConfiguration' })), ...inPolicies.map(p => ({ displayName: p.name, type: 'configurationPolicy' }))]);
      return fail(test, 'No Windows Update policy was found in Intune.');
    },

    '24560': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => /laps|local admin password/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /laps|local admin password/i.test(p.name || ''));
      const drpCheck = await getDeviceRegistrationPolicy();
      const drpLaps = drpCheck?.localAdminPassword?.isEnabled === true;
      const found = inConfigs.length + inPolicies.length;
      if (found || drpLaps) return pass(test, `Cloud LAPS is configured. ${found} Intune policy(s) found${drpLaps ? ', enabled in device registration policy' : ''}.`);
      return fail(test, 'No Windows Cloud LAPS policy was found in Intune or the device registration policy.');
    },

    '24575': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('defender') || lower(c['@odata.type']).includes('antivirus') || /defender|antivirus/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /defender|antivirus/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} Windows Defender/Antivirus policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName, type: 'deviceConfiguration' })), ...inPolicies.map(p => ({ displayName: p.name, type: 'configurationPolicy' }))]);
      return fail(test, 'No Windows Defender Antivirus policy was found in Intune.');
    },

    '21795': async (test, ctx) => {
      ctx?.report('Querying sign-in logs for last 30 days (this can take longer)...', 10);
      const legacyClients = ['exchange activesync', 'imap4', 'mapi over http', 'smtp auth', 'pop3', 'other clients'];
      const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
      const signIns = await getSignInLogs(`createdDateTime ge ${thirtyDaysAgo}T00:00:00Z and status/errorCode eq 0`, 100, { timeoutMs: 120000 });
      ctx?.report('Processing sign-in dataset...', 75);
      if (signIns === null) return skip(test, 'Sign-in logs are not accessible (AuditLog.Read.All required, or sign-in logs may not be available in this tenant tier).');
      const legacy = signIns.filter(s => legacyClients.some(c => lower(s.clientAppUsed || '').includes(c)));
      if (!legacy.length) return pass(test, `No legacy authentication sign-ins found in the last 30 days (${signIns.length} sign-ins sampled).`);
      return fail(test, `Found ${legacy.length} legacy authentication sign-in(s) in the last 30 days.`, legacy.slice(0, 20).map(s => ({ userPrincipalName: s.userPrincipalName, clientAppUsed: s.clientAppUsed, createdDateTime: s.createdDateTime })));
    },

    '21780': async (test) => {
      const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
      const signIns = await getSignInLogs(`createdDateTime ge ${thirtyDaysAgo}T00:00:00Z and status/errorCode eq 0`, 100);
      if (signIns === null) return skip(test, 'Sign-in logs are not accessible (AuditLog.Read.All required).');
      const adalSignIns = signIns.filter(s => {
        const lib = lower(s.authenticationLibrary?.name || s.authenticationLibrary || '');
        return lib.includes('adal') || lib.includes('azure active directory authentication library');
      });
      if (!adalSignIns.length) return pass(test, `No ADAL-based sign-ins detected in the last 30 days (${signIns.length} sign-ins sampled).`);
      return fail(test, `Found ${adalSignIns.length} sign-in(s) using ADAL (deprecated authentication library) in the last 30 days.`, adalSignIns.slice(0, 20).map(s => ({ userPrincipalName: s.userPrincipalName, appDisplayName: s.appDisplayName, authenticationLibrary: s.authenticationLibrary })));
    },

    '21800': async (test, ctx) => {
      ctx?.report('Querying sign-in authentication requirement logs...', 10);
      const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
      const signIns = await getSignInLogs(`createdDateTime ge ${thirtyDaysAgo}T00:00:00Z and status/errorCode eq 0`, 100, { timeoutMs: 120000 });
      ctx?.report('Calculating single-factor share...', 75);
      if (signIns === null) return skip(test, 'Sign-in logs are not accessible (AuditLog.Read.All required).');
      const weakAuth = signIns.filter(s => lower(s.authenticationRequirement) === 'singlefactorauthentication');
      const total = signIns.length;
      if (!weakAuth.length) return pass(test, `All ${total} sampled successful sign-ins in the last 30 days used multi-factor or strong authentication.`);
      const pct = Math.round((weakAuth.length / total) * 100);
      return fail(test, `${weakAuth.length} of ${total} sampled sign-ins (${pct}%) in the last 30 days used only single-factor authentication.`, weakAuth.slice(0, 20).map(s => ({ userPrincipalName: s.userPrincipalName, appDisplayName: s.appDisplayName, authenticationRequirement: s.authenticationRequirement, createdDateTime: s.createdDateTime })));
    },

    '21801': async (test) => {
      const users = await getUserRegistrationDetails(500);
      if (users === null) return skip(test, 'User registration details report is not accessible (Reports.Read.All required).');
      const notCapable = users.filter(u => !u.isMfaCapable && !lower(u.userPrincipalName).includes('#ext#'));
      const total = users.length;
      if (!notCapable.length) return pass(test, `All ${total} users have at least one strong authentication method registered.`);
      return fail(test, `${notCapable.length} of ${total} users do not have a strong authentication method registered.`, { total, notCapable: notCapable.slice(0, 30).map(u => ({ userPrincipalName: u.userPrincipalName, methodsRegistered: u.methodsRegistered })) });
    },

    '21810': async (test) => {
      const policy = await getAuthorizationPolicy();
      const permissionPolicies = toArray(policy?.permissionGrantPoliciesAssigned);
      const hasManageAppConsent = permissionPolicies.some(p => lower(p).includes('managepermisgrantpolicies'));
      const noRSC = permissionPolicies.filter(p => lower(p).includes('resourcespecific')).length === 0;
      const isDefault = permissionPolicies.some(p => lower(p) === 'managepermisgrantpolicies_microsoft-user-default-legacy');
      if (noRSC && !isDefault) return pass(test, 'Resource-specific consent appears restricted.', { permissionGrantPoliciesAssigned: permissionPolicies });
      const hasLimitedConsent = policy?.allowedToSignUpEmailBasedSubscriptions === false || permissionPolicies.some(p => lower(p).includes('low') || lower(p).includes('limited'));
      // check if RSC is enabled
      const rscPermissions = permissionPolicies.filter(p => lower(p).includes('resource'));
      if (!rscPermissions.length) return pass(test, 'No resource-specific consent policies are assigned.', { permissionGrantPoliciesAssigned: permissionPolicies });
      return fail(test, 'Resource-specific consent policies are assigned that may allow over-permissive application consent.', { permissionGrantPoliciesAssigned: permissionPolicies });
    },

    '21829': async (test) => {
      const payload = await graphWithFallback('domains', [
        { query: { '$top': '50' } },
        { query: { '$select': 'id,authenticationType,isDefault,isManaged,isInitial', '$top': '50' } },
      ]);
      const domains = toArray(payload?.value);
      const managedDomains = domains.filter(d => !d.isInitial && !lower(d.id).endsWith('.onmicrosoft.com'));
      const domainsToCheck = managedDomains.length ? managedDomains : domains;
      const federated = domainsToCheck.filter(d => lower(d.authenticationType) === 'federated');
      if (!federated.length) return pass(test, 'All custom domains use cloud authentication (Managed) — no federated domains found.', domainsToCheck.map(d => ({ id: d.id, authenticationType: d.authenticationType })));
      return fail(test, `Found ${federated.length} federated domain(s). Verify that Entra ID's cloud authentication stack is in use (not AD FS pass-through).`, federated.map(d => ({ id: d.id, authenticationType: d.authenticationType })));
    },

    '21831': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const protectedActionPolicies = policies.filter(policy =>
        lower(policy.state) === 'enabled' &&
        toArray(policy?.conditions?.applications?.includeAuthenticationContextClassReferences).length > 0
      );
      if (protectedActionPolicies.length) return pass(test, `Found ${protectedActionPolicies.length} Conditional Access policy(s) using authentication context references (protected actions pattern).`, protectedActionPolicies.map(p => ({ id: p.id, displayName: p.displayName, authContextRefs: p.conditions?.applications?.includeAuthenticationContextClassReferences })));
      return fail(test, 'No enabled Conditional Access policies use authentication context references. Conditional Access protected actions may not be configured.');
    },

    '21835': async (test) => {
      const [roles, assignments] = await Promise.all([getPrivilegedBuiltInRoles(), getRoleAssignments()]);
      const gaRole = roles.find(r => /^global administrator$/i.test(r.displayName));
      if (!gaRole) return skip(test, 'Global Administrator role not found.');
      const gaAssignees = assignments.filter(a => a.roleDefinitionId === gaRole.id).map(a => a.principalId);
      const gaUsers = await getUsersByIds(gaAssignees);
      const enabledPolicies = (await getConditionalAccessPolicies()).filter(p => lower(p.state) === 'enabled');
      const bgPattern = /break.?glass|emergency|bg-|emerg|e\.?acc/i;
      const nameMatch = gaUsers.filter(u => bgPattern.test(u.displayName || '') || bgPattern.test(u.userPrincipalName || ''));
      // Accounts excluded from all enabled CA policies (classic breakglass indicator)
      const excludedFromAll = gaUsers.filter(u =>
        enabledPolicies.length > 0 &&
        enabledPolicies.every(p => toArray(p?.conditions?.users?.excludeUsers).includes(u.id))
      );
      const cloudOnly = gaUsers.filter(u => !u.onPremisesSyncEnabled);
      if (nameMatch.length) return pass(test, `Found ${nameMatch.length} emergency/breakglass Global Administrator account(s) by naming pattern.`, nameMatch.map(u => ({ displayName: u.displayName, userPrincipalName: u.userPrincipalName })));
      if (excludedFromAll.length) return pass(test, `Found ${excludedFromAll.length} Global Administrator account(s) excluded from all enabled Conditional Access policies — consistent with breakglass pattern.`, excludedFromAll.map(u => ({ displayName: u.displayName, userPrincipalName: u.userPrincipalName })));
      return fail(test, 'No emergency/breakglass accounts were detected by naming pattern or CA policy exclusions. Verify that at least one cloud-only emergency Global Administrator account exists and is excluded from Conditional Access policies.', { gaCount: gaUsers.length, cloudOnly: cloudOnly.length });
    },

    '21847': async (test) => {
      const settings = await getDirectorySettings();
      const pwdSettings = settings.find(s => lower(s.displayName) === 'passwordrulesettings' || s.templateId === '5cf42378-d67d-4f36-ba46-e8b86229381d');
      if (!pwdSettings) return skip(test, 'Password rule settings not found in directory settings. On-premises password protection may be managed outside the tenant portal.');
      const onPremEnabled = toArray(pwdSettings.values).find(v => lower(v.name) === 'enablebannedpasswordonpremises')?.value;
      const mode = toArray(pwdSettings.values).find(v => lower(v.name) === 'bannedpasswordcheckonpremisesmode')?.value;
      if (lower(onPremEnabled) === 'true') return pass(test, 'On-premises password protection is enabled.', { enableBannedPasswordOnPremises: onPremEnabled, mode });
      return fail(test, 'On-premises password protection does not appear to be enabled in directory settings.', { enableBannedPasswordOnPremises: onPremEnabled || null, mode: mode || null });
    },

    '21848': async (test) => {
      const settings = await getDirectorySettings();
      const pwdSettings = settings.find(s => lower(s.displayName) === 'passwordrulesettings' || s.templateId === '5cf42378-d67d-4f36-ba46-e8b86229381d');
      if (!pwdSettings) return skip(test, 'Password rule settings not found in directory settings.');
      const bannedList = toArray(pwdSettings.values).find(v => lower(v.name) === 'bannedpasswordlist')?.value;
      if (bannedList && bannedList.trim().length > 0) return pass(test, 'Custom banned password list is configured.', { bannedPasswordList: bannedList.slice(0, 200) });
      return fail(test, 'No custom banned password list is configured in tenant directory settings.');
    },

    '21866': async (test) => {
      const recs = await getEntraRecommendations();
      if (recs === null) return skip(test, 'Entra recommendations are not available in this tenant configuration.');
      const active = recs.filter(r => lower(r.status) === 'active');
      if (!active.length) return pass(test, `All ${recs.length} Entra recommendations have been addressed (no active items).`);
      return fail(test, `Found ${active.length} active (unaddressed) Entra recommendation(s).`, active.map(r => ({ id: r.id, displayName: r.displayName, priority: r.priority, status: r.status })));
    },

    '21870': async (test) => {
      const policy = await getSelfServicePasswordResetPolicy();
      if (policy === null) return skip(test, 'Self-service password reset policy is not accessible (requires tenant-level permissions).');
      if (policy.isEnabled) return pass(test, 'Self-service password reset (SSPR) is enabled.', {
        isEnabled: policy.isEnabled,
        allowedToUseSspr: policy.allowedToUseSspr,
        registrationEnforcement: policy.registrationEnforcement,
      });
      return fail(test, 'Self-service password reset (SSPR) is not enabled.', { isEnabled: policy.isEnabled });
    },

    '21877': async (test) => {
      const guests = await getGuestUsers();
      if (!guests.length) return pass(test, 'No guest users found in the tenant.');
      const noSponsor = [];
      for (const guest of guests.slice(0, 50)) {
        try {
          const sponsors = await Api.graphAll(`users/${guest.id}/sponsors`, { query: { '$select': 'id,displayName', '$top': '5' } });
          if (!sponsors.length) noSponsor.push({ displayName: guest.displayName, userPrincipalName: guest.userPrincipalName });
        } catch {
          noSponsor.push({ displayName: guest.displayName, userPrincipalName: guest.userPrincipalName });
        }
      }
      if (!noSponsor.length) return pass(test, `All ${Math.min(guests.length, 50)} sampled guest accounts have at least one sponsor.`);
      return fail(test, `${noSponsor.length} of ${Math.min(guests.length, 50)} sampled guest accounts have no sponsor.`, noSponsor.slice(0, 20));
    },

    '21878': async (test) => {
      const policies = await getEntitlementAssignmentPolicies();
      if (policies === null) return skip(test, 'Entitlement management assignment policies are not accessible (EntitlementManagement.Read.All required).');
      if (!policies.length) return skip(test, 'No entitlement management assignment policies found in this tenant.');
      const noExpiry = policies.filter(p => !p.expiration || lower(p.expiration?.type) === 'noexpiration' || (!p.expiration?.duration && !p.expiration?.endDateTime));
      if (!noExpiry.length) return pass(test, `All ${policies.length} entitlement management assignment policies have an expiration configured.`);
      return fail(test, `${noExpiry.length} of ${policies.length} assignment policies have no expiration configured.`, noExpiry.map(p => ({ id: p.id, displayName: p.displayName, expiration: p.expiration })));
    },

    '21879': async (test) => {
      const policies = await getEntitlementAssignmentPolicies();
      if (policies === null) return skip(test, 'Entitlement management assignment policies are not accessible (EntitlementManagement.Read.All required).');
      if (!policies.length) return skip(test, 'No entitlement management assignment policies found in this tenant.');
      const externalPolicies = policies.filter(p => lower(p.allowedTargetScope).includes('allexternalusers') || lower(p.allowedTargetScope).includes('specificconnectedorganizationusers'));
      if (!externalPolicies.length) return pass(test, 'No entitlement management assignment policies apply to external users.', { total: policies.length });
      const noApproval = externalPolicies.filter(p => !p.requestApprovalSettings?.isApprovalRequired);
      if (!noApproval.length) return pass(test, `All ${externalPolicies.length} external-facing assignment policies require approval.`);
      return fail(test, `${noApproval.length} of ${externalPolicies.length} external-facing assignment policies do not require approval.`, noApproval.map(p => ({ id: p.id, displayName: p.displayName, allowedTargetScope: p.allowedTargetScope })));
    },

    '21875': async (test) => {
      const orgs = await getEntitlementConnectedOrgs();
      if (orgs === null) return skip(test, 'Entitlement management connected organizations are not accessible (EntitlementManagement.Read.All required).');
      if (orgs.length) return pass(test, `Found ${orgs.length} connected organization(s) in entitlement management.`, orgs.map(o => ({ displayName: o.displayName, state: o.state })));
      return fail(test, 'No connected organizations found in entitlement management. External collaboration may not be governed through entitlement management.');
    },

    '21929': async (test) => {
      const policies = await getEntitlementAssignmentPolicies();
      if (policies === null) return skip(test, 'Entitlement management assignment policies are not accessible.');
      const guestPolicies = policies.filter(p => lower(p.allowedTargetScope).includes('allexternalusers') || lower(p.allowedTargetScope).includes('specificconnectedorganizationusers'));
      if (!guestPolicies.length) return pass(test, 'No entitlement management assignment policies target guest or external users.');
      const noExpiryOrReview = guestPolicies.filter(p => {
        const hasExpiry = p.expiration && lower(p.expiration?.type) !== 'noexpiration' && (p.expiration?.duration || p.expiration?.endDateTime);
        const hasReview = p.accessReviewSettings?.isEnabled === true;
        return !hasExpiry && !hasReview;
      });
      if (!noExpiryOrReview.length) return pass(test, `All ${guestPolicies.length} guest-facing entitlement management policies have expiration or access reviews configured.`);
      return fail(test, `${noExpiryOrReview.length} guest-facing entitlement management policies lack both expiration and access reviews.`, noExpiryOrReview.map(p => ({ id: p.id, displayName: p.displayName, allowedTargetScope: p.allowedTargetScope })));
    },

    '21890': async (test) => {
      const policy = await getSelfServicePasswordResetPolicy();
      if (policy === null) return skip(test, 'Self-service password reset policy is not accessible.');
      if (!policy.isEnabled) return skip(test, 'SSPR is not enabled — notification settings are irrelevant.');
      return skip(test, 'SSPR user notification settings are not exposed via Microsoft Graph in delegated browser mode. Validate this in Entra Admin Center.');
    },

    '21891': async (test) => {
      const policy = await getSelfServicePasswordResetPolicy();
      if (policy === null) return skip(test, 'Self-service password reset policy is not accessible.');
      if (!policy.isEnabled) return skip(test, 'SSPR is not enabled — notification settings are irrelevant.');
      return skip(test, 'SSPR admin notification settings are not exposed via Microsoft Graph in delegated browser mode. Validate this in Entra Admin Center.');
    },

    '21893': async (test) => {
      const mfaPolicy = await getIdentityProtectionMfaPolicy();
      if (mfaPolicy === null) {
        // Fallback: check security defaults
        const defaults = await Api.graph('policies/identitySecurityDefaultsEnforcementPolicy', { query: { '$select': 'isEnabled' } }).catch(() => null);
        if (defaults?.isEnabled) return pass(test, 'Security defaults are enabled, which enforces MFA registration for all users.');
        return skip(test, 'Identity Protection MFA registration policy is not accessible (P2 license required) and security defaults are not enabled.');
      }
      const state = lower(mfaPolicy?.state || '');
      if (state === 'enabled') return pass(test, 'Identity Protection MFA registration policy is enabled.', { state: mfaPolicy.state, includedUsers: mfaPolicy.includedUsers, excludedUsers: mfaPolicy.excludedUsers });
      return fail(test, `Identity Protection MFA registration policy is not enabled (state: ${mfaPolicy.state || 'unknown'}).`, { state: mfaPolicy.state, includedUsers: mfaPolicy.includedUsers });
    },

    '21798': async (test) => {
      const notifPolicy = await getIdentityProtectionNotificationPolicy();
      if (notifPolicy !== null) {
        const alerts = toArray(notifPolicy?.value || [notifPolicy]).filter(Boolean);
        const hasEnabled = alerts.some(n => n?.isEnabled !== false && n?.state !== 'disabled');
        if (hasEnabled) return pass(test, 'Identity Protection notifications appear to be configured.', alerts.slice(0, 5));
        return fail(test, 'Identity Protection notifications do not appear to be enabled.', alerts.slice(0, 5));
      }
      const riskDetections = await getRiskDetections("riskState eq 'atRisk'");
      if (riskDetections === null) return skip(test, 'Identity Protection notification configuration is not exposed by Graph, and Identity Protection data is inaccessible (license or permission limitation).');
      return skip(test, 'Identity Protection notification configuration is not exposed by Graph. Identity Protection data is accessible; validate notification recipients in Entra Admin Center.');
    },

    '21818': async (test, ctx) => {
      const notificationRuleIds = [
        'Notification_Admin_Eligibility',
        'Notification_Requestor_Admin_Eligibility',
        'Notification_Approver_Admin_Eligibility',
        'Notification_Admin_Admin_Assignment',
        'Notification_Requestor_Admin_Assignment',
        'Notification_Approver_Admin_Assignment',
        'Notification_Admin_EndUser_Assignment',
        'Notification_Requestor_EndUser_Assignment',
        'Notification_Approver_EndUser_Assignment',
      ];

      ctx?.report('Loading privileged directory role definitions...', 10);
      const roleDefs = await Api.graphAll('roleManagement/directory/roleDefinitions', {
        beta: true,
        query: {
          '$select': 'id,displayName,isBuiltIn,isPrivileged',
          '$filter': 'isPrivileged eq true',
          '$top': '999',
        },
      });
      const privilegedRoles = toArray(roleDefs).filter(r => r.isPrivileged);
      if (!privilegedRoles.length) return skip(test, 'No privileged role definitions could be enumerated.');

      ctx?.report('Loading PIM role policy assignments...', 30);
      const assignments = await getRoleManagementPolicyAssignments("scopeId eq '/' and scopeType eq 'DirectoryRole'", 200);
      if (assignments === null) return skip(test, 'PIM role management policy assignments are not accessible (P2 PIM license/provider/permission limitation).');

      const policyByRole = new Map(toArray(assignments).map(a => [a.roleDefinitionId, a.policyId]).filter(x => x[0] && x[1]));
      const roleSample = privilegedRoles.slice(0, 40);
      const findings = [];

      for (let i = 0; i < roleSample.length; i++) {
        const role = roleSample[i];
        const pct = 35 + Math.round((60 * (i + 1)) / roleSample.length);
        ctx?.report(`Checking notification rules for ${role.displayName}...`, pct);
        const policyId = policyByRole.get(role.id);
        if (!policyId) {
          findings.push({ roleDisplayName: role.displayName, issue: 'No PIM policy assignment found' });
          continue;
        }
        for (const ruleId of notificationRuleIds) {
          const rule = await getRoleManagementPolicyRule(policyId, ruleId);
          if (!rule) continue;
          const hasRecipients = toArray(rule.notificationRecipients).length > 0;
          const hasDefault = rule.isDefaultRecipientsEnabled === true;
          if (!hasDefault && !hasRecipients) {
            findings.push({ roleDisplayName: role.displayName, ruleId, issue: 'No default recipients and no additional recipients' });
            return fail(test, 'Privileged role activation notifications are not fully configured.', findings);
          }
        }
      }

      return pass(test, `Privileged role activation notifications are configured for sampled privileged roles (${roleSample.length}).`);
    },

    '21819': async (test, ctx) => {
      const gaTemplateId = '62e90394-69f5-4237-9190-012177145e10';
      ctx?.report('Loading Global Administrator role definition...', 15);
      const roleDefs = await Api.graphAll('roleManagement/directory/roleDefinitions', {
        beta: true,
        query: {
          '$select': 'id,displayName,templateId',
          '$filter': `templateId eq '${gaTemplateId}'`,
          '$top': '5',
        },
      });
      const gaRole = toArray(roleDefs)[0] || null;
      if (!gaRole) return fail(test, 'Could not find Global Administrator role definition.');

      ctx?.report('Loading Global Administrator policy assignment...', 40);
      const assignments = await getRoleManagementPolicyAssignments(`scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '${gaRole.id}'`, 10);
      if (assignments === null) return skip(test, 'PIM role management policy assignments are not accessible (P2 PIM license/provider/permission limitation).');
      const policyId = toArray(assignments)[0]?.policyId || null;
      if (!policyId) return fail(test, 'No PIM policy assignment found for Global Administrator role.');

      ctx?.report('Reading activation notification rule...', 75);
      const rule = await getRoleManagementPolicyRule(policyId, 'Notification_Admin_EndUser_Assignment');
      if (!rule) return fail(test, 'Could not read Global Administrator activation notification rule.');
      const hasDefault = rule.isDefaultRecipientsEnabled === true;
      const recipients = toArray(rule.notificationRecipients);
      if (hasDefault || recipients.length > 0) {
        return pass(test, 'Activation alerts are configured for Global Administrator role assignments.', {
          roleDisplayName: gaRole.displayName,
          isDefaultRecipientsEnabled: hasDefault,
          notificationRecipients: recipients,
        });
      }
      return fail(test, 'Activation alerts are not configured for Global Administrator role assignments.', {
        roleDisplayName: gaRole.displayName,
        isDefaultRecipientsEnabled: hasDefault,
        notificationRecipients: recipients,
      });
    },

    '21820': async (test, ctx) => {
      ctx?.report('Loading privileged roles and PIM assignments...', 20);
      const [roleDefs, assignments] = await Promise.all([
        Api.graphAll('roleManagement/directory/roleDefinitions', {
          beta: true,
          query: {
            '$select': 'id,displayName,isPrivileged',
            '$filter': 'isPrivileged eq true',
            '$top': '999',
          },
        }),
        getRoleManagementPolicyAssignments("scopeId eq '/' and scopeType eq 'DirectoryRole'", 300),
      ]);
      if (assignments === null) return skip(test, 'PIM role management policy assignments are not accessible (P2 PIM license/provider/permission limitation).');

      const privilegedRoles = toArray(roleDefs).filter(r => r.isPrivileged);
      if (!privilegedRoles.length) return skip(test, 'No privileged roles found for evaluation.');

      const policyByRole = new Map(toArray(assignments).map(a => [a.roleDefinitionId, a.policyId]).filter(x => x[0] && x[1]));
      const rolesWithIssues = [];
      for (let i = 0; i < privilegedRoles.length; i++) {
        const role = privilegedRoles[i];
        const pct = 30 + Math.round((65 * (i + 1)) / privilegedRoles.length);
        ctx?.report(`Checking activation notification for ${role.displayName}...`, pct);

        const policyId = policyByRole.get(role.id);
        if (!policyId) {
          rolesWithIssues.push({ roleDisplayName: role.displayName, issue: 'No PIM policy assignment found' });
          continue;
        }

        const rule = await getRoleManagementPolicyRule(policyId, 'Notification_Admin_EndUser_Assignment');
        if (!rule) {
          rolesWithIssues.push({ roleDisplayName: role.displayName, issue: 'Activation notification rule not found' });
          continue;
        }

        const hasDefault = rule.isDefaultRecipientsEnabled === true;
        const hasRecipients = toArray(rule.notificationRecipients).length > 0;
        if (!hasDefault && !hasRecipients) {
          rolesWithIssues.push({ roleDisplayName: role.displayName, issue: 'No default recipients and no additional recipients' });
        }
      }

      if (!rolesWithIssues.length) return pass(test, `Activation alerts are configured for all evaluated privileged role assignments (${privilegedRoles.length}).`);
      return fail(test, 'Activation alerts are missing or improperly configured for one or more privileged role assignments.', rolesWithIssues.slice(0, 30));
    },

    '21896': async (test) => {
      const sps = await Api.graphAll('servicePrincipals', {
        query: {
          '$filter': "servicePrincipalType eq 'Application'",
          '$select': 'id,displayName,appId,keyCredentials,passwordCredentials,appOwnerOrganizationId',
          '$top': '200',
        },
        beta: true,
      });
      const org = await getOrganization();
      const tenantId = org?.id;
      const tenantSps = tenantId ? sps.filter(sp => sp.appOwnerOrganizationId === tenantId) : sps;
      const withCreds = tenantSps.filter(sp => toArray(sp.keyCredentials).length > 0 || toArray(sp.passwordCredentials).length > 0);
      if (!withCreds.length) return pass(test, 'No tenant-owned service principals (enterprise apps) have key or password credentials stored directly.');
      return fail(test, `Found ${withCreds.length} service principal(s) with credentials. Managed Identity or federated identity credentials are preferred over stored secrets/certificates.`, withCreds.slice(0, 20).map(sp => ({ displayName: sp.displayName, appId: sp.appId, keys: toArray(sp.keyCredentials).length, passwords: toArray(sp.passwordCredentials).length })));
    },

    '21992': async (test) => {
      const apps = await Api.graphAll('applications', { query: { '$select': 'id,displayName,appId,keyCredentials', '$top': '200' } });
      const now = new Date();
      const ninetyDays = 90 * 24 * 60 * 60 * 1000;
      const expiredOrSoon = [];
      for (const app of apps) {
        for (const key of toArray(app.keyCredentials)) {
          const expiry = key.endDateTime ? new Date(key.endDateTime) : null;
          if (!expiry) continue;
          const msTillExpiry = expiry.getTime() - now.getTime();
          if (msTillExpiry < ninetyDays) {
            expiredOrSoon.push({ appDisplayName: app.displayName, appId: app.appId, keyDisplayName: key.displayName || key.customKeyIdentifier, endDateTime: key.endDateTime, daysLeft: Math.floor(msTillExpiry / 86400000) });
          }
        }
      }
      if (!expiredOrSoon.length) return pass(test, `All application certificates in ${apps.length} app registrations are valid for more than 90 days.`);
      const expired = expiredOrSoon.filter(c => c.daysLeft < 0).length;
      return fail(test, `Found ${expiredOrSoon.length} certificate(s) expiring within 90 days or already expired (${expired} expired).`, expiredOrSoon.slice(0, 20));
    },

    '21789': async (test) => {
      const audits = await getDirectoryAuditLogs("activityDisplayName eq 'Add company'", 25);
      if (audits === null) return skip(test, 'Directory audit logs are not accessible (AuditLog.Read.All required).');
      if (!audits.length) return pass(test, 'No tenant creation events found in audit logs.');
      // These are unusual events — if found and recent, flag for review
      const recentCreations = audits.filter(a => {
        const daysSince = daysBetween(new Date().toISOString(), a.activityDateTime);
        return daysSince < 365;
      });
      if (!recentCreations.length) return pass(test, `Found ${audits.length} tenant creation event(s) but none in the last 365 days.`, audits.slice(0, 5).map(a => ({ activityDisplayName: a.activityDisplayName, activityDateTime: a.activityDateTime, result: a.result })));
      return fail(test, `Found ${recentCreations.length} recent tenant creation event(s) — verify these are authorized.`, recentCreations.slice(0, 10).map(a => ({ activityDisplayName: a.activityDisplayName, activityDateTime: a.activityDateTime, initiatedBy: a.initiatedBy })));
    },

    '21860': async (test) => {
      const settings = await getEntraDiagnosticSettings();
      if (settings === null) return skip(test, 'Cannot access Entra diagnostic settings via ARM API (missing Azure scope consent or insufficient Azure RBAC permissions).');
      if (!settings.length) return fail(test, 'No Entra diagnostic settings are configured.');

      const requiredLogs = [
        'AuditLogs',
        'SignInLogs',
        'NonInteractiveUserSignInLogs',
        'ServicePrincipalSignInLogs',
        'ManagedIdentitySignInLogs',
        'ProvisioningLogs',
        'ADFSSignInLogs',
        'RiskyUsers',
        'UserRiskEvents',
        'NetworkAccessTrafficLogs',
        'RiskyServicePrincipals',
        'ServicePrincipalRiskEvents',
        'EnrichedOffice365AuditLogs',
        'MicrosoftGraphActivityLogs',
        'RemoteNetworkHealthLogs',
      ];

      const enabledLogs = new Set();
      for (const setting of settings) {
        for (const logCfg of toArray(setting?.properties?.logs)) {
          if (logCfg?.enabled === true && logCfg?.category) enabledLogs.add(`${logCfg.category}`);
        }
      }

      const missingLogs = requiredLogs.filter(name => !enabledLogs.has(name));
      if (!missingLogs.length) {
        return pass(test, 'Diagnostic settings are configured for all expected Microsoft Entra log categories.', {
          diagnosticSettings: settings.map(s => s.name),
          enabledCategoryCount: enabledLogs.size,
        });
      }

      return fail(test, `${missingLogs.length} required Entra log categories are not enabled in diagnostic settings.`, {
        missingLogs,
        diagnosticSettings: settings.map(s => s.name),
      });
    },

    '22098': async (test) => {
      return skip(test, 'Integration of Entra audit logs with Azure Monitor cannot be verified via Microsoft Graph. Check Azure Portal -> Entra ID -> Monitoring -> Diagnostic settings to confirm audit logs are forwarded to a Log Analytics workspace.', { note: 'Azure Monitor integration requires ARM API access, not available in this browser-only tool.' });
    },

    '22099': async (test) => {
      return skip(test, 'Integration of Entra sign-in logs with Azure Monitor cannot be verified via Microsoft Graph. Check Azure Portal -> Entra ID -> Monitoring -> Diagnostic settings to confirm sign-in logs are forwarded to a Log Analytics workspace.', { note: 'Azure Monitor integration requires ARM API access, not available in this browser-only tool.' });
    },

    '24540': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('windowsfirewall') || /firewall/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /firewall/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} Windows Firewall policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName, type: 'deviceConfiguration' })), ...inPolicies.map(p => ({ displayName: p.name, type: 'configurationPolicy' }))]);
      return fail(test, 'No Windows Firewall policy was found in Intune.');
    },

    '24552': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('macos') && /firewall/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /macos.*firewall|firewall.*macos/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} macOS Firewall policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName, type: 'deviceConfiguration' })), ...inPolicies.map(p => ({ displayName: p.name, type: 'configurationPolicy' }))]);
      return fail(test, 'No macOS Firewall policy was found in Intune.');
    },

    '24554': async (test) => {
      const configs = await getDeviceConfigurations();
      if (configs === null) return skip(test, 'Intune device configuration is not accessible.');
      const ios = toArray(configs).filter(c => lower(c['@odata.type']).includes('iosupdate') || (lower(c['@odata.type']).includes('ios') && /update/i.test(c.displayName)));
      if (ios.length) return pass(test, `Found ${ios.length} iOS update policy(s).`, ios.map(c => ({ displayName: c.displayName, type: c['@odata.type'] })));
      return fail(test, 'No iOS update policy was found in Intune.');
    },

    '24555': async (test) => {
      const tags = await getDeviceRoleScopeTags();
      if (tags === null) return skip(test, 'Intune role scope tags are not accessible (DeviceManagementRBAC.Read.All required).');
      // The Default scope tag always exists — if only 1 and it is "Default", it's not configured
      const custom = tags.filter(t => lower(t.displayName) !== 'default');
      if (custom.length) return pass(test, `Found ${custom.length} custom Intune scope tag(s) configured for delegated administration.`, custom.map(t => ({ id: t.id, displayName: t.displayName })));
      return fail(test, 'No custom Intune scope tags are configured. Scope tags are needed for delegated administration of different device groups or regions.');
    },

    '24564': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('windowslocalaccountmanager') || /local.?account|local.?admin/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /local.?account|local.?admin/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} Windows Local Account Protection policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName })), ...inPolicies.map(p => ({ displayName: p.name }))]);
      return fail(test, 'No Windows Local Account Protection policy was found in Intune.');
    },

    '24568': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('macos') && /platform.?sso|sso|single.?sign.?on/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /platform.?sso|macos.*sso/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} macOS Platform SSO policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName })), ...inPolicies.map(p => ({ displayName: p.name }))]);
      return fail(test, 'No macOS Platform SSO policy was found in Intune.');
    },

    '24569': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('macosfilevault') || /filevault/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /filevault/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} macOS FileVault policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName })), ...inPolicies.map(p => ({ displayName: p.name }))]);
      return fail(test, 'No macOS FileVault policy was found in Intune.');
    },

    '24572': async (test) => {
      const notifications = await getEnrollmentNotificationConfigurations();
      if (notifications === null) return skip(test, 'Intune device enrollment notification configurations are not accessible (additional Intune service config scopes may be required).');
      const assigned = toArray(notifications).filter(n => toArray(n.assignments).length > 0);
      if (assigned.length) return pass(test, `Found ${assigned.length} assigned device enrollment notification configuration(s).`, assigned.map(n => ({ displayName: n.displayName, assignments: toArray(n.assignments).length })));
      return fail(test, 'No assigned device enrollment notification configuration was found in Intune.');
    },

    '24573': async (test) => {
      const [configs, policies, intents] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies(), getDeviceIntents()]);
      if (configs === null && policies === null && intents === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => /security.?baseline|baseline/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /security.?baseline|baseline/i.test(p.name || ''));
      const inIntents = toArray(intents).filter(i => /security.?baseline|baseline/i.test(i.displayName || '') || /security.?baseline/i.test(i.description || ''));
      const found = inConfigs.length + inPolicies.length + inIntents.length;
      if (found) return pass(test, `Found ${found} Windows security baseline policy(s).`);
      return fail(test, 'No Windows security baseline policy was found in Intune.');
    },

    '24574': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('endpointprotection') || /attack.?surface|asr/i.test(c.displayName));
      const inPolicies = toArray(policies).filter(p => /attack.?surface|asr/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} Attack Surface Reduction (ASR) policy(s).`, [...inConfigs.map(c => ({ displayName: c.displayName })), ...inPolicies.map(p => ({ displayName: p.name }))]);
      return fail(test, 'No Attack Surface Reduction (ASR) policies were found in Intune.');
    },

    '24576': async (test) => {
      const configs = await graphAllWithFallback('deviceManagement/deviceConfigurations', [
        { beta: true, query: { '$select': 'id,displayName,@odata.type,assignments', '$expand': 'assignments', '$top': '200' } },
        { beta: true, query: { '$expand': 'assignments', '$top': '200' } },
      ]).catch(err => {
        if (err.status === 400 || err.status === 403 || err.status === 404) return null;
        throw err;
      });

      if (configs === null) return skip(test, 'Intune device configurations are not accessible (required Intune scopes may be missing).');

      const windowsHealthMonitoring = toArray(configs).filter(c => lower(c['@odata.type']) === '#microsoft.graph.windowshealthmonitoringconfiguration');
      const assigned = windowsHealthMonitoring.filter(c => toArray(c.assignments).length > 0);
      if (assigned.length) return pass(test, `Found ${assigned.length} assigned Endpoint Analytics (Windows Health Monitoring) policy(s).`, assigned.map(p => ({ displayName: p.displayName, assignments: toArray(p.assignments).length })));
      return fail(test, 'Endpoint Analytics policy is not created or not assigned.');
    },

    '24690': async (test) => {
      const configs = await getDeviceConfigurations();
      if (configs === null) return skip(test, 'Intune device configuration is not accessible.');
      const macos = toArray(configs).filter(c => (lower(c['@odata.type']).includes('macos') && /update/i.test(c.displayName)) || lower(c['@odata.type']).includes('macosupdateconfiguration'));
      if (macos.length) return pass(test, `Found ${macos.length} macOS update policy(s).`, macos.map(c => ({ displayName: c.displayName, type: c['@odata.type'] })));
      return fail(test, 'No macOS update policy was found in Intune.');
    },

    '24784': async (test) => {
      const [configs, policies] = await Promise.all([getDeviceConfigurations(), getConfigurationPolicies()]);
      if (configs === null && policies === null) return skip(test, 'Intune device configuration is not accessible.');
      const inConfigs = toArray(configs).filter(c => lower(c['@odata.type']).includes('macos') && (/defender|antivirus/i.test(c.displayName)));
      const inPolicies = toArray(policies).filter(p => /defender.*macos|macos.*defender|antivirus.*macos|macos.*antivirus/i.test(p.name || ''));
      const found = inConfigs.length + inPolicies.length;
      if (found) return pass(test, `Found ${found} Microsoft Defender Antivirus policy(s) for macOS.`);
      return fail(test, 'No Microsoft Defender Antivirus policy for macOS was found in Intune.');
    },

    '24794': async (test) => {
      const tac = await getTermsAndConditionsWithAssignments();
      if (tac === null) return skip(test, 'Intune Terms and Conditions are not accessible (additional Intune service config scopes may be required).');
      const assigned = toArray(tac).filter(t => toArray(t.assignments).length > 0);
      if (assigned.length) return pass(test, `Found ${assigned.length} assigned Terms and Conditions policy(s) in Intune.`, assigned.map(t => ({ displayName: t.displayName, assignments: toArray(t.assignments).length })));
      return fail(test, 'No assigned Terms and Conditions policy was found in Intune.');
    },

    '24802': async (test) => {
      const rules = await getDeviceCleanupRules();
      if (rules === null) return skip(test, 'Device cleanup rules are not accessible.');
      if (rules.length) return pass(test, `Found ${rules.length} device cleanup rule(s).`, rules.map(r => ({ displayName: r.displayName, daysInactive: r.deviceInactivityBeforeRetirementInDays })));
      return fail(test, 'No device cleanup rules are configured. Stale devices should be automatically removed from Intune.');
    },

    '21832': async (test) => {
      const policies = await getConditionalAccessPolicies();
      const groupIds = new Set();
      for (const policy of policies.filter(p => lower(p.state) === 'enabled')) {
        for (const gid of toArray(policy?.conditions?.users?.includeGroups)) groupIds.add(gid);
        for (const gid of toArray(policy?.conditions?.users?.excludeGroups)) groupIds.add(gid);
      }
      if (!groupIds.size) return pass(test, 'No groups are referenced in enabled Conditional Access policies.');
      const unprotected = [];
      let lookupFailures = 0;
      for (const gid of [...groupIds].slice(0, 20)) {
        try {
          const memberOf = await graphAllWithFallback(`groups/${gid}/memberOf/microsoft.graph.administrativeUnit`, [
            { beta: true, query: { '$select': 'id,displayName,isMemberManagementRestricted', '$top': '10' } },
            { beta: true, query: { '$top': '10' } },
          ]);
          const hasRestricted = memberOf.some(au => au.isMemberManagementRestricted === true);
          if (!hasRestricted) {
            const group = await Api.graph(`groups/${gid}`, { query: { '$select': 'id,displayName' } }).catch(() => null);
            unprotected.push({ id: gid, displayName: group?.displayName || gid });
          }
        } catch {
          lookupFailures += 1;
        }
      }
      if (lookupFailures > 0 && lookupFailures === Math.min(groupIds.size, 20)) {
        return skip(test, 'Could not validate administrative unit membership for Conditional Access groups due Graph API query limitations (HTTP 400/permission variance).');
      }
      if (!unprotected.length) return pass(test, `All ${groupIds.size} groups used in Conditional Access policies are members of restricted-management administrative units.`);
      return fail(test, `${unprotected.length} of ${Math.min(groupIds.size, 20)} sampled groups in Conditional Access policies are not in a restricted management administrative unit.`, unprotected.slice(0, 20));
    },

    '21833': async (test) => {
      const syncUsers = await Api.graphAll('users', {
        beta: true,
        query: {
          '$filter': "startsWith(userPrincipalName,'Sync_')",
          '$select': 'id,displayName,userPrincipalName,lastPasswordChangeDateTime,passwordPolicies,onPremisesSyncEnabled',
          '$top': '10',
        },
      }).catch(() => []);
      const syncSps = await graphAllWithFallback('servicePrincipals', [
        {
          beta: true,
          query: { '$filter': "contains(displayName,'Azure AD Connect') or contains(displayName,'Entra Connect Sync')", '$top': '5' },
        },
        {
          beta: true,
          query: { '$top': '5' },
        },
      ]).catch(() => []);
      if (!syncUsers.length && !syncSps.length) return skip(test, 'No directory synchronization service accounts found. Tenant may not use Entra Connect hybrid sync.');
      const staleCredentials = [];
      for (const user of syncUsers) {
        const lastChange = user.lastPasswordChangeDateTime;
        if (!lastChange) { staleCredentials.push({ type: 'user', displayName: user.displayName, userPrincipalName: user.userPrincipalName, lastPasswordChangeDateTime: null }); continue; }
        const daysSince = daysBetween(new Date().toISOString(), lastChange);
        if (daysSince > 90) staleCredentials.push({ type: 'user', displayName: user.displayName, userPrincipalName: user.userPrincipalName, lastPasswordChangeDateTime: lastChange, daysSince });
      }
      if (!staleCredentials.length) return pass(test, `Directory sync service account(s) found — credentials appear current.`, syncUsers.map(u => ({ userPrincipalName: u.userPrincipalName, lastPasswordChangeDateTime: u.lastPasswordChangeDateTime })));
      return fail(test, `Found ${staleCredentials.length} sync service account(s) with stale or unknown credential rotation.`, staleCredentials);
    },

    '21834': async (test) => {
      const syncUsers = await Api.graphAll('users', {
        beta: true,
        query: { '$filter': "startsWith(userPrincipalName,'Sync_')", '$select': 'id,displayName,userPrincipalName', '$top': '10' },
      }).catch(() => []);
      if (!syncUsers.length) return skip(test, 'No directory synchronization service accounts found (Sync_ pattern). Tenant may not use Entra Connect hybrid sync.');
      const syncUserIds = new Set(syncUsers.map(u => u.id));
      const policies = await getConditionalAccessPolicies();
      const coveringPolicies = policies.filter(policy => {
        if (lower(policy.state) !== 'enabled') return false;
        const includesUsers = toArray(policy?.conditions?.users?.includeUsers);
        const conditions = policy?.conditions?.locations;
        const hasLocationCondition = !!conditions?.includeLocations?.length || !!conditions?.excludeLocations?.length;
        return syncUserIds.size > 0 && includesUsers.some(id => syncUserIds.has(id)) && hasLocationCondition;
      });
      if (coveringPolicies.length) return pass(test, `Found ${coveringPolicies.length} Conditional Access policy(s) targeting sync accounts with location conditions.`, coveringPolicies.map(p => ({ displayName: p.displayName })));
      return fail(test, `Found ${syncUsers.length} sync account(s) but no Conditional Access policy restricts their access to specific named locations.`, syncUsers.map(u => ({ userPrincipalName: u.userPrincipalName })));
    },

    '21882': async (test) => {
      try {
        const instances = await Api.graphAll('identityGovernance/privilegedAccess/group/assignmentScheduleInstances', {
          beta: true, query: { '$select': 'id,principalId,groupId,accessId,memberType', '$top': '100' },
        });
        if (!instances.length) return pass(test, 'No PIM for Groups assignments found — no nested group risk.');
        const groupPrincipalIds = [...new Set(instances.map(i => i.principalId))];
        const nestedGroups = [];
        for (const principalId of groupPrincipalIds.slice(0, 20)) {
          try {
            const grp = await Api.graph(`groups/${principalId}`, { query: { '$select': 'id,displayName' } });
            if (grp?.id) nestedGroups.push({ id: principalId, displayName: grp.displayName });
          } catch { /* not a group */ }
        }
        if (!nestedGroups.length) return pass(test, `No nested groups found among ${groupPrincipalIds.length} PIM for Groups principals.`, { total: instances.length });
        return fail(test, `Found ${nestedGroups.length} group(s) assigned as principals in PIM for Groups — nested group memberships can bypass JIT access controls.`, nestedGroups.slice(0, 20));
      } catch (err) {
        if (err.status === 400 || err.status === 403 || err.status === 404) return skip(test, 'PIM for Groups assignment schedule is not accessible (P2 license required, or endpoint query shape unsupported for this tenant).');
        throw err;
      }
    },

    '21886': async (test) => {
      try {
        const provLogs = await Api.graphAll('auditLogs/provisioning', {
          beta: true, query: { '$top': '5' },
        });
        if (provLogs.length) return pass(test, `Automatic provisioning is active — ${provLogs.length} recent provisioning log entries found.`, provLogs.slice(0, 5).map(l => ({ jobId: l.jobId, servicePrincipal: l.servicePrincipal?.displayName, activityDateTime: l.activityDateTime })));
        return fail(test, 'No recent automatic provisioning activity was found. Applications that support Entra-based provisioning should have provisioning configured.');
      } catch (err) {
        if (err.status === 400 || err.status === 403 || err.status === 404) return skip(test, 'Provisioning audit logs are not accessible (AuditLog.Read.All required, or query shape unsupported for this tenant).');
        throw err;
      }
    },

    '21887': async (test) => {
      let apps = [];
      try {
        apps = await graphAllWithFallback('applications', [
          { query: { '$select': 'id,displayName,appId,web,spa,publicClient', '$top': '200' } },
          { beta: true, query: { '$select': 'id,displayName,appId,web,spa,publicClient', '$top': '200' } },
          { query: { '$top': '200' } },
        ]);
      } catch (err) {
        if (err.status === 400 || err.status === 403 || err.status === 404) {
          return skip(test, 'Could not enumerate application redirect URIs due Graph API limitations or insufficient permissions.');
        }
        throw err;
      }
      const suspicious = [];
      for (const app of apps) {
        const allUris = [...toArray(app.web?.redirectUris), ...toArray(app.spa?.redirectUris), ...toArray(app.publicClient?.redirectUris)];
        for (const uri of allUris) {
          try {
            const parsed = new URL(uri);
            const host = parsed.hostname;
            // Flag wildcards, IP addresses, generic catch-all patterns
            if (host.includes('*') || /^(\d+\.){3}\d+$/.test(host) && !['127.0.0.1'].includes(host)) {
              suspicious.push({ appDisplayName: app.displayName, appId: app.appId, uri, reason: 'wildcard or bare IP address' });
            }
          } catch {
            suspicious.push({ appDisplayName: app.displayName, appId: app.appId, uri, reason: 'invalid URL format' });
          }
        }
      }
      if (!suspicious.length) return pass(test, `No obviously invalid redirect URIs found in ${apps.length} app registrations. Note: full DNS ownership verification requires DNS lookup capability not available in this browser tool.`);
      return fail(test, `Found ${suspicious.length} suspicious redirect URI(s) that should be reviewed for DNS ownership.`, suspicious.slice(0, 20));
    },

    '21897': async (test) => {
      const packages = await getEntitlementAccessPackages();
      if (packages === null) return skip(test, 'Entitlement management access packages are not accessible (EntitlementManagement.Read.All required).');
      if (!packages.length) return fail(test, 'No entitlement management access packages found. App role assignments and group memberships may not be governed through entitlement management.');
      return pass(test, `Found ${packages.length} entitlement management access package(s). App assignments and group memberships can be governed through these packages.`, { count: packages.length });
    },

    '21898': async (test) => {
      const packages = await getEntitlementAccessPackages();
      if (packages === null) return skip(test, 'Entitlement management access packages are not accessible (EntitlementManagement.Read.All required).');
      if (!packages.length) return fail(test, 'No entitlement management access packages found. Supported access lifecycle resources should be managed through entitlement management packages.');
      const policies = await getEntitlementAssignmentPolicies();
      if (!policies?.length) return fail(test, `Found ${packages.length} access package(s) but no assignment policies are configured to manage access lifecycle.`, { packages: packages.length });
      const withExpiry = policies.filter(p => p.expiration && lower(p.expiration?.type) !== 'noexpiration' && (p.expiration?.duration || p.expiration?.endDateTime)).length;
      if (withExpiry) return pass(test, `Found ${packages.length} access package(s) and ${withExpiry} policy(s) with expiration configured for lifecycle management.`, { packages: packages.length, policiesWithExpiry: withExpiry });
      return fail(test, `Found ${packages.length} access package(s) but none of the ${policies.length} assignment policies have expiration configured for lifecycle management.`);
    },

    '21985': async (test) => {
      const org = await getOrganization();
      if (!org?.onPremisesSyncEnabled) return pass(test, 'On-premises directory sync is not enabled — Seamless SSO is not applicable.');
      // Look for the Seamless SSO service principal — it uses the Kerberos delegation token
      const ssoSps = await Api.graphAll('servicePrincipals', {
        query: { '$filter': "contains(displayName,'Seamless SSO') or contains(displayName,'AZUREADSSOACC')", '$select': 'id,displayName,appId,accountEnabled', '$top': '5' },
        beta: true,
      }).catch(() => []);
      if (!ssoSps.length) return skip(test, 'Could not detect Seamless SSO service principal. Verify Seamless SSO usage in Entra ID → Hybrid Identity → Entra Connect → Single sign-on.');
      const enabled = ssoSps.filter(sp => sp.accountEnabled !== false);
      if (!enabled.length) return pass(test, 'Seamless SSO service principal exists but is disabled — this control recommends disabling if not in use.', ssoSps.map(s => ({ displayName: s.displayName, accountEnabled: s.accountEnabled })));
      // Check if it's being used via sign-in logs
      const signIns = await getSignInLogs(`appId eq '${enabled[0].appId}'`, 5);
      if (signIns?.length) return pass(test, `Seamless SSO is enabled and sign-in activity was detected — usage appears active.`, { recentSignIns: signIns.length, sp: enabled[0].displayName });
      return fail(test, 'Seamless SSO is enabled but no recent sign-in activity was detected — consider disabling Seamless SSO if it is no longer in use.', enabled.map(s => ({ displayName: s.displayName, appId: s.appId })));
    },

    '22102': async (test) => {
      const payload = await graphWithFallback('domains', [
        { query: { '$select': 'id,isDefault,isInitial,isVerified', '$top': '50' } },
        { query: { '$top': '50' } },
      ]).catch(err => {
        if (err.status === 400 || err.status === 403 || err.status === 404) return null;
        throw err;
      });
      if (!payload) return skip(test, 'Could not enumerate verified domains due Graph API limitations or insufficient permissions.');
      const domains = toArray(payload?.value);
      const customDomains = domains.filter(d => !lower(d.id).endsWith('.onmicrosoft.com') && !d.isInitial);
      if (customDomains.length) return pass(test, `Found ${customDomains.length} verified custom domain(s).`, customDomains.map(d => ({ id: d.id, isDefault: d.isDefault })));
      return fail(test, 'No verified custom domains found. The tenant only uses the default .onmicrosoft.com domain — configure a custom domain for a professional identity.');
    },

    '21778': async (test) => {
      const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
      const signIns = await getSignInLogs(`createdDateTime ge ${thirtyDaysAgo}T00:00:00Z and status/errorCode eq 0`, 50);
      if (signIns === null) return skip(test, 'Sign-in logs are not accessible or timed out (AuditLog.Read.All required).');
      const nonMsal = signIns.filter(s => {
        const lib = lower(s.authenticationLibrary?.name || s.authenticationLibrary || '');
        return lib.length > 0 && !lib.includes('msal') && !lib.includes('microsoft authentication library');
      });
      if (!nonMsal.length) return pass(test, `All ${signIns.length} sampled sign-ins with identifiable authentication library use MSAL or no library info is available.`);
      const legacy = nonMsal.filter(s => {
        const lib = lower(s.authenticationLibrary?.name || s.authenticationLibrary || '');
        return lib.includes('adal') || lib.includes('wia') || lib.includes('passport') || lib.includes('azure active directory authentication library');
      });
      if (!legacy.length) return pass(test, `${nonMsal.length} sign-in(s) use non-MSAL libraries but none match known legacy libraries. Review if custom apps are using current SDKs.`, nonMsal.slice(0, 10).map(s => ({ appDisplayName: s.appDisplayName, authenticationLibrary: s.authenticationLibrary })));
      return fail(test, `Found ${legacy.length} sign-in(s) using non-MSAL authentication libraries. Line-of-business and partner apps should use MSAL.`, legacy.slice(0, 20).map(s => ({ appDisplayName: s.appDisplayName, userPrincipalName: s.userPrincipalName, authenticationLibrary: s.authenticationLibrary })));
    },

    '21779': async (test) => {
      const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
      const signIns = await getSignInLogs(`createdDateTime ge ${thirtyDaysAgo}T00:00:00Z and status/errorCode eq 0 and appDisplayName eq 'Microsoft Office'`, 25);
      if (signIns === null) return skip(test, 'Sign-in logs are not accessible or timed out. Verifying Microsoft application versions requires sign-in log access (AuditLog.Read.All required).');
      // We can check client app user agent from signIn, but Graph doesn't expose full user agent version in basic sign-in fields
      return skip(test, 'Application version data is not available through the Microsoft Graph sign-in logs API at the required granularity. Use Microsoft Intune or Endpoint Analytics to verify application versions.');
    },

    '21788': async (test) => {
      const privilegedAccessAdminRoleDefinitionId = '18d7d88d-d35e-4fb5-a5c3-7773c20a72d9';
      const arm = await getArmRootRoleAssignments();
      if (arm.scopePrompted) return skip(test, 'Azure management scope consent has been requested. Re-run the assessment after completing consent for https://management.azure.com/user_impersonation.');
      if (arm.assignments === null) return skip(test, 'Azure role assignments at tenant root are not accessible with current Azure permissions.');

      const standingRootAssignments = toArray(arm.assignments).filter(a => lower(a?.properties?.roleDefinitionId || '').endsWith(`/${privilegedAccessAdminRoleDefinitionId}`));
      if (!standingRootAssignments.length) {
        return pass(test, 'No standing access to Azure root management group was found.');
      }

      const principalIds = uniqueBy(standingRootAssignments.map(a => a?.properties?.principalId).filter(Boolean), x => x);
      const objects = await Promise.all(principalIds.map(async (id) => {
        try {
          return await Api.graph(`directoryObjects/${id}`, { query: { '$select': 'id,displayName,userPrincipalName' } });
        } catch {
          return { id, displayName: id, userPrincipalName: null };
        }
      }));
      const objById = new Map(objects.map(o => [o.id, o]));

      return fail(test, `Found ${standingRootAssignments.length} standing access assignment(s) to Azure root management group.`, standingRootAssignments.slice(0, 30).map(a => {
        const pid = a?.properties?.principalId;
        const obj = objById.get(pid) || null;
        return {
          principalId: pid,
          principalType: a?.properties?.principalType,
          principal: obj?.userPrincipalName || obj?.displayName || pid,
          roleDefinitionId: a?.properties?.roleDefinitionId,
        };
      }));
    },

    '21859': async (test) => {
      return skip(test, 'GDAP (Granular Delegated Admin Privileges) admin least privilege verification requires Microsoft Partner Center API access, which is not available in this browser-only tool. Review GDAP roles in the Microsoft 365 Admin Center → Settings → Partner relationships.');
    },

    '21881': async (test) => {
      return skip(test, 'Verifying that Azure subscriptions used by Identity Governance only allow access from privileged roles requires Azure Resource Manager (ARM) RBAC API access, which is not available in this browser-only tool.');
    },

    '21894': async (test) => {
      const apps = await graphAllWithFallback('applications', [
        { query: { '$select': 'id,displayName,appId,keyCredentials', '$top': '200' } },
        { beta: true, query: { '$select': 'id,displayName,appId,keyCredentials', '$top': '200' } },
        { query: { '$top': '200' } },
      ]).catch(err => {
        if (err.status === 400 || err.status === 403 || err.status === 404) return null;
        throw err;
      });
      if (apps === null) return skip(test, 'Could not enumerate application certificates due Graph API limitations or insufficient permissions.');
      const certsFound = [];
      for (const app of apps) {
        for (const key of toArray(app.keyCredentials)) {
          if (key.type === 'AsymmetricX509Cert' || key.usage === 'Verify') {
            certsFound.push({ appDisplayName: app.displayName, appId: app.appId, keyDisplayName: key.displayName || '', endDateTime: key.endDateTime, startDateTime: key.startDateTime });
          }
        }
      }
      if (!certsFound.length) return pass(test, 'No application certificate credentials found in app registrations.');
      return fail(test, `Found ${certsFound.length} certificate credential(s) in app registrations. Verify each certificate is issued by an approved CA. Note: full CA chain validation requires certificate parsing not available in this browser tool.`, { note: 'Check certificate issuers in Azure Portal → App registrations → Certificates & secrets.', count: certsFound.length, sample: certsFound.slice(0, 10) });
    },

    '21895': async (test) => {
      return skip(test, 'Verifying that application certificate credentials are managed using HSM (Hardware Security Module) requires inspection of key vault and certificate provisioning infrastructure, which is not accessible via the Microsoft Graph API in this browser tool.');
    },

    '21912': async (test) => {
      return skip(test, 'Verifying that Azure resources used by Microsoft Entra only allow access from privileged roles requires Azure Resource Manager (ARM) RBAC API access, which is not available in this browser-only tool.');
    },

    '22100': async (test) => {
      return skip(test, 'WAF (Web Application Firewall) configuration for ciamlogin.com endpoints requires Azure Front Door or Application Gateway API access, which is not available via Microsoft Graph in this browser tool. Verify in Azure Portal → Azure Front Door → WAF policies.');
    },

    '22101': async (test) => {
      return skip(test, 'ciamlogin endpoint management requires CIAM (External ID) tenant configuration access, which is not available via the standard Microsoft Graph API in this browser tool.');
    },

  };

  window.ZtaImpl = {
    get: (id) => impl[id] || null,
    count: () => Object.keys(impl).length,
  };
})();
