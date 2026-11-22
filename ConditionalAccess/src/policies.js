// Pure policy logic: import validation, principal references, lockout check, diffs and
// the starter set. No DOM, no network, so tests/test.html covers all of it.
(() => {
  // Loose on purpose: Microsoft first-party app ids such as 00000002-0000-0ff1-ce00-...
  // are not RFC 4122 v1-5.
  const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const MAX_FILE_BYTES = 5 * 1024 * 1024;
  const MAX_POLICIES = 500;
  const STATES = ['disabled', 'enabledForReportingButNotEnforced', 'enabled'];
  const DEFAULT_STATE = 'enabledForReportingButNotEnforced';

  const TOP_KEYS = ['displayName', 'state', 'conditions', 'grantControls', 'sessionControls'];
  const READ_ONLY_KEYS = ['id', 'createdDateTime', 'modifiedDateTime', 'deletedDateTime', 'templateId', 'partialEnablementStrategy', 'conditionsNotSatisfied', 'description'];
  const CONDITION_KEYS = ['applications', 'users', 'clientApplications', 'clientAppTypes', 'platforms', 'locations', 'devices',
    'deviceStates', 'signInRiskLevels', 'userRiskLevels', 'servicePrincipalRiskLevels', 'insiderRiskLevels', 'authenticationFlows', 'times'];
  const GRANT_KEYS = ['operator', 'builtInControls', 'customAuthenticationFactors', 'termsOfUse', 'authenticationStrength'];
  const BAD_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

  // Every slot that can hold an object id, and what kind of object it is. Tenant-specific
  // kinds must exist in the target tenant; roles and apps are global ids.
  const REF_SLOTS = [
    { path: ['users', 'includeUsers'], kind: 'user', label: 'Include users' },
    { path: ['users', 'excludeUsers'], kind: 'user', label: 'Exclude users' },
    { path: ['users', 'includeGroups'], kind: 'group', label: 'Include groups' },
    { path: ['users', 'excludeGroups'], kind: 'group', label: 'Exclude groups' },
    { path: ['users', 'includeRoles'], kind: 'role', label: 'Include roles' },
    { path: ['users', 'excludeRoles'], kind: 'role', label: 'Exclude roles' },
    { path: ['applications', 'includeApplications'], kind: 'app', label: 'Include apps' },
    { path: ['applications', 'excludeApplications'], kind: 'app', label: 'Exclude apps' },
    { path: ['clientApplications', 'includeServicePrincipals'], kind: 'servicePrincipal', label: 'Include workload identities' },
    { path: ['clientApplications', 'excludeServicePrincipals'], kind: 'servicePrincipal', label: 'Exclude workload identities' },
    { path: ['locations', 'includeLocations'], kind: 'location', label: 'Include locations' },
    { path: ['locations', 'excludeLocations'], kind: 'location', label: 'Exclude locations' },
  ];
  const MAPPABLE_KINDS = new Set(['user', 'group', 'servicePrincipal', 'location']);

  function isGuid(v) {
    return typeof v === 'string' && GUID_RE.test(v.trim());
  }

  function isPlainObject(v) {
    return v !== null && typeof v === 'object' && !Array.isArray(v);
  }

  function clone(v) {
    return v === undefined ? undefined : JSON.parse(JSON.stringify(v));
  }

  // Deep JSON copy without prototype keys and OData annotations. Reports what it dropped.
  function cleanValue(v, path, dropped) {
    if (Array.isArray(v)) return v.map((x, i) => cleanValue(x, `${path}[${i}]`, dropped));
    if (!isPlainObject(v)) return v;
    const out = {};
    for (const k of Object.keys(v)) {
      if (BAD_KEYS.has(k) || k.startsWith('@odata') || k.includes('@odata.')) {
        dropped.push({ path: path ? `${path}.${k}` : k, reason: BAD_KEYS.has(k) ? 'unsafe key' : 'annotation' });
        continue;
      }
      out[k] = cleanValue(v[k], path ? `${path}.${k}` : k, dropped);
    }
    return out;
  }

  // parseImportPayload(text or object) -> { policies: raw[], metadata }
  function parseImportPayload(raw) {
    if (typeof raw === 'string' && raw.length > MAX_FILE_BYTES) throw new Error('File is larger than 5 MB. That is not a policy export.');
    let parsed;
    try { parsed = typeof raw === 'string' ? JSON.parse(raw) : raw; }
    catch (e) { throw new Error(`Not valid JSON: ${e.message}`); }
    let policies;
    let metadata = {};
    if (Array.isArray(parsed)) policies = parsed;
    else if (isPlainObject(parsed) && Array.isArray(parsed.policies)) {
      policies = parsed.policies;
      if (isPlainObject(parsed.metadata)) metadata = cleanValue(parsed.metadata, 'metadata', []);
    } else if (isPlainObject(parsed) && Array.isArray(parsed.value)) {
      policies = parsed.value;
    } else {
      throw new Error('Expected an array of policies, { "policies": [...] } or a Graph list response.');
    }
    if (policies.length > MAX_POLICIES) throw new Error(`File holds ${policies.length} policies, the limit is ${MAX_POLICIES}.`);
    return { policies, metadata };
  }

  // Strict check of one policy from an untrusted file. Returns { policy, fileState, dropped, errors }.
  function sanitizePolicy(input, index = 0) {
    const errors = [];
    const dropped = [];
    const where = `Policy ${index + 1}`;
    if (!isPlainObject(input)) return { policy: null, fileState: null, dropped, errors: [`${where} is not an object.`] };

    const out = {};
    for (const k of Object.keys(input)) {
      if (TOP_KEYS.includes(k)) continue;
      dropped.push({ path: k, reason: READ_ONLY_KEYS.includes(k) ? 'read-only' : BAD_KEYS.has(k) ? 'unsafe key' : k.includes('@odata') ? 'annotation' : 'unknown field' });
    }

    const name = input.displayName;
    if (typeof name !== 'string' || !name.trim()) errors.push(`${where} has no displayName.`);
    else if (name.length > 256) errors.push(`${where}: displayName is longer than 256 characters.`);
    else out.displayName = name.trim();
    const label = out.displayName ? `"${out.displayName}"` : where;

    const fileState = STATES.includes(input.state) ? input.state : null;

    const cond = input.conditions;
    if (!isPlainObject(cond)) errors.push(`${label} has no conditions object.`);
    else {
      out.conditions = {};
      for (const k of Object.keys(cond)) {
        if (!CONDITION_KEYS.includes(k)) {
          dropped.push({ path: `conditions.${k}`, reason: BAD_KEYS.has(k) ? 'unsafe key' : k.includes('@odata') ? 'annotation' : 'unknown field' });
          continue;
        }
        if (cond[k] !== null && typeof cond[k] !== 'object') { errors.push(`${label}: conditions.${k} has the wrong type.`); continue; }
        out.conditions[k] = cleanValue(cond[k], `conditions.${k}`, dropped);
      }
      if (!isPlainObject(out.conditions.users)) errors.push(`${label} has no conditions.users.`);
      if (!isPlainObject(out.conditions.applications)) errors.push(`${label} has no conditions.applications.`);
      for (const slot of REF_SLOTS) {
        const list = getPath(out.conditions, slot.path);
        if (list == null) continue;
        if (!Array.isArray(list) || list.some((v) => typeof v !== 'string')) errors.push(`${label}: conditions.${slot.path.join('.')} must be a list of strings.`);
      }
    }

    const grant = input.grantControls;
    if (grant != null) {
      if (!isPlainObject(grant)) errors.push(`${label}: grantControls has the wrong type.`);
      else {
        out.grantControls = {};
        for (const k of Object.keys(grant)) {
          if (!GRANT_KEYS.includes(k)) {
            dropped.push({ path: `grantControls.${k}`, reason: BAD_KEYS.has(k) ? 'unsafe key' : k.includes('@odata') ? 'annotation' : 'unknown field' });
            continue;
          }
          out.grantControls[k] = cleanValue(grant[k], `grantControls.${k}`, dropped);
        }
        // Graph only takes the id; the rest of an exported strength object is read-only.
        const strength = out.grantControls.authenticationStrength;
        if (strength != null) {
          if (isPlainObject(strength) && isGuid(strength.id)) out.grantControls.authenticationStrength = { id: strength.id };
          else errors.push(`${label}: grantControls.authenticationStrength has no valid id.`);
        }
        if (out.grantControls.builtInControls != null && (!Array.isArray(out.grantControls.builtInControls) || out.grantControls.builtInControls.some((v) => typeof v !== 'string'))) {
          errors.push(`${label}: grantControls.builtInControls must be a list of strings.`);
        }
      }
    } else if (grant === null) out.grantControls = null;

    const session = input.sessionControls;
    if (session != null) {
      if (!isPlainObject(session)) errors.push(`${label}: sessionControls has the wrong type.`);
      else out.sessionControls = cleanValue(session, 'sessionControls', dropped);
    } else if (session === null) out.sessionControls = null;

    if (out.grantControls == null && out.sessionControls == null && !errors.length) errors.push(`${label} has no grant or session controls.`);
    return { policy: errors.length ? null : out, fileState, dropped, errors };
  }

  // Whole file: throws on any structural problem so nothing half-valid gets imported.
  function validateImport(raw) {
    const { policies, metadata } = parseImportPayload(raw);
    if (!policies.length) throw new Error('The file holds no policies.');
    const results = policies.map((p, i) => sanitizePolicy(p, i));
    const errors = results.flatMap((r) => r.errors);
    if (errors.length) {
      const shown = errors.slice(0, 6).join(' ');
      throw new Error(`File rejected. ${shown}${errors.length > 6 ? ` (+${errors.length - 6} more)` : ''}`);
    }
    return { metadata, policies: results.map(({ policy, fileState, dropped }) => ({ policy, fileState, dropped })) };
  }

  function getPath(obj, path) {
    let cur = obj;
    for (const p of path) {
      if (!isPlainObject(cur)) return undefined;
      cur = cur[p];
    }
    return cur;
  }

  // The body Graph gets: no read-only fields, chosen name and state.
  function normalizePolicyForCreateOrUpdate(policy, displayName, state) {
    const src = isPlainObject(policy) ? policy : {};
    const out = {};
    for (const k of TOP_KEYS) if (src[k] !== undefined) out[k] = clone(src[k]);
    out.displayName = displayName;
    out.state = STATES.includes(state) ? state : DEFAULT_STATE;
    return out;
  }

  // [{ slot, kind, label, value, isId }] for every value in every reference slot.
  function extractRefs(policy) {
    const cond = (policy || {}).conditions || {};
    const out = [];
    for (const slot of REF_SLOTS) {
      const list = getPath(cond, slot.path);
      if (!Array.isArray(list)) continue;
      for (const v of list) {
        if (typeof v !== 'string') continue;
        out.push({ slot: slot.path.join('.'), kind: slot.kind, label: slot.label, value: v, isId: isGuid(v) });
      }
    }
    return out;
  }

  // User and group GUIDs only (kept for the resolver tests and callers).
  function extractPrincipalIds(policy) {
    return Array.from(new Set(extractRefs(policy)
      .filter((r) => r.isId && (r.kind === 'user' || r.kind === 'group'))
      .map((r) => r.value.toLowerCase())));
  }

  // Replace tenant-specific ids by their mapped values. Roles and apps are left alone.
  function applyPrincipalMapping(policy, mapByLowerId) {
    const out = clone(policy);
    const cond = (out || {}).conditions || {};
    for (const slot of REF_SLOTS) {
      if (!MAPPABLE_KINDS.has(slot.kind)) continue;
      const parent = getPath(cond, slot.path.slice(0, -1));
      const key = slot.path[slot.path.length - 1];
      if (!parent || !Array.isArray(parent[key])) continue;
      parent[key] = Array.from(new Set(parent[key].map((v) => (isGuid(v) && mapByLowerId[v.toLowerCase()]) || v)));
    }
    return out;
  }

  function addUserExclusion(policy, userId) {
    const out = clone(policy);
    if (!isGuid(userId)) return out;
    out.conditions = out.conditions || {};
    out.conditions.users = out.conditions.users || {};
    const list = Array.isArray(out.conditions.users.excludeUsers) ? out.conditions.users.excludeUsers : [];
    if (!list.some((v) => String(v).toLowerCase() === userId.toLowerCase())) list.push(userId);
    out.conditions.users.excludeUsers = list;
    return out;
  }

  // Can this policy lock out the admin who writes it?
  // me: { userId, groupIds: [], roleIds: [] }. Returns null or { kind: 'block'|'noExclusions', reason }.
  function lockoutRisk(policy, me = {}) {
    const users = ((policy || {}).conditions || {}).users || {};
    const arr = (k) => (Array.isArray(users[k]) ? users[k].map((v) => String(v).toLowerCase()) : []);
    if (!arr('includeUsers').includes('all')) return null;
    const exU = arr('excludeUsers'), exG = arr('excludeGroups'), exR = arr('excludeRoles');
    const lower = (list) => (list || []).map((v) => String(v).toLowerCase());
    const excludesMe = (me.userId && exU.includes(String(me.userId).toLowerCase()))
      || lower(me.groupIds).some((g) => exG.includes(g))
      || lower(me.roleIds).some((r) => exR.includes(r));
    if (excludesMe) return null;
    const controls = (((policy || {}).grantControls || {}).builtInControls || []).map((v) => String(v).toLowerCase());
    if (controls.includes('block')) return { kind: 'block', reason: 'Blocks all users and does not exclude you.' };
    const excludesNobody = !exU.length && !exG.length && !exR.length && !users.excludeGuestsOrExternalUsers;
    if (excludesNobody) return { kind: 'noExclusions', reason: 'Applies to all users and excludes nobody, not even an emergency account.' };
    return null;
  }

  function stripNoiseForDiff(policy) {
    const p = clone(policy) || {};
    for (const k of READ_ONLY_KEYS) delete p[k];
    for (const k of Object.keys(p)) if (k.includes('@odata')) delete p[k];
    return p;
  }

  function stableSortObjectKeys(obj) {
    if (Array.isArray(obj)) return obj.map(stableSortObjectKeys);
    if (!obj || typeof obj !== 'object') return obj;
    const out = {};
    Object.keys(obj).sort().forEach((k) => {
      if (BAD_KEYS.has(k)) return;
      out[k] = stableSortObjectKeys(obj[k]);
    });
    return out;
  }

  function diffObjects(a, b, path = '') {
    const changes = [];
    const left = a === undefined ? null : a;
    const right = b === undefined ? null : b;
    if (JSON.stringify(left) === JSON.stringify(right)) return changes;

    const leftObj = left && typeof left === 'object';
    const rightObj = right && typeof right === 'object';

    if (!leftObj || !rightObj || Array.isArray(left) !== Array.isArray(right)) {
      changes.push(path || '(root)');
      return changes;
    }

    if (Array.isArray(left) && Array.isArray(right)) {
      if (left.length !== right.length) {
        changes.push(path + '.length');
        return changes;
      }
      for (let i = 0; i < left.length; i++) {
        changes.push(...diffObjects(left[i], right[i], `${path}[${i}]`));
      }
      return changes;
    }

    const keys = new Set([...Object.keys(left), ...Object.keys(right)]);
    for (const k of keys) {
      const childPath = path ? `${path}.${k}` : k;
      changes.push(...diffObjects(left[k], right[k], childPath));
    }
    return Array.from(new Set(changes));
  }

  // Empty arrays, empty objects and null all mean "not set" in a CA policy.
  function isEmptyish(v) {
    return v == null || (Array.isArray(v) && !v.length) || (isPlainObject(v) && Object.keys(v).every((k) => isEmptyish(v[k])));
  }

  // Field-level diff for review: [{ path, before, after }]. Lists of plain values are
  // compared as sets, since CA does not care about their order.
  function fieldDiff(a, b, path = '') {
    if (isEmptyish(a) && isEmptyish(b)) return [];
    if (isPlainObject(a) && isPlainObject(b)) {
      const keys = Array.from(new Set([...Object.keys(a), ...Object.keys(b)])).filter((k) => !BAD_KEYS.has(k)).sort();
      return keys.flatMap((k) => fieldDiff(a[k], b[k], path ? `${path}.${k}` : k));
    }
    if (Array.isArray(a) && Array.isArray(b) && [...a, ...b].every((v) => v === null || typeof v !== 'object')) {
      const sa = JSON.stringify([...a].map(String).sort());
      const sb = JSON.stringify([...b].map(String).sort());
      return sa === sb ? [] : [{ path: path || '(root)', before: a, after: b }];
    }
    if (JSON.stringify(stableSortObjectKeys(a)) === JSON.stringify(stableSortObjectKeys(b))) return [];
    return [{ path: path || '(root)', before: a === undefined ? null : a, after: b === undefined ? null : b }];
  }

  function comparePolicySets(oldPolicies, newPolicies) {
    const oldMap = new Map();
    const newMap = new Map();
    for (const p of oldPolicies || []) if (isPlainObject(p)) oldMap.set(String(p.displayName || '').trim(), p);
    for (const p of newPolicies || []) if (isPlainObject(p)) newMap.set(String(p.displayName || '').trim(), p);

    const added = [];
    const removed = [];
    const changed = [];
    const unchanged = [];

    for (const [name, oldP] of oldMap.entries()) {
      if (!newMap.has(name)) {
        removed.push(name);
        continue;
      }
      const left = stableSortObjectKeys(stripNoiseForDiff(oldP));
      const right = stableSortObjectKeys(stripNoiseForDiff(newMap.get(name)));
      const paths = diffObjects(left, right);
      if (paths.length) changed.push({ name, paths, fields: fieldDiff(left, right) });
      else unchanged.push(name);
    }
    for (const [name] of newMap.entries()) {
      if (!oldMap.has(name)) added.push(name);
    }

    return { added, removed, changed, unchanged };
  }

  function buildStarterPolicies() {
    return [
      {
        displayName: 'Baseline - Block legacy authentication',
        state: DEFAULT_STATE,
        conditions: {
          users: { includeUsers: ['All'], excludeUsers: [] },
          clientAppTypes: ['exchangeActiveSync', 'other'],
          applications: { includeApplications: ['All'], excludeApplications: [] },
        },
        grantControls: { operator: 'OR', builtInControls: ['block'] },
      },
      {
        displayName: 'Baseline - Require MFA for admin roles',
        state: DEFAULT_STATE,
        conditions: {
          users: {
            includeRoles: [
              '62e90394-69f5-4237-9190-012177145e10', // Global Administrator
              '194ae4cb-b126-40b2-bd5b-6091b380977d', // Security Administrator
              '729827e3-9c14-49f7-bb1b-9608f156bbb8', // Helpdesk Administrator
              'f28a1f50-f6e7-4571-818b-6a12f2af6b6c', // SharePoint Administrator
              '29232cdf-9323-42fd-ade2-1d097af3e4de', // Exchange Administrator
              'fe930be7-5e62-47db-91af-98c3a49a38b1', // User Administrator
              'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9', // Conditional Access Administrator
              'e8611ab8-c189-46e8-94e1-60213ab1f814', // Privileged Role Administrator
            ],
            excludeUsers: [],
            includeUsers: [],
            includeGroups: [],
            excludeGroups: [],
          },
          applications: { includeApplications: ['All'], excludeApplications: [] },
          clientAppTypes: ['all'],
        },
        grantControls: { operator: 'OR', builtInControls: ['mfa'] },
      },
      {
        displayName: 'Baseline - Require MFA for Microsoft admin portals',
        state: DEFAULT_STATE,
        conditions: {
          users: { includeUsers: ['All'], excludeUsers: [] },
          applications: { includeApplications: ['MicrosoftAdminPortals'], excludeApplications: [] },
          clientAppTypes: ['all'],
        },
        grantControls: { operator: 'OR', builtInControls: ['mfa'] },
      },
      {
        displayName: 'Baseline - Require MFA for all users',
        state: DEFAULT_STATE,
        conditions: {
          users: { includeUsers: ['All'], excludeUsers: [] },
          applications: { includeApplications: ['All'], excludeApplications: [] },
          clientAppTypes: ['browser', 'mobileAppsAndDesktopClients'],
        },
        grantControls: { operator: 'OR', builtInControls: ['mfa'] },
      },
      {
        displayName: 'Baseline - Protect security info registration with MFA',
        state: DEFAULT_STATE,
        conditions: {
          users: { includeUsers: ['All'], excludeUsers: ['GuestsOrExternalUsers'] },
          applications: { includeUserActions: ['urn:user:registersecurityinfo'] },
          locations: { includeLocations: ['All'], excludeLocations: ['AllTrusted'] },
          clientAppTypes: ['all'],
        },
        grantControls: { operator: 'OR', builtInControls: ['mfa'] },
      },
      {
        displayName: 'Baseline - Block device code authentication',
        state: DEFAULT_STATE,
        conditions: {
          users: { includeUsers: ['All'], excludeUsers: [] },
          applications: { includeApplications: ['All'], excludeApplications: [] },
          authenticationFlows: { transferMethods: 'deviceCodeFlow' },
          clientAppTypes: ['all'],
        },
        grantControls: { operator: 'OR', builtInControls: ['block'] },
      },
    ];
  }

  window.PolicyUtils = {
    isGuid,
    STATES,
    DEFAULT_STATE,
    MAPPABLE_KINDS,
    parseImportPayload,
    sanitizePolicy,
    validateImport,
    normalizePolicyForCreateOrUpdate,
    extractRefs,
    extractPrincipalIds,
    applyPrincipalMapping,
    addUserExclusion,
    lockoutRisk,
    stripNoiseForDiff,
    fieldDiff,
    comparePolicySets,
    buildStarterPolicies,
  };
})();
