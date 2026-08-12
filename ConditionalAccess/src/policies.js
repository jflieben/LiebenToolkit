(() => {
  const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

  function isGuid(v) {
    return typeof v === 'string' && GUID_RE.test(v.trim());
  }

  function clone(v) {
    return JSON.parse(JSON.stringify(v));
  }

  function parseImportPayload(raw) {
    const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (Array.isArray(parsed)) {
      return { policies: parsed, metadata: { source: 'array' } };
    }
    if (parsed && Array.isArray(parsed.policies)) {
      return { policies: parsed.policies, metadata: parsed.metadata || {} };
    }
    throw new Error('Invalid import JSON. Expected array of policies or { policies: [...] }.');
  }

  function normalizePolicyForCreateOrUpdate(policy, displayName, state) {
    const out = clone(policy);
    delete out.id;
    delete out.createdDateTime;
    delete out.modifiedDateTime;
    delete out.templateId;
    delete out.conditionsNotSatisfied;
    out.displayName = displayName;
    out.state = state;
    return out;
  }

  function extractPrincipalIds(policy) {
    const users = (((policy || {}).conditions || {}).users || {});
    const ids = [];
    const slots = [
      'includeUsers', 'excludeUsers',
      'includeGroups', 'excludeGroups',
    ];
    for (const slot of slots) {
      const list = users[slot];
      if (!Array.isArray(list)) continue;
      for (const v of list) {
        if (isGuid(v)) ids.push(v.toLowerCase());
      }
    }
    return Array.from(new Set(ids));
  }

  function applyPrincipalMapping(policy, mapByLowerId) {
    const out = clone(policy);
    const users = (((out || {}).conditions || {}).users || {});
    const slots = ['includeUsers', 'excludeUsers', 'includeGroups', 'excludeGroups'];
    for (const slot of slots) {
      if (!Array.isArray(users[slot])) continue;
      users[slot] = users[slot].map((v) => {
        if (!isGuid(v)) return v;
        const replacement = mapByLowerId[v.toLowerCase()];
        return replacement || v;
      });
    }
    return out;
  }

  function stripNoiseForDiff(policy) {
    const p = clone(policy);
    delete p.id;
    delete p.createdDateTime;
    delete p.modifiedDateTime;
    delete p.templateId;
    return p;
  }

  function stableSortObjectKeys(obj) {
    if (Array.isArray(obj)) return obj.map(stableSortObjectKeys);
    if (!obj || typeof obj !== 'object') return obj;
    const out = {};
    Object.keys(obj).sort().forEach((k) => {
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

  function comparePolicySets(oldPolicies, newPolicies) {
    const oldMap = new Map();
    const newMap = new Map();
    for (const p of oldPolicies || []) oldMap.set((p.displayName || '').trim(), p);
    for (const p of newPolicies || []) newMap.set((p.displayName || '').trim(), p);

    const added = [];
    const removed = [];
    const changed = [];
    const unchanged = [];

    for (const [name, oldP] of oldMap.entries()) {
      if (!newMap.has(name)) {
        removed.push(name);
        continue;
      }
      const newP = newMap.get(name);
      const left = stableSortObjectKeys(stripNoiseForDiff(oldP));
      const right = stableSortObjectKeys(stripNoiseForDiff(newP));
      const paths = diffObjects(left, right);
      if (paths.length) {
        changed.push({ name, paths });
      } else {
        unchanged.push(name);
      }
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
        state: 'enabledForReportingButNotEnforced',
        conditions: {
          users: { includeUsers: ['All'], excludeUsers: [] },
          clientAppTypes: ['exchangeActiveSync', 'other'],
          applications: { includeApplications: ['All'], excludeApplications: [] },
        },
        grantControls: {
          operator: 'OR',
          builtInControls: ['block'],
        },
      },
      {
        displayName: 'Baseline - Require MFA for admin roles',
        state: 'enabledForReportingButNotEnforced',
        conditions: {
          users: {
            includeRoles: [
              '62e90394-69f5-4237-9190-012177145e10',
              '194ae4cb-b126-40b2-bd5b-6091b380977d',
              '729827e3-9c14-49f7-bb1b-9608f156bbb8',
            ],
            excludeUsers: [],
            includeUsers: [],
            includeGroups: [],
            excludeGroups: [],
          },
          applications: { includeApplications: ['All'], excludeApplications: [] },
          clientAppTypes: ['all'],
        },
        grantControls: {
          operator: 'OR',
          builtInControls: ['mfa'],
        },
      },
      {
        displayName: 'Baseline - Require MFA for all users',
        state: 'enabledForReportingButNotEnforced',
        conditions: {
          users: {
            includeUsers: ['All'],
            excludeUsers: [],
          },
          applications: { includeApplications: ['All'], excludeApplications: [] },
          clientAppTypes: ['browser', 'mobileAppsAndDesktopClients'],
        },
        grantControls: {
          operator: 'OR',
          builtInControls: ['mfa'],
        },
      },
      {
        displayName: 'Baseline - Block device code authentication',
        state: 'enabledForReportingButNotEnforced',
        conditions: {
          users: {
            includeUsers: ['All'],
            excludeUsers: [],
          },
          applications: { includeApplications: ['All'], excludeApplications: [] },
          authenticationFlows: {
            transferMethods: ['deviceCodeFlow'],
          },
        },
        grantControls: {
          operator: 'OR',
          builtInControls: ['block'],
        },
      }
    ];
  }

  window.PolicyUtils = {
    isGuid,
    parseImportPayload,
    normalizePolicyForCreateOrUpdate,
    extractPrincipalIds,
    applyPrincipalMapping,
    comparePolicySets,
    buildStarterPolicies,
  };
})();