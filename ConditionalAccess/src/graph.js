// Microsoft Graph calls for ConditionalAccess.
(() => {
  const V1 = 'https://graph.microsoft.com/v1.0';
  const read = { label: 'Graph', token: () => Auth.getReadToken() };
  const write = { label: 'Graph', token: () => Auth.getWriteToken() };

  function chunk(list, size) {
    const out = [];
    for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
    return out;
  }
  const guids = (ids) => Array.from(new Set((ids || []).filter(Http.isGuid).map((x) => x.toLowerCase())));

  const listConditionalAccessPolicies = () => Http.pageAll(`${V1}/identity/conditionalAccess/policies`, read);
  const listNamedLocations = () => Http.pageAll(`${V1}/identity/conditionalAccess/namedLocations?$select=id,displayName`, read);
  const listRoleDefinitions = () => Http.pageAll(`${V1}/roleManagement/directory/roleDefinitions?$select=id,displayName,templateId`, read);

  function createConditionalAccessPolicy(body) {
    return Http.json(`${V1}/identity/conditionalAccess/policies`, { ...write, method: 'POST', body });
  }

  function updateConditionalAccessPolicy(id, body) {
    if (!Http.isGuid(id)) throw new Error(`Refusing to update policy with an invalid id: ${id}`);
    return Http.json(Http.url`${Http.raw(V1)}/identity/conditionalAccess/policies/${id}`, { ...write, method: 'PATCH', body });
  }

  // Users, groups and service principals by object id. Missing ids are simply absent.
  async function getDirectoryObjectsByIds(ids, types = ['user', 'group', 'servicePrincipal']) {
    const all = [];
    for (const part of chunk(guids(ids), 1000)) {
      const data = await Http.json(`${V1}/directoryObjects/getByIds`, { ...read, method: 'POST', body: { ids: part, types } });
      if (data && Array.isArray(data.value)) all.push(...data.value);
    }
    return all;
  }

  // Enterprise apps by appId, 15 per call (the limit of the "in" operator).
  async function servicePrincipalsByAppIds(appIds) {
    const all = [];
    for (const part of chunk(guids(appIds), 15)) {
      const filter = `appId in (${part.map(Http.odataString).join(',')})`;
      const url = `${V1}/servicePrincipals?$filter=${encodeURIComponent(filter)}&$select=id,appId,displayName`;
      all.push(...await Http.pageAll(url, read));
    }
    return all;
  }

  // Groups and directory roles of the signed-in admin, for the lockout check.
  const myMemberships = () => Http.pageAll(`${V1}/me/transitiveMemberOf?$top=999`, read);

  window.Graph = {
    listConditionalAccessPolicies,
    listNamedLocations,
    listRoleDefinitions,
    createConditionalAccessPolicy,
    updateConditionalAccessPolicy,
    getDirectoryObjectsByIds,
    servicePrincipalsByAppIds,
    myMemberships,
  };
})();
