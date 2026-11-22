// Microsoft Graph calls for InactiveGuests. Guests come from /beta/users because
// signInActivity is reliable there; everything else uses v1.0.
(() => {
  const V1 = 'https://graph.microsoft.com/v1.0';
  const BETA = 'https://graph.microsoft.com/beta';

  function call(url, { write = false, ...opts } = {}) {
    return Http.json(url, { ...opts, label: 'Graph', token: write ? Auth.getWriteToken : Auth.getReadToken });
  }

  const GUEST_FIELDS = [
    'id', 'userPrincipalName', 'displayName', 'mail', 'userType',
    'externalUserState', 'externalUserStateChangeDateTime',
    'createdDateTime', 'creationType', 'accountEnabled', 'companyName', 'signInActivity',
  ].join(',');

  // onProgress(loadedCount, pageNumber)
  function listAllGuests(onProgress, signal) {
    const url = `${BETA}/users?$filter=userType eq 'Guest'&$select=${GUEST_FIELDS}&$top=999`;
    return Http.pageAll(url, { label: 'Graph', token: Auth.getReadToken, headers: { ConsistencyLevel: 'eventual' }, signal },
      (items, page, total) => onProgress && onProgress(total, page));
  }

  function listDeletedGuests(onProgress) {
    const select = 'id,displayName,userPrincipalName,mail,userType,accountEnabled,deletedDateTime';
    const url = `${V1}/directory/deletedItems/microsoft.graph.user?$filter=userType eq 'Guest'&$select=${select}&$top=999`;
    return Http.pageAll(url, { label: 'Graph', token: Auth.getReadToken, headers: { ConsistencyLevel: 'eventual' } },
      (items, page, total) => onProgress && onProgress(total, page));
  }

  const disableUser = (id) => call(Http.url`${Http.raw(V1)}/users/${id}`, { method: 'PATCH', body: { accountEnabled: false }, write: true });
  const enableUser = (id) => call(Http.url`${Http.raw(V1)}/users/${id}`, { method: 'PATCH', body: { accountEnabled: true }, write: true });
  const deleteUser = (id) => call(Http.url`${Http.raw(V1)}/users/${id}`, { method: 'DELETE', write: true });
  const restoreUser = (id) => call(Http.url`${Http.raw(V1)}/directory/deletedItems/${id}/restore`, { method: 'POST', write: true });

  window.Graph = { call, listAllGuests, listDeletedGuests, disableUser, enableUser, deleteUser, restoreUser };
})();
