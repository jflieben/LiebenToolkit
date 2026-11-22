// Auth and API clients for PermView. Each workload has its own scope set, asked by
// popup the first time that workload runs. Retry and backoff come from Http.
(() => {
  const READ = ['User.Read', 'Directory.Read.All'];
  const SETS = {
    read: READ,
    apps: [...READ, 'Application.Read.All'],
    sites: [...READ, 'Sites.Read.All'],
    files: [...READ, 'Files.Read.All'],
    calendars: [...READ, 'Calendars.Read.Shared'],
    arm: ['https://management.azure.com/user_impersonation'],
    powerbi: ['https://analysis.windows.net/powerbi/api/Workspace.Read.All'],
    powerplatform: ['https://service.powerapps.com/User'],
    devops: ['499b84ac-1321-427f-aa17-267ca6975798/user_impersonation'],
  };
  window.Auth = ToolkitAuth.create({ scopes: SETS });

  // ---------- Microsoft Graph ----------
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  function Graph(set = 'read') {
    const opts = (o) => ({ ...o, label: 'Graph', token: () => Auth.getToken(set) });
    return {
      json: (path, o = {}) => Http.json(GRAPH + path, opts(o)),
      pageAll: (path, o = {}) => Http.pageAll(GRAPH + path, opts(o)),
    };
  }

  // ---------- Azure Resource Manager ----------
  const ARM = 'https://management.azure.com';
  const armToken = () => Auth.getToken('arm');
  // Http.pageAll follows nextLink and refuses a next page on another host.
  const armPageAll = (path) => Http.pageAll(ARM + path, { label: 'ARM', token: armToken });

  // ---------- Power BI ----------
  const PBI = 'https://api.powerbi.com/v1.0/myorg';
  const PowerBI = { json: (path) => Http.json(PBI + path, { label: 'Power BI', token: () => Auth.getToken('powerbi') }) };

  // ---------- Power Platform (BAP admin API) ----------
  const BAP = 'https://api.bap.microsoft.com';
  const PowerPlatform = { json: (path) => Http.json(BAP + path, { label: 'BAP', token: () => Auth.getToken('powerplatform') }) };

  // ---------- Azure DevOps ----------
  const devopsToken = () => Auth.getToken('devops');
  function devopsUrl(url) {
    if (!Guards.isDevOpsUrl(url)) throw new Error(`Not an Azure DevOps URL: ${url}`);
    return url;
  }
  const DevOps = {
    json: (url) => Http.json(devopsUrl(url), { label: 'DevOps', token: devopsToken }),
    request: (url) => Http.request(devopsUrl(url), { label: 'DevOps', token: devopsToken }),
  };

  // ---------- Dataverse Web API ----------
  // Instance URLs come from the BAP response, so the host is checked before a token
  // is requested for it or a call is made.
  function dataverseBase(instanceUrl) {
    const origin = Guards.dataverseOrigin(instanceUrl);
    if (!origin) throw new Error(`Dataverse instance URL is not an https *.dynamics.com address: ${instanceUrl}`);
    return origin;
  }
  const DV_HEADERS = { 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' };
  function dataversePageAll(instanceUrl, path, onPage) {
    const origin = dataverseBase(instanceUrl);
    return Http.pageAll(`${origin}/api/data/v9.2${path}`, {
      label: 'Dataverse', headers: DV_HEADERS, token: () => Auth.getToken([`${origin}/user_impersonation`]),
    }, onPage);
  }

  window.Api = {
    SETS, Graph, armPageAll, PowerBI, PowerPlatform, DevOps, dataversePageAll, dataverseBase,
  };
})();
