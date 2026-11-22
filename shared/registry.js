// Every tool and product the toolkit links to. readOnly: the tool never changes the tenant.
// privacyNote: data that goes somewhere other than Microsoft, shown in About and footer.
// The header menu, related tools,
// footer and the landing page are all built from this list.
(() => {
  const TOOLKIT = {
    name: 'LC Toolkit',
    owner: 'JSolve B.V.',
    ownerUrl: 'https://jsolve.nl',
    softwareUrl: 'https://jsolve.nl/software.html',
    commercialUrl: 'https://jsolve.nl/commercial-use.html',
    privacyUrl: 'https://jsolve.nl/privacy.html',
    repoUrl: 'https://github.com/jflieben/LiebenToolkit',
    contact: 'hello@jsolve.nl',
  };

  const GROUPS = [
    { id: 'identity', label: 'Identity and devices' },
    { id: 'sharepoint', label: 'SharePoint and OneDrive' },
    { id: 'assess', label: 'Assessments and permissions' },
  ];

  const TOOLS = [
    {
      id: 'ConditionalAccess', group: 'identity',
      tagline: 'Export, import and compare Conditional Access policies',
      short: 'Export, import, diff and baseline CA policies.',
      description: 'Import, export and compare Conditional Access policies. Prefix or suffix names, pick the state per policy, map missing users and groups, and start from a baseline set.',
      related: ['SimpleZTA', 'SimpleMaester', 'InactiveGuests'],
    },
    {
      id: 'DupedDevices', group: 'identity',
      tagline: 'Find and clean up duplicate Entra device registrations',
      short: 'Duplicate Entra devices, keep one, retire the rest.',
      description: 'Find duplicate device registrations in Entra ID, pick the one to keep and disable or delete the rest. Optional grace period via an extension attribute: disable now, delete safely later.',
      related: ['InactiveDevices', 'InactiveGuests', 'ConditionalAccess'],
    },
    {
      id: 'InactiveDevices', group: 'identity',
      tagline: 'Find and clean up dormant devices in Entra',
      short: 'Dormant Entra devices, filtered and bulk-handled.',
      description: 'Find dormant Entra devices by inactivity, OS, join type, Autopilot and managed state. Bulk disable, enable or delete what you select.',
      related: ['DupedDevices', 'InactiveGuests', 'SimpleMaester'],
    },
    {
      id: 'InactiveGuests', group: 'identity',
      tagline: 'Find and clean up dormant guest accounts in Entra',
      short: 'Dormant guests: disable, delete or restore.',
      description: 'Find dormant guest accounts in Entra ID by inactivity, redemption state and account age. Bulk disable, delete or restore them from the recycle bin.',
      related: ['InactiveDevices', 'PermView', 'ConditionalAccess'],
    },
    {
      id: 'SPNRoleMgr', group: 'identity',
      tagline: 'Manage API permissions of service principals and managed identities',
      short: 'API permissions for SPNs, MIs and Sites.Selected.',
      description: 'Inspect or change Graph and other API permissions for any service principal or managed identity, and manage Sites.Selected grants on SharePoint sites.',
      related: ['PermView', 'M365Permissions', 'SPOTrim'],
    },
    {
      id: 'SPOMove', group: 'sharepoint',
      tagline: 'Move or copy SharePoint files between sites and tenants',
      short: 'Move or copy files between sites or tenants.',
      description: 'Move or copy SharePoint files between sites, within one tenant (server-side copy) or between two tenants. Wizard driven, resumable jobs, optional versions and metadata.',
      related: ['SPOTrim', 'PermView', 'VFS365'],
    },
    {
      id: 'SPOTrim', group: 'sharepoint',
      tagline: 'Find stale SharePoint and OneDrive sites',
      short: 'Stale SharePoint and OneDrive sites, explained.',
      description: 'Find stale SharePoint and OneDrive sites, see who owns them and when they were last changed or viewed, and how that verdict was reached.',
      related: ['SPOMove', 'PermView', 'InactiveGuests'],
    },
    {
      id: 'PermView', group: 'assess', readOnly: true,
      tagline: 'Quick top-level permission view across Microsoft 365 and Azure',
      short: 'Top-level permissions across M365 and Azure.',
      description: 'Quick top-level permission view across Entra, SharePoint, OneDrive, mailboxes, Azure, Azure DevOps, Power BI and Power Platform. The first look before a full audit.',
      related: ['SPNRoleMgr', 'SPOTrim', 'M365Permissions'],
    },
    {
      id: 'SimpleMaester', group: 'assess', readOnly: true,
      privacyNote: 'Domain names are looked up through Cloudflare DNS-over-HTTPS for the SPF, DKIM and DMARC tests.',
      tagline: 'Maester security tests, straight from your browser',
      short: 'Maester tests in the browser, with trends.',
      description: 'Browser port of Maester. Run the Microsoft 365 security tests (EIDSCA, CISA, CIS, conditional access) without PowerShell and follow your posture over time.',
      related: ['SimpleZTA', 'ConditionalAccess', 'DriftMaester'],
    },
    {
      id: 'SimpleZTA', group: 'assess', readOnly: true,
      tagline: 'Zero Trust Assessment, straight from your browser',
      short: 'Zero Trust Assessment in the browser.',
      description: 'Browser port of the Microsoft Zero Trust Assessment. Run the checks with delegated access, compare with earlier runs and drill into the evidence.',
      related: ['SimpleMaester', 'ConditionalAccess', 'AzCmply'],
    },
  ];

  const PRODUCTS = [
    { id: 'M365Permissions', url: 'https://m365permissions.com', host: 'm365permissions.com', kind: 'Hosted',
      description: 'Know who has what access in your tenant, at any level, at any point in time. The full audit PermView is the quick look for.' },
    { id: 'AzCmply', url: 'https://azcmply.jsolve.nl/', host: 'azcmply.jsolve.nl', kind: 'Browser',
      description: 'Azure subscription assessment against MCSB, CIS, NIST, ISO, DORA, WAF, PCI, SOC 2 and Azure LZ, with a full report and trend over time.' },
    { id: 'DriftMaester', url: 'https://github.com/jflieben/DriftMaester', host: 'github.com', kind: 'GitHub',
      description: 'Plug and play Azure Automation wrapper for Maester: continuous drift detection, alerting and HTML reporting.' },
    { id: 'M365AutoLink', url: 'https://github.com/jflieben/M365AutoLink', host: 'github.com', kind: 'GitHub',
      description: 'Discovers SharePoint and Teams libraries and creates OneDrive shortcuts, with exclusions and tray-based run controls.' },
    { id: 'M365AutoRevocate', url: 'https://github.com/jflieben/M365AutoRevocate', host: 'github.com', kind: 'GitHub',
      description: 'Offboarding cleanup for Microsoft 365. Reacts to inactive, disabled or deleted users and runs the revocations you enable, via a managed identity.' },
    { id: 'VFS365', url: 'https://github.com/jflieben/VFS365', host: 'github.com', kind: 'GitHub',
      description: 'Shows SharePoint, Teams and OneDrive libraries as a drive and in Explorer, with each user\'s own permissions, so legacy apps keep working on ordinary file paths. No sync, no per-library setup.' },
    { id: 'WeatherRadius', url: 'https://weatherradius.com', host: 'weatherradius.com', kind: 'Hosted',
      description: 'Pick a starting point and a date range and find the best weather within driving or public transport distance.' },
  ];

  const byId = (id) => TOOLS.find((t) => t.id === id) || PRODUCTS.find((p) => p.id === id) || null;

  window.ToolkitRegistry = { TOOLKIT, GROUPS, TOOLS, PRODUCTS, byId };
})();
