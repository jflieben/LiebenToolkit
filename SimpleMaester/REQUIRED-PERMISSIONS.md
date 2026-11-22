# SimpleMaester - required API permissions

SimpleMaester is a browser SPA (MSAL PKCE, delegated, read only). Scopes are defined in `src/config.js`.
The app registration (`4bc18533-99fb-401d-9ef1-b9e03b4ca51b`) needs every permission below added
(admin consent optional) so users can consent on demand.

## Asked at sign-in (Microsoft Graph, delegated)

`User.Read`, `Directory.Read.All`, `Policy.Read.All`, `RoleManagement.Read.Directory`,
`Application.Read.All`, `UserAuthenticationMethod.Read.All`, `AuditLog.Read.All`, `Reports.Read.All`.

## Asked before a scan whose selection needs them

The runner collects the `requiredScopes` of the selected tests and gets one token per resource
before the first test runs (popup, redirect when the popup is blocked; the run resumes after the redirect).

| Permission | Used by |
| --- | --- |
| Graph `OrgSettings-Forms.Read.All` | CIS 1.3.5 (Forms phishing protection) |
| Graph `OrgSettings-AppsAndServices.Read.All` | CIS 5.5.1 (user-owned apps and trials) |
| Graph `SecurityEvents.Read.All` | Secure Score based tests |
| Graph `SharePointTenantSettings.Read.All` | CISA SharePoint 1.1, 1.3 |
| Graph `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All`, `DeviceManagementServiceConfig.Read.All` | Intune tests |
| Graph `AccessReview.Read.All` | Access review tests |
| Graph `EntitlementManagement.Read.All` | Entitlement management tests |
| Graph `InformationProtectionPolicy.Read` | Purview sensitivity label test |
| Azure Service Management `user_impersonation` (requested as `https://management.azure.com/.default`) | Azure RBAC tests, CISA AAD 4.1 (diagnostic settings) |

ARM has no read-only delegated scope. The token carries the user's Azure RBAC rights; SimpleMaester
only sends GET requests with it, and only asks for it when an Azure test is in the selection.

## Not available from a browser

- **Exchange Online admin REST** (`outlook.office365.com/adminapi`): no CORS headers, so the Exchange
  based CIS and CISA tests report Skipped. SimpleMaester does not request an Exchange token.
- **Teams admin REST**: not exposed for delegated browser auth; the Teams tests report Skipped.

Run those with the [Maester PowerShell module](https://maester.dev).

## Other outbound calls

The SPF, DKIM and DMARC tests query `https://cloudflare-dns.com/dns-query` (DNS-over-HTTPS) for the
tenant's mail domains. No token is sent. The site CSP must allow this origin in `connect-src`.
