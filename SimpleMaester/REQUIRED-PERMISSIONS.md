# SimpleMaester - Required API Permissions

The CIS test suite (38 tests) covers Entra, Exchange Online, Forms, Teams policy and admin consent settings. SimpleMaester runs as a true browser SPA. That has implications for which APIs are actually reachable - see "Important: browser CORS limitations" below.

## What works without extra permissions

These tests work with just the default Graph scopes (User.Read, Directory.Read.All, etc.) the user already consents to on first sign-in:

- **CIS.M365.1.1.1** Cloud-only admins
- **CIS.M365.1.1.3** GA count between 2 and 4
- **CIS.M365.1.3.1** Password expiration
- **CIS.M365.2.1.11** Guest user role restriction
- **CIS.M365.5.1.2.3** Non-admin tenant creation
- **CIS.M365.5.1.2.5** Third-party app registration
- **CIS.M365.5.1.5.2** Admin consent workflow
- **CIS.M365.5.1.6.1** User consent disallowed
- **CIS.M365.5.2.2.1** Weak auth methods
- **CIS.M365.7.2.4** Third-party storage SP

## What needs extra delegated permissions (asked for at runtime)

These trigger an additional consent popup the first time the user runs the relevant test. The app registration (`4bc18533-99fb-401d-9ef1-b9e03b4ca51b`) needs the permissions added (without granting admin consent) so the user can consent on demand.

### Microsoft Graph (delegated)

| Permission | Used by |
|---|---|
| `Group.Read.All` | 1.2.1, 5.1.8.2 |
| `OrgSettings-Forms.Read.All` | 1.3.5 (internal phishing protection in Forms) |
| `OrgSettings-AppsAndServices.Read.All` | 5.5.1 (user-owned add-ins / trials) |
| `DeviceManagementConfiguration.Read.All` | 4.1 (Intune compliance default) |

### Office 365 Exchange Online (resource ID `00000002-0000-0ff1-ce00-000000000000`, delegated)

| Permission | Used by |
|---|---|
| `Exchange.Manage` | 1.2.2, 1.3.3, 1.3.6, 2.1.1-2.1.10, 2.1.14, 3.1.1, 6.5.3 (16 tests) |

The signed-in user also needs an Exchange admin role (Exchange Admin or Global Admin) for these calls to succeed at the data layer.

### Microsoft Teams

7 tests need the Teams admin REST API (`Get-CsTeamsMeetingPolicy`, `Get-CsTenantFederationConfiguration`, etc.). Microsoft does not expose these over a browser-friendly delegated PKCE flow, so SimpleMaester reports them as Skipped:

- 2.1.12, 5.1.8.1, 7.2.5, 8.5.1, 8.6.1, 8.2.1, 8.2.2

Run the equivalent `Test-MtCis*` checks in the [Maester PowerShell module](https://maester.dev) for those.

## Important: browser CORS limitations

The Exchange Online admin endpoint (`https://outlook.office365.com/adminapi/beta/...`) does **not** serve CORS headers. We confirmed this empirically - the OPTIONS preflight gets denied with no `Access-Control-Allow-Origin` header. The PowerShell module gets away with this because it isn't running in a browser.

This means **all 16 Exchange-based CIS tests report "Skipped" with a CORS message when SimpleMaester runs as a pure SPA** - which is the deployment model we picked. It is enforced by the browser (Chrome/Edge/Firefox) and there is nothing we can do about it client-side.

If you need those Exchange tests to actually run, options are:

1. Run `Test-MtCis*` in the [Maester PowerShell module](https://maester.dev) - same tests, runs locally without CORS.
2. Add a small server-side helper at the deployment origin (e.g. `https://lieben.nu/tools/SimpleMaester/api/exo`) that proxies the InvokeCommand call. Same-origin, so no CORS issue. The bearer token comes from the browser; the helper is stateless. A previous prototype lived in `serve.ps1` but was removed at user request.
3. Ship as a browser extension (different threat model, no CORS for extensions with host permissions).

## Granting consent

Once the listed permissions are added to the app registration in your tenant, sign in to SimpleMaester. The runtime consent popups appear when a test first needs a scope it doesn't have yet. For tenants that require admin consent, an admin needs to approve in `Enterprise applications > <app> > Permissions > Grant admin consent`.
