# SPOTrim

Finds stale SharePoint and OneDrive sites, shows who owns them, when they were last changed or viewed and how that verdict was reached. Part of the [LC Toolkit](https://github.com/jflieben/LCToolKit).

Runs in the browser. It signs you in with MSAL (PKCE, delegated) using the client id in [.auth](.auth) and calls Microsoft Graph and SharePoint REST as you. Shared code (shell, auth, http, log, export) lives in `../shared`; see [../shared/README.md](../shared/README.md).

## Recommendations

For every selected site the tool collects these signals (all visible in the row detail and the export):

| Source | Signal |
|---|---|
| Discovery | `lastModifiedDateTime` |
| `_api/web` | `LastItemUserModifiedDate` (no system or crawl updates) |
| `_api/web/lists` | latest `LastItemUserModifiedDate` of the visible document libraries |
| Graph `/sites/{id}/pages` | last edited modern page or news post |
| Graph `/sites/{id}/analytics/allTime` and `lastSevenDays` | access count and distinct actors |
| Graph `drive/items/root/getActivitiesByInterval` | daily activity for 30 days, gives the last active day |
| Graph `drive?$select=quota` | storage used |

The Graph signals go in one `$batch` per site.

- **Keep**: modified within the threshold, or a system site.
- **Review**: not modified within the threshold but still opened (30 or 7 days, or all-time views).
- **Clean up**: not modified within the threshold and no views, or not viewed within the threshold.
- **No access**: `_api/web` and `_api/web/lists` both 401/403.
- **Unknown**: no modification date for another reason.

Exports (Excel, CSV) contain the rows shown, with every raw signal, so you can re-score in Excel.

## Permissions

| When | Resource | Scope |
|---|---|---|
| Sign-in | Graph | `User.Read`, `Sites.Read.All` |
| First discovery | SharePoint (`{tenant}.sharepoint.com`) | `AllSites.Read` |
| OneDrive discovery and force scan only | SharePoint admin (`{tenant}-admin.sharepoint.com`) | `AllSites.FullControl`, plus the SharePoint Administrator role |

Missing consent is asked before a run starts, in one popup for both SharePoint resources when both are needed. During a run tokens are only taken silently (`noRedirect`), so the page never navigates away halfway.

## Force scan

Makes the signed-in admin temporary site collection admin on selected sites they cannot read.

1. Checks `_api/web/currentuser` on every selected site. `IsSiteAdmin=false` or a 403 means not admin. Anything else (401, 429 after retries, 5xx, network) is unknown: that site is left alone and listed after the run.
2. Shows the exact list of sites, the account and the tenant, and waits for confirmation. OneDrives are skipped unless ticked.
3. Writes each site to `localStorage` (`spotrim-elevations`: site URL, account, tenant id, time, no tokens) before calling `SetSiteAdmin`. Clearly refused grants are taken off the list again.
4. Removes the rights after the analysis, also after cancel or errors. Failed removals stay on the list.
5. On the next visit a notice offers to remove what an interrupted run left behind.

## Run it locally

```powershell
pwsh ./serve.ps1             # http://localhost:1985/, same CSP as production
pwsh ./serve.ps1 -Port 1996
```

Sign-in only works on the port in `redirect-url-local` of [.auth](.auth).

## Tests

[tests/test.html](tests/test.html): classification, host checks, export rows, Graph batch parsing, scoring, and the force scan safety rules (admin check, plan, ledger, grant and revoke). Open `/tests/test.html` on the local server and look for `TEST-RESULT: PASS`.

## Files

```
index.html          page (shared shell contract)
style.css           tool-only styles
src/sites.js        site types, tenant host checks, export rows
src/elevation.js    force scan: admin check, plan, ledger, grant, revoke
src/graph.js        Graph: host detection, engagement $batch
src/sharepoint.js   SharePoint REST, search, tenant admin API, SetSiteAdmin
src/analysis.js     signals per site and the recommendation
src/app.js          UI
```
