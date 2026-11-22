# SPOMove

Move or copy SharePoint Online files between sites, within one tenant or between two tenants.
Part of the [LC Toolkit](../index.html) by JSolve.

A client-side single-page app. It signs you in to Microsoft 365 with OAuth2 PKCE using the public multi-tenant app registration in `.auth` (`5ead7eeb-c250-472b-8e37-f59ef93e197d`) and calls Microsoft Graph and the SharePoint REST API as you. No backend; file contents go from Microsoft to your browser and back to Microsoft.

## Wizard

1. **Type**: same tenant or between two tenants.
2. **Source login**: popup sign-in to the tenant that holds the files.
3. **Source location**: **Sites** (search or paste a URL of the connected tenant), **OneDrive (by user)** or **All sites (admin)** (every site collection, SharePoint Administrator only). Then a library and optionally a subfolder.
4. **Target login**: cross-tenant only, a second popup so both connections stay live.
5. **Target location**: missing folders are created. A target that is the source folder, or inside it (or the other way round), is refused.
6. **Options**: copy or move, keep or reset metadata, previous versions, overwrite (optionally only when the source is newer), parallelism.
7. **Review and start**.

A finished job opens its report: result chart, totals, duration, average speed and every file with status and error, filterable and exportable to Excel or CSV.

## Transfer mechanics

| Mode | How files move |
|---|---|
| Same tenant | `SP.MoveCopyUtil.CopyFileByPath` / `MoveFileByPath`, server-side. A move does not bypass co-authoring locks. |
| Between tenants | Download from the source, upload to the target (chunked above 10 MB). |
| Previous versions | Every older version is replayed in order onto the target, best effort. |
| Metadata | Author, Editor, Created and Modified are set where the user resolves in the target web. |
| Overwrite off | An existing target file is skipped before anything is uploaded, versions included. |
| Deletes | Move sources, replaced targets and sync deletes go to the site recycle bin (`recycle()`), never a hard delete. |

## Jobs, resume and sync

Job records live in `localStorage` (`SPOMove-jobs`), per-file lists in IndexedDB (`SPOMove/jobItems`). No tokens are stored there; `Jobs` refuses to write anything that looks like one. Interrupted jobs resume with only the remaining files, and several jobs can run at once.

Re-run of a finished job:
- **Full re-run** transfers everything that is in the source now.
- **Sync changes** transfers new and changed files. For copy jobs it lists the target files this job created earlier that are gone from the source, and deletes them only after you confirm that list. When any source folder cannot be listed, nothing is deleted in that run and the job and report show which folders failed.

**Clear all job data** (Jobs tab, or when signing out) removes the history from the browser.

## Taking ownership

When a site denies access and you hold the SharePoint Administrator role, **Take ownership** adds you as site collection admin (`SetSiteAdmin` through the tenant admin API) after a confirmation that names the site and the account. The grant is recorded before it is made, on the job that uses the site or in the Jobs tab otherwise. It stays until you click **Remove my site admin rights** (job card, report, or Jobs tab), which removes exactly the grants SPOMove made.

## Sign-in

- MSAL token cache in `sessionStorage`: tokens are gone when the tab closes.
- Source and target are separate connections on one MSAL instance, popup only.
- After a reload a role only re-attaches the account this tab signed in for it. Resuming a job reuses a cached account only when tenant id and UPN both match the job; otherwise you sign in again with the job's UPN as login hint, and a sign-in to another tenant is refused.
- A SharePoint token is only requested for hosts of the connected tenant (`tenant`, `tenant-my`, `tenant-admin`, multi-geo variants).
- A running job's connection cannot be switched until it finishes.

## Permissions (delegated)

| Permission | Why |
|---|---|
| Graph `User.Read` | Identity and tenant detection |
| Graph `User.ReadBasic.All` | OneDrive (by user) picker, asked when first used |
| SharePoint `https://{tenant}.sharepoint.com/.default` (`AllSites.FullControl` on the app) | Read, write and delete files, set metadata, list all sites, take ownership |

## Run it locally

```powershell
.\serve.ps1               # http://localhost:1985/, same CSP as production
.\serve.ps1 -Port 1995
```

Sign-in only works on the origin in `redirect-url-local` (`http://localhost:1985`).

## Tests

`tests/test.html` (served over HTTP) covers path helpers, the source/target overlap guard, existing-target and overwrite rules, the sync delete plan (including unlisted folders), SharePoint host checks, the resume account rule, the no-token guard for stored jobs and folder enumeration that reports unlisted folders. It prints `TEST-RESULT: PASS n/n`.

## Limitations

- Cross-tenant transfers move the bytes through the browser; very large libraries are slow and memory-bound. Keep the tab open.
- Version replay and cross-tenant Author/Editor are best effort, logged per file on the Debug tab.
- SharePoint throttling (429/503) is handled by one shared backoff for all jobs.
