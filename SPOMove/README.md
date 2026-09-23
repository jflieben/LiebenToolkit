# SPOMove

> Browser-based SharePoint file mover. Move or copy files between SharePoint Online sites, within one tenant or between two tenants.
> Part of the [Cloud Toolkit](../index.html) by Lieben Consulting.

SPOMove is a fully client-side single-page application. It signs you in to Microsoft 365 via OAuth2 PKCE using the public Lieben Consulting multi-tenant app registration (`271bf19c-6935-42dc-be52-67cb3bd93962`) and calls the Microsoft Graph and SharePoint REST APIs directly as the signed-in user.

**No backend. No telemetry. No file data leaves your browser.**

---

## What it does

A guided wizard walks you through a transfer:

1. **Type** - same tenant or between two tenants.
2. **Source login** - sign in to the tenant holding the files.
3. **Source location** - find the files three ways: **Sites** (search by name or paste a URL), **OneDrive (by user)** (type a person's name or UPN), or **All sites (admin)** (list every site collection in the tenant - SharePoint Administrator only). Then pick a document library and drill into a subfolder.
4. **Target login** - only for cross-tenant: a second sign-in popup so both connections stay live.
5. **Target location** - where the files should land (missing folders are created automatically).
6. **Options** - copy vs move, preserve/reset metadata, include previous versions, overwrite (optionally only when the source file is newer), parallelism.
7. **Review & start**.

When a job finishes, a **report** opens automatically: a success/failure donut chart, totals, duration and a searchable, filterable list of every file so you can drill into failures and their error messages, or export the full list as CSV.

### Transfer mechanics

| Mode | How files move |
|---|---|
| **Same tenant** | SharePoint server-side `SP.MoveCopyUtil.CopyFileByPath` / `MoveFileByPath` - bytes never pass through the browser. Fast, efficient for large libraries. |
| **Between tenants** | Download each file from the source, upload to the target (chunked upload above 10 MB). Bound by your bandwidth and memory. |
| **Previous versions** | Best-effort: each historical version is replayed in order onto the target (streams through the browser). |
| **Metadata** | Author, Editor, Created and Modified are preserved where the principal resolves in the target web (`ValidateUpdateListItem`). Turn it off to reset metadata to the current user / now. |

### Resumable jobs

Every job and its per-file progress are stored in your browser's local storage. If a run is interrupted (tab closed, network drop, cancel), open the **Jobs** tab and **Resume** - only the remaining files are processed. Multiple jobs can run at the same time.

### Taking ownership

If you cannot read a site you need and you hold the **SharePoint Administrator** role, SPOMove can temporarily add you as a site collection administrator (`SetSiteAdmin`) so the transfer can proceed.

## Required permissions (delegated)

| Permission | Why |
|---|---|
| `User.Read` | Sign-in identity and tenant detection |
| `User.ReadBasic.All` | Look up a user's OneDrive by name/UPN (requested only when the OneDrive-by-user picker is used) |
| SharePoint `AllSites.FullControl` / `Sites.FullControl.All` | Read source files, write target files, set metadata, list all sites, take ownership |

The first sign-in may trigger an admin-consent prompt if your tenant has not consented to the app yet.

## Run it locally

MSAL refuses to run from `file://`, so serve the folder over HTTP. A PowerShell helper is included:

```powershell
pwsh ./serve.ps1                 # http://localhost:1985/
pwsh ./serve.ps1 -Port 1985      # explicit port
pwsh ./serve.ps1 -NoBrowser      # don't auto-open
```

The origin must match `redirect-url-local` in `.auth` (`http://localhost:1985`) and that URI must be registered as a SPA redirect on the app registration.

## Tests

Open `tests/test.html` in a browser (served over HTTP) to run the pure-helper unit tests (path handling, OData encoding, claim login names, byte formatting, concurrency limiter).

## Limitations

- Cross-tenant transfers move file bytes through the browser; very large libraries are slow and memory-bound. Keep the tab open until the job finishes.
- Version-history replay and cross-tenant Author/Editor preservation are best-effort and logged per file in the Debug tab.
- SharePoint per-request throttling (429) is retried automatically with backoff.
