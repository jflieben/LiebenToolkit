# SPOTrim

> Browser-based SharePoint &amp; OneDrive stale-site analysis &amp; cleanup tool.
> Part of the [LCToolKit](https://github.com/jflieben/LCToolKit) by Lieben Consultancy.

SPOTrim is a fully client-side single-page application. It signs you in to your Microsoft 365 tenant via OAuth2 PKCE using the public Lieben Consultancy multi-tenant app registration (`271bf19c-6935-42dc-be52-67cb3bd93962`) and queries the Microsoft Graph &amp; SharePoint REST APIs as the signed-in user.

**No backend. No telemetry. No data leaves your browser.**

---

## Feature 1 - Stale site analysis &amp; cleanup advice

For every selected SharePoint or OneDrive site, SPOTrim collects multiple "last user activity" signals and produces a single recommendation: **Clean up**, **Review**, **Keep**, or **Unknown**.

Signals collected (visible in the per-row "Why" tooltip):

| Source | Signal |
|---|---|
| Microsoft Graph `/sites/{id}` | `lastModifiedDateTime` |
| SharePoint REST `_api/web` | `LastItemUserModifiedDate` (excludes system/crawl updates) |
| SharePoint REST `_api/web/lists` | max `LastItemUserModifiedDate` across non-hidden document libraries |
| Graph `/sites/{id}/pages` | most recent modern page / news edit (counts as a modification signal too) |
| Graph `/sites/{id}/analytics/allTime` | all-time access count + distinct actors |
| Graph `/sites/{id}/analytics/lastSevenDays` | last 7 days access count |
| Graph `/sites/{id}/drive/items/root/getActivitiesByInterval` | per-day drive activity for the last 30 days (yields a precise "last viewed" date) |
| Graph `/sites/{id}/drive?$select=quota` | live storage usage in bytes |

All view / engagement signals come from `graph.microsoft.com` (CORS-friendly) and are batched into a single Graph `$batch` call per site - no proxy needed.

The site is classified as:
- **Clean up** - not modified in &gt; threshold AND no all-time views, OR not viewed for &gt; threshold.
- **Review** - not modified in &gt; threshold but still has activity in the last 30 / 7 days OR has nonzero all-time views (archive / reference site).
- **Keep** - modified within the threshold.
- **No access** - both `_api/web` and `_api/web/lists` returned 401/403.
- **Unknown** - no modification signal could be retrieved for non-permission reasons (e.g. transient 5xx).

Results are exportable to **XLSX** or **CSV** with all raw signals (so you can re-score in Excel with your own rules).

## Required Microsoft Graph delegated permissions

| Permission | Why |
|---|---|
| `Sites.Read.All` | Enumerate sites; per-site analytics, drive activity, site pages, drive quota |
| `Sites.FullControl.All` | Read web/list modification details via SharePoint REST |
| `User.Read` | Sign-in identity |

The Lieben Consultancy multi-tenant app already exposes these scopes; the first sign-in will trigger an admin-consent prompt if your tenant has not yet consented to this app.

> **OneDrive enumeration:** to enumerate other users' OneDrive sites the signed-in user must additionally hold the **SharePoint Administrator** role in Entra ID (the SharePoint Admin Tenant REST API is used; it is CORS-friendly and works directly from the browser).

## Run it

A PowerShell helper is included that serves the folder over HTTP (MSAL refuses `file://`):

```powershell
pwsh ./serve.ps1                 # http://localhost:8080/
pwsh ./serve.ps1 -Port 8123      # custom port
pwsh ./serve.ps1 -NoBrowser      # don't auto-open
```

The script reads `redirect-url-local` from [.auth](.auth) and warns if the chosen port does not match. The tool itself reads [.auth](.auth) at startup and picks `redirect-url-local` when served from `localhost`/`127.0.0.1` and `redirect-url-web` otherwise - so the same build works both locally and in production.

The first time you sign in, MSAL caches the refresh token in `localStorage` so subsequent loads restore your session. A "Sign out" button is in the top-right when signed in.

## Tests

Unit tests for pure logic (`.auth` parsing, redirect-URI selection, staleness scoring, pmap concurrency, classify) are at [tests/test.html](tests/test.html). Run them by serving the folder and opening <http://localhost:8080/tests/test.html>.

## Architecture

```
SPOTrim/
  index.html           # SPA shell with tabs (Discover, Results, Debug, About)
  style.css            # White + blue UI, dark mode, responsive
  src/
    log.js             # Centralized log buffer + Debug tab renderer
    auth.js            # MSAL.js wrapper (PKCE), per-resource token cache mutex
    concurrency.js     # Promise-based pmap with cancellation + ETA progress
    graph.js           # Microsoft Graph client (retry/backoff, paging, per-site engagement $batch)
    sharepoint.js      # SharePoint REST client + Search-based site enumeration
    analysis.js        # Per-site signal aggregation & recommendation logic
    export.js          # XLSX (SheetJS) and CSV exporters
    app.js             # UI orchestration: tabs, discover, analyze, results
```

External libraries (loaded via CDN):
- [`msal-browser`](https://github.com/AzureAD/microsoft-authentication-library-for-js) - OAuth2 PKCE
- [`SheetJS / xlsx`](https://github.com/SheetJS/sheetjs) - XLSX export

## Roadmap

The `.description` file in this folder defines additional features to be implemented. Feature 1 (stale-site analysis) is implemented; future features (configuring versioning, file-version cleanup, deletion workflows) will be added incrementally.
