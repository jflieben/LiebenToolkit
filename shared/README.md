# LC Toolkit shared layer

Everything tools have in common lives here. A tool folder only holds what is specific to
that tool: its markup, its API calls, its logic and a small `style.css`.

| File | What it is |
| --- | --- |
| `toolkit.css` | Design system: tokens (light and dark), buttons, forms, cards, tabs, tables, badges, notices, progress, modals, toasts, animations. |
| `icons.svg` | Icon sprite. `t-<ToolId>` per tool and product, `i-<name>` for interface icons. 24 grid, 1.75 stroke. |
| `registry.js` | Every tool and product: id, group, tagline, descriptions, related tools. Header menu, related tools, footer and landing page are built from it. |
| `toolkit.js` | Shell and UI helpers (`window.Toolkit`). |
| `log.js` | Debug log (`window.Log`), masks tokens before anything is written. |
| `concurrency.js` | `Concurrency.pmap`, `CancelToken`, `sleep`, `eta`. |
| `http.js` | `Http.json/request/pageAll` with retry, backoff, Retry-After and readable errors. `Http.url` template encodes path parts, `Http.odataString` quotes `$filter` values. |
| `auth.js` | `ToolkitAuth.create({ scopes })`, MSAL wrapper with session-only cache and popup consent. |
| `export.js` | `Export.csv` and `Export.xlsx`, formula injection neutralised. |
| `analytics.js` | Google Analytics: loads gtag.js, `config` before `consent` (same as the rest of lieben.nu), path only. |
| `serve.ps1` | Local server for one tool, used by each tool's `serve.ps1`. |
| `tests/test.html` | Unit tests for this folder. |

## Page contract

```html
<head>
  <link rel="icon" type="image/png" href="../shared/img/favicon.png" />
  <link rel="stylesheet" href="../shared/toolkit.css" />
  <link rel="stylesheet" href="style.css" />
  <script src="../shared/registry.js"></script>
  <script src="../shared/toolkit.js"></script>          <!-- not deferred: sets the theme before paint -->
  <script src="../shared/analytics.js" defer></script>
  <script src="../shared/log.js" defer></script>
  <script src="../shared/concurrency.js" defer></script>
  <script src="../shared/http.js" defer></script>
  <script src="../shared/auth.js" defer></script>
  <script src="../shared/export.js" defer></script>      <!-- when the tool exports -->
  <script src="src/....js" defer></script>
</head>
<body data-tool="ToolId">
  <div data-tk="header"></div>
  <main class="tk-page">
    <nav class="tabs" id="tabs"> <button class="tab active" data-tab="x"><i data-icon="i-..."></i>Label</button> ... </nav>
    <section class="tab-panel active" id="tab-x"> ... </section>
    <section class="tab-panel" id="tab-debug"><div data-tk="debug"></div></section>
    <section class="tab-panel" id="tab-about"><div class="card about"> ... <div data-tk="about"></div> ... </div></section>
  </main>
  <div data-tk="related"></div>
  <div data-tk="footer"></div>
</body>
```

The shell renders on `DOMContentLoaded`, before tool code that also waits for it, so a
tool's boot can use `Toolkit.setUser`, the debug panel and hydrated icons.

## Rules

- **No inline script, no `on*=` attributes**, in HTML files or in HTML strings built by JS.
  The CSP is `script-src 'self'`. Use `addEventListener` or delegated listeners on
  `data-act` attributes. Avoid `style=` in markup; setting `el.style.x` from JS is fine.
- **Tenant data never goes into HTML unescaped.** Prefer `Toolkit.h()` or `textContent`.
  When building strings, every value goes through `Toolkit.esc()`, attributes are quoted,
  and links built from data go through `Toolkit.safeUrl()`.
- **URLs**: ids and names in API paths via `Http.url` or `encodeURIComponent`; values in
  `$filter` via `Http.odataString`. Hosts taken from API responses (SharePoint, Dataverse)
  are checked against the expected suffix before a token is requested for them.
- **Auth**: `ToolkitAuth.create`. Sign in with the read set; write scopes are requested by
  `getWriteToken()` the first time a change is confirmed. Ask the least privileged scope
  that works, and list the scopes on the connect card. Call `Toolkit.setUser(Auth.userInfo(onSignOut))`
  after `Auth.init()`. In boot: `if (ToolkitAuth.isAuthPopup()) return;` first.
- **Changes need a preview and an explicit confirm.** `Toolkit.confirm({ title, body, items, danger, okLabel })`
  lists exactly what will be touched. Selections hidden by a filter are dropped first.
  Destructive defaults are off and are not remembered across sessions.
- **Feedback**: `Toolkit.toast(msg, 'ok'|'warn'|'err'|'info')`, `Toolkit.busy(button, fn)`
  for anything async behind a button, `.progress-box` with a count and `Concurrency.eta` for long runs.
- **Exports** through `Export.csv` / `Export.xlsx` only.
- **Classes**: `btn btn-primary | btn-danger | btn-ghost | btn-small`, `btn-icon`, `card`,
  `card-head`, `tabs/tab/tab-panel`, `data-table` in `table-wrap`, `badge ok|warn|danger|info|neutral`,
  `notice notice-info|warn|err|ok` (icon first), `stat` tiles, `empty-state`, `action-bar`.
  Tokens: `--primary`, `--surface`, `--line`, `--ink-*`, `--ok|warn|danger(-soft)`, spacing `--space-1..7`.
- **Icons**: `<i data-icon="i-name"></i>` in markup, `Toolkit.icon('i-name')` or `Toolkit.iconHtml()` in JS.
  No emoji or text glyphs as icons.
- **Text**: short, factual, no marketing, no em dashes.

## Local testing

```powershell
.\InactiveGuests\serve.ps1          # http://localhost:1985/, same CSP as production
.\serve.ps1                         # landing page and shared tests (/shared/tests/test.html)
```

Headless check: `chrome --headless=new --virtual-time-budget=5000 --dump-dom http://localhost:1985/tests/test.html`
and look for `TEST-RESULT: PASS`.

## Credentials for automated test runs

`.auth` is public: the browser reads it. Test accounts go in `.testauth` next to it, which
is git-ignored, never synced to the public repo and refused by the web server.
