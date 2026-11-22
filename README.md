# LC Toolkit

Small, focused tools for Microsoft 365 and Azure admins, by [JSolve B.V.](https://jsolve.nl). Live at https://lieben.nu/tools/.

| Tool | What it does |
| --- | --- |
| ConditionalAccess | Export, import, diff and baseline Conditional Access policies |
| DupedDevices | Find duplicate Entra device registrations, keep one, retire the rest |
| InactiveDevices | Find dormant Entra devices and disable or delete them in bulk |
| InactiveGuests | Find dormant guest accounts, disable, delete or restore them |
| SPNRoleMgr | Manage API permissions of service principals, managed identities and Sites.Selected |
| SPOMove | Move or copy SharePoint files between sites or tenants |
| SPOTrim | Find stale SharePoint and OneDrive sites |
| PermView | Quick top-level permission view across Microsoft 365 and Azure |
| SimpleMaester | Maester security tests in the browser |
| SimpleZTA | Microsoft Zero Trust Assessment in the browser |

## How they work

Every tool runs in your browser. You sign in with your own account (delegated permissions, PKCE), and the page talks to Microsoft directly. There is no backend and no application consent. Write permissions are only asked for when you confirm your first change, and every change is previewed first. Sign-in tokens live in the browser tab and are gone when you close it.

## Layout

- `shared/` design system, icons, header/menu/footer shell, auth, http, log, export and the local dev server. See [shared/README.md](shared/README.md).
- `<Tool>/` one folder per tool: `index.html`, `style.css`, `src/`, `tests/`, `.auth` (client id and redirect URIs) and `serve.ps1`.
- `vendor/` pinned third-party browser libraries.

Run a tool locally with `.\<Tool>\serve.ps1` and open http://localhost:1985/.

## Self-hosting

Possible with the source in this repo. Register your own app (SPA platform, redirect URIs for your host), put its client id in each tool's `.auth`, and serve the folder with the headers from `.htaccess`. Check the [commercial use terms](https://jsolve.nl/commercial-use.html) first.

## Contributing

Issues and suggestions are welcome via GitHub issues.

## Usage of AI

AI assists (human in the loop) with writing code, documentation and reviews. It does not touch anything live.
