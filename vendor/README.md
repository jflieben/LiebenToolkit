# Vendored browser dependencies

This folder contains locally hosted copies of third-party browser libraries used by LCToolKit tools.

## Libraries and pinned versions

- `msal-browser.3.27.0.min.js` from `@azure/msal-browser@3.27.0`
- `xlsx.0.18.5.full.min.js` from `xlsx@0.18.5`
- `marked.12.0.2.min.js` from `marked@12.0.2`
- `chart.4.4.4.umd.min.js` from `chart.js@4.4.4`

## Refresh process

1. Download updated files from the upstream package CDN or release artifacts.
2. Keep version numbers in filenames.
3. Update tool HTML/JS references if filenames change.
4. Re-check CSP in `.htaccess` files after changes.
