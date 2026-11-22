# Vendored browser dependencies

This folder contains locally hosted copies of third-party browser libraries used by LCToolKit tools.

## Libraries and pinned versions

- `msal-browser.3.27.0.min.js` from `@azure/msal-browser@3.27.0`
- `xlsx.0.20.3.full.min.js` from SheetJS 0.20.3 (https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js, sha256 cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41). SheetJS no longer publishes to npm; 0.20.2+ fixes CVE-2023-30533 and CVE-2024-22363
- `marked.12.0.2.min.js` from `marked@12.0.2`
- `chart.4.4.4.umd.min.js` from `chart.js@4.4.4`

## Refresh process

1. Download updated files from the upstream package CDN or release artifacts.
2. Keep version numbers in filenames.
3. Update tool HTML/JS references if filenames change.
4. Re-check CSP in `.htaccess` files after changes.
