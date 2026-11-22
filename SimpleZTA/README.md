# SimpleZTA

Browser port of the Microsoft Zero Trust Assessment for Microsoft 365. Part of the LC Toolkit; shell,
auth, HTTP, log and export come from `../shared` (see `../shared/README.md`).

## Credit and copyright

The controls, their metadata, the remediation guidance and the official report template come from the
Microsoft [ZeroTrustAssessment](https://github.com/microsoft/zerotrustassessment) PowerShell module, bundled
in `./powershell`. Copyright Microsoft Corporation, MIT licence. SimpleZTA is an independent browser port,
not made or supported by Microsoft.

## Highlights

- Runs client-side, no backend. Delegated Microsoft Graph; Azure Resource Manager only when a run contains
  Azure controls (separate `arm` scope set, asked at run start, never at sign-in).
- Every module control (`./powershell/tests`) has a browser implementation. Controls that need an admin
  PowerShell session with no delegated API (Purview, Exchange Online, SharePoint Online) return Skipped with
  the reason.
- Dashboard: pillar scores, sankeys (CA/MFA coverage, auth methods for all and privileged users, desktop and
  mobile device compliance), tenant KPIs, pillar by risk heatmap, top fixes. Charts use the toolkit tokens,
  so they follow light and dark mode.
- Tenant insight collection (`src/tenantinfo.js`): port of the module's Invoke-ZtTenantInfo stage.
- Official report export: splices the results into the bundled `ReportTemplate.html`, same contract as
  Get-HtmlReport.ps1. The spliced JSON has every `<` written as `\u003c`.
- Native HTML report, Zero Trust Workshop JSON, Excel, CSV and JSON exports.
- Scan history in `localStorage`: trends, posture metrics over time and a scan-to-scan diff. Stored API
  responses are capped (25 calls per control, 1000 characters each). History, Clear history removes it.
- Scan profiles (built-in plus saved ones).

## Local run

    pwsh -ExecutionPolicy Bypass -File .\serve.ps1          # http://localhost:1985/

## Tests

- `pwsh -ExecutionPolicy Bypass -File .\tests\validate-simplezta.ps1`: catalog versus module, page contract,
  CSP rules, runtime fetch allowlist, report template markers.
- `http://localhost:1985/tests/test.html`: unit tests, expect `TEST-RESULT: PASS n/n`.
- `http://localhost:1985/tests/smoke.html`: charts, dashboard and both report exports against the real
  template, expect `SMOKE-RESULT: PASS`.
- `http://localhost:1985/tests/preview.html`: dashboard with fixture data, `?theme=dark` for dark mode.

## Runtime files

Production serves only these local files besides the app itself: `powershell/assets/*`,
`powershell/tests/Test-Assessment.*.md`, `powershell/ZeroTrustAssessment.psd1` and `tests/TestMeta.json`.
The validate script fails when code fetches anything else.

## Control catalog

`tests/TestMeta.json` is generated from the bundled module. After updating `./powershell`:

    pwsh -ExecutionPolicy Bypass -File .\tools\Generate-TestMeta.ps1

Implementations: `src/tests-impl.js` (core Graph), `src/tests-graph.js` (Global Secure Access, Intune,
AI/agent), `src/tests-azure.js` (Azure Resource Manager / Resource Graph), `src/tests-purview.js`
(Purview / Exchange / SharePoint). Controls marked "partial" approximate a complex PowerShell helper, a
database export join or a preview API; the UI says what is and is not evaluated.
