# SimpleZTA

SimpleZTA is a browser-based, client-side implementation of Zero Trust Assessment workflows for Microsoft 365.

## Highlights

- Runs fully client-side, no backend
- Uses delegated authentication with Microsoft Graph
- Loads official control metadata from the bundled Microsoft ZeroTrustAssessment PowerShell module
- Implements the full module control set (all controls in `./powershell/tests`) as browser-native checks
- Controls that require an admin PowerShell session with no browser-reachable delegated API (Purview /
  Exchange Online / SharePoint Online) return an explicit Skipped result with the reason
- **Dashboard** with pillar scorecards, sankey diagrams (CA/MFA coverage, auth methods for all and
  privileged users, desktop/mobile device compliance flows), tenant KPIs, a pillar × risk heatmap and a
  ranked "top fixes" panel
- **Tenant insight collection** (`src/tenantinfo.js`): browser port of the module's Invoke-ZtTenantInfo
  stage — the aggregate data behind all report graphics, stored with each scan
- **Official report export**: generates the pixel-identical Microsoft Zero Trust Assessment HTML report
  in the browser by splicing results into the bundled `ReportTemplate.html` (same contract as
  Get-HtmlReport.ps1)
- **Native HTML report export**: self-contained, print-friendly SimpleZTA report with graphics,
  findings and remediation guidance
- **Zero Trust Workshop export**: findings as a Workshop import JSON (port of
  Convert-ZtAssessmentToWorkshop)
- Result details include the module's markdown guidance ("what was checked" + remediation actions)
  and full metadata (user impact, effort, minimum license, SFI pillar)
- Stores scan history locally: per-pillar pass-score trend lines, posture metrics over time
  (MFA-protected sign-ins, phish-resistant users, compliant devices) and a scan-to-scan diff
  (regressions / fixes)
- Scan profiles (built-in "quick wins" / high-risk selections plus user-saved sets)
- Exports scan results as JSON and CSV

## Credit and copyright

The reference PowerShell module in ./powershell is Microsoft ZeroTrustAssessment and remains copyright Microsoft Corporation.
SimpleZTA is a browser companion implementation built in this toolkit.

## Local run

1. Start local server:

   pwsh -ExecutionPolicy Bypass -File .\serve.ps1

2. Open:

   http://localhost:1985

## Validation

Run:

pwsh -ExecutionPolicy Bypass -File .\tests\validate-simplezta.ps1

For a browser smoke test of the chart/report/dashboard modules (needs the local server running):

open http://localhost:1985/tests/smoke.html   (expect "SMOKE-RESULT: PASS" at the bottom)

A dashboard visual preview with fixture data is available at http://localhost:1985/tests/preview.html
(append `?theme=dark` for the dark palette).

## Control catalog

The catalog (`tests/TestMeta.json`) is generated from the bundled module — after updating `./powershell`, run:

    pwsh -ExecutionPolicy Bypass -File .\tools\Generate-TestMeta.ps1

Implementations live in `src/tests-impl.js` (core Graph), `src/tests-graph.js` (Global Secure Access, Intune, AI/agent),
`src/tests-azure.js` (Azure Resource Manager / Resource Graph), and `src/tests-purview.js` (Purview / Exchange / SharePoint).

## Notes

- Every module control has a browser implementation. Some are marked "partial" where the browser port approximates a
  complex PowerShell helper, a database export join, or a preview API surface — the UI shows what is/isn't evaluated.
- Controls that are not evaluable in browser mode (no delegated API) are marked Skipped with an explicit reason.
- Scan data is stored only in browser local storage.
