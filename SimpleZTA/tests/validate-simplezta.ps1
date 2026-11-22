[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )
    if (-not $Condition) {
        throw "Validation failed: $Message"
    }
}

Write-Output 'SimpleZTA validation starting...'

$required = @(
    'index.html',
    'style.css',
    '.auth',
    'serve.ps1',
    'src/app.js',
    'src/api.js',
    'src/registry.js',
    'src/tests-impl.js',
    'src/tests-graph.js',
    'src/tests-azure.js',
    'src/tests-purview.js',
    'src/runner.js',
    'src/storage.js',
    'src/ui-results.js',
    'src/ui-history.js',
    'src/ui-dashboard.js',
    'src/md.js',
    'src/charts.js',
    'src/tenantinfo.js',
    'src/report-official.js',
    'src/report-native.js',
    'tests/TestMeta.json',
    'tests/test.html',
    'tests/smoke.html',
    'tests/preview.html',
    'tools/Generate-TestMeta.ps1'
)

# Moved to the shared layer (tools/shared); a local copy must not come back.
foreach ($gone in @('src/auth.js', 'src/log.js', 'src/concurrency.js', 'src/export.js')) {
    Assert-True (-not (Test-Path (Join-Path $root $gone))) "Local copy of a shared module found: $gone (use ../shared instead)"
}

foreach ($r in $required) {
    Assert-True (Test-Path (Join-Path $root $r) -PathType Leaf) "Missing required file: $r"
}

# Catalog must reflect every control in the bundled PowerShell module.
$metaPath = Join-Path $root 'tests/TestMeta.json'
$meta = Get-Content -Path $metaPath -Raw | ConvertFrom-Json
$catalogIds = @($meta.PSObject.Properties.Name)
$metaCount = $catalogIds.Count

$moduleTestCount = @(Get-ChildItem -Path (Join-Path $root 'powershell/tests') -Filter 'Test-Assessment.*.ps1').Count
Assert-True ($metaCount -eq $moduleTestCount) "Catalog control count ($metaCount) does not match the PowerShell module test count ($moduleTestCount). Re-run tools/Generate-TestMeta.ps1."

# Every catalogued control must have a browser implementation across the impl files.
$implIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($file in @('src/tests-impl.js', 'src/tests-graph.js', 'src/tests-azure.js', 'src/tests-purview.js')) {
    $text = Get-Content -Path (Join-Path $root $file) -Raw
    # Match control IDs whether registered as object keys ('id':) or, in tests-purview.js,
    # as array elements ('id', / 'id']) that are bulk-registered via impl[id] = ...
    foreach ($m in [regex]::Matches($text, "'(\d{4,6})'\s*[:,\]]")) {
        [void]$implIds.Add($m.Groups[1].Value)
    }
}

$missing = @($catalogIds | Where-Object { -not $implIds.Contains($_) })
Assert-True ($missing.Count -eq 0) "Catalogued controls without a browser implementation: $($missing -join ', ')"

$orphans = @($implIds | Where-Object { $catalogIds -notcontains $_ })
Assert-True ($orphans.Count -eq 0) "Implementations without a catalog entry (orphans): $($orphans -join ', ')"

# Script references must be present in index.html.
$indexText = Get-Content -Path (Join-Path $root 'index.html') -Raw
foreach ($script in @('../shared/toolkit.js', '../shared/log.js', '../shared/http.js', '../shared/auth.js', '../shared/export.js',
        'src/app.js', 'src/api.js', 'src/registry.js', 'src/tests-impl.js', 'src/tests-graph.js', 'src/tests-azure.js', 'src/tests-purview.js',
        'src/md.js', 'src/charts.js', 'src/tenantinfo.js', 'src/report-official.js', 'src/report-native.js', 'src/ui-dashboard.js')) {
    Assert-True ($indexText.Contains($script)) "index.html missing script reference: $script"
}

# Page contract and CSP: no inline script, no inline handlers or style attributes.
Assert-True ($indexText -match '<body data-tool="SimpleZTA">') 'index.html must have <body data-tool="SimpleZTA">'
Assert-True (-not ($indexText -match '<script>|<script\s+(?![^>]*\bsrc=)[^>]*>')) 'index.html has an inline script block'
Assert-True (-not ($indexText -match '\son[a-z]+\s*=')) 'index.html has an inline on*= handler'
Assert-True (-not ($indexText -match '\sstyle\s*=')) 'index.html has a style attribute'
foreach ($file in Get-ChildItem -Path (Join-Path $root 'src') -Filter '*.js') {
    $text = Get-Content -Path $file.FullName -Raw
    Assert-True (-not ($text -match '\son[a-z]+=\s*["'']')) "src/$($file.Name) builds HTML with an inline on*= handler"
    Assert-True (-not ($text -match 'javascript:\S')) "src/$($file.Name) contains a javascript: URL"
    Assert-True (-not $text.Contains([string][char]0x2014)) "src/$($file.Name) contains an em dash"
}

# Runtime fetches of local files must stay inside what production serves under SimpleZTA:
# powershell/assets/*, powershell/tests/Test-Assessment.*.md, powershell/ZeroTrustAssessment.psd1, tests/TestMeta.json.
$allowed = '^(powershell/assets/[^/]+|powershell/tests/Test-Assessment\.[^/]+\.md|powershell/ZeroTrustAssessment\.psd1|tests/TestMeta\.json|\.auth)$'
$fetchCount = 0
foreach ($file in Get-ChildItem -Path (Join-Path $root 'src') -Filter '*.js') {
    $text = Get-Content -Path $file.FullName -Raw
    foreach ($m in [regex]::Matches($text, '[`''"]\./([^`''"]+)[`''"]')) {
        $fetchCount++
        $rel = $m.Groups[1].Value
        Assert-True ($rel -match $allowed) "src/$($file.Name) fetches ./$rel, which production does not serve"
        if ($rel -notmatch '\$\{') {
            Assert-True (Test-Path (Join-Path $root $rel) -PathType Leaf) "src/$($file.Name) fetches ./$rel, which does not exist"
        }
    }
}

# The official-report exporter splices JSON between fixed markers in the bundled template
# (same contract as Get-HtmlReport.ps1). Guard against silent template drift on module sync.
$templatePath = Join-Path $root 'powershell/assets/ReportTemplate.html'
Assert-True (Test-Path $templatePath -PathType Leaf) 'Missing bundled ReportTemplate.html (official report export needs it).'
$templateText = Get-Content -Path $templatePath -Raw
Assert-True ($templateText.Contains('reportData={')) "ReportTemplate.html start marker 'reportData={' not found. Update src/report-official.js markers."
Assert-True ($templateText.Contains('EndOfJson:"EndOfJson"}')) "ReportTemplate.html end marker not found. Update src/report-official.js markers."
$officialText = Get-Content -Path (Join-Path $root 'src/report-official.js') -Raw
Assert-True ($officialText.Contains("replace(/</g, '\\u003c')")) 'src/report-official.js must escape every "<" in the spliced JSON as \u003c'

# The workshop export and official report rely on these bundled module assets.
Assert-True (Test-Path (Join-Path $root 'powershell/assets/ztw-task-mapping.json') -PathType Leaf) 'Missing ztw-task-mapping.json (workshop export needs it).'
Assert-True (Test-Path (Join-Path $root 'powershell/ZeroTrustAssessment.psd1') -PathType Leaf) 'Missing module manifest (report version stamp needs it).'

# TenantInfo collectors must emit the exact keys the official report template consumes.
$tenantInfoText = Get-Content -Path (Join-Path $root 'src/tenantinfo.js') -Raw
foreach ($key in @('TenantOverview', 'OverviewCaMfaAllUsers', 'OverviewCaDevicesAllUsers', 'OverviewAuthMethodsAllUsers',
        'OverviewAuthMethodsPrivilegedUsers', 'DeviceOverview', 'ConfigWindowsEnrollment', 'ConfigDeviceEnrollmentRestriction',
        'ConfigDeviceCompliancePolicies', 'ConfigDeviceAppProtectionPolicies')) {
    Assert-True ($tenantInfoText.Contains($key)) "src/tenantinfo.js missing TenantInfo key: $key"
}

# Markdown guidance files must exist for a healthy majority of catalogued controls.
$mdCount = @(Get-ChildItem -Path (Join-Path $root 'powershell/tests') -Filter 'Test-Assessment.*.md').Count
Assert-True ($mdCount -gt 100) "Expected module markdown guidance files, found only $mdCount."

Write-Output "Module tests:        $moduleTestCount"
Write-Output "Catalog controls:    $metaCount"
Write-Output "Implemented controls: $($implIds.Count)"
Write-Output "Guidance md files:    $mdCount"
Write-Output "Local runtime fetches: $fetchCount (all inside the served allowlist)"
Write-Output 'SimpleZTA validation passed.'
