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
    'src/auth.js',
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
    'tools/Generate-TestMeta.ps1'
)

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
foreach ($script in @('src/app.js', 'src/auth.js', 'src/registry.js', 'src/tests-impl.js', 'src/tests-graph.js', 'src/tests-azure.js', 'src/tests-purview.js',
        'src/md.js', 'src/charts.js', 'src/tenantinfo.js', 'src/report-official.js', 'src/report-native.js', 'src/ui-dashboard.js')) {
    Assert-True ($indexText.Contains($script)) "index.html missing script reference: $script"
}

# The official-report exporter splices JSON between fixed markers in the bundled template
# (same contract as Get-HtmlReport.ps1). Guard against silent template drift on module sync.
$templatePath = Join-Path $root 'powershell/assets/ReportTemplate.html'
Assert-True (Test-Path $templatePath -PathType Leaf) 'Missing bundled ReportTemplate.html (official report export needs it).'
$templateText = Get-Content -Path $templatePath -Raw
Assert-True ($templateText.Contains('reportData={')) "ReportTemplate.html start marker 'reportData={' not found. Update src/report-official.js markers."
Assert-True ($templateText.Contains('EndOfJson:"EndOfJson"}')) "ReportTemplate.html end marker not found. Update src/report-official.js markers."

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
Write-Output 'SimpleZTA validation passed.'
