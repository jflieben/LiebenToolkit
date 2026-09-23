<#
.SYNOPSIS
    Regenerates tests/TestMeta.json from the bundled ZeroTrustAssessment PowerShell
    module by reading the [ZtTest(...)] attribute of every Test-Assessment-<id> function.

.DESCRIPTION
    This is the single source of truth for the SimpleZTA control catalog. It AST-parses
    every powershell/tests/Test-Assessment.*.ps1, extracts the ZtTest metadata attribute,
    and writes one JSON object per control, keyed by TestId, sorted numerically.

    Run this whenever the powershell/ module is updated so the browser catalog never drifts.

.EXAMPLE
    pwsh -NoProfile -ExecutionPolicy Bypass -File .\tools\Generate-TestMeta.ps1
#>
[CmdletBinding()]
param(
    [string] $Root = (Split-Path -Parent $PSScriptRoot),
    [switch] $CheckOnly
)

$ErrorActionPreference = 'Stop'
$testRoot = Join-Path $Root 'powershell/tests'
$outPath = Join-Path $Root 'tests/TestMeta.json'

# Fields we surface in the catalog. Multi-value fields stay arrays; the rest are scalars.
$multiValue = @('TenantType', 'Pillar', 'Service', 'MinimumLicense', 'CompatibleLicense')
$fields = @('Category', 'ImplementationCost', 'MinimumLicense', 'CompatibleLicense', 'Service',
    'Pillar', 'RiskLevel', 'SfiPillar', 'TenantType', 'TestId', 'Title', 'UserImpact')

$files = Get-ChildItem -Path $testRoot -Filter 'Test-Assessment.*.ps1' | Sort-Object Name
$catalog = [ordered]@{}

foreach ($file in $files) {
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
    $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) |
        Where-Object { $_.Name -like 'Test-Assessment-*' } | Select-Object -First 1
    if (-not $fn) {
        Write-Warning "No Test-Assessment function in $($file.Name)"
        continue
    }

    $attr = $fn.Body.ParamBlock.Attributes | Where-Object { $_.TypeName.FullName -eq 'ZtTest' }
    $idFromFile = ($file.BaseName -replace 'Test-Assessment\.', '')

    $entry = [ordered]@{}
    $values = @{}
    if ($attr) {
        foreach ($na in $attr.NamedArguments) {
            $name = $na.ArgumentName
            try { $val = $na.Argument.SafeGetValue() } catch { $val = $null }
            $values[$name] = $val
        }
    }

    foreach ($field in ($fields | Sort-Object)) {
        if (-not $values.ContainsKey($field)) { continue }
        $val = $values[$field]
        if ($null -eq $val) { continue }
        if ($multiValue -contains $field) {
            $arr = @($val) | ForEach-Object { "$_" } | Where-Object { $_ -ne '' }
            if ($arr.Count -gt 0) { $entry[$field] = $arr }
        }
        else {
            $entry[$field] = "$val"
        }
    }

    $id = if ($entry.Contains('TestId') -and $entry['TestId']) { "$($entry['TestId'])" } else { $idFromFile }
    $entry['TestId'] = $id
    if (-not $entry.Contains('Title') -or -not $entry['Title']) { $entry['Title'] = "Test $id" }

    $catalog[$id] = $entry
}

# Sort numerically by TestId for a stable, reviewable file.
$sorted = [ordered]@{}
foreach ($key in ($catalog.Keys | Sort-Object { [int]$_ })) { $sorted[$key] = $catalog[$key] }

$json = $sorted | ConvertTo-Json -Depth 6
if ($CheckOnly) {
    Write-Output "Parsed $($sorted.Count) controls."
    return
}

# UTF-8 without BOM
[System.IO.File]::WriteAllText($outPath, $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Output "Wrote $($sorted.Count) controls to $outPath"
