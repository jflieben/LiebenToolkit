# Extracts the EIDSCA test definitions from the embedded Maester PowerShell module
# (../src/internal/eidsca/Test-MtEidsca*.ps1) into a single JSON file the SimpleMaester
# web app can consume without a backend.
#
# Each .ps1 follows the @templateps1.txt template, so we can scrape it with regex.

[CmdletBinding()]
param(
    [string]$EidscaSource = (Join-Path $PSScriptRoot '..\src\internal\eidsca'),
    [string]$Output       = (Join-Path $PSScriptRoot '..\data\eidsca-tests.json')
)

$ErrorActionPreference = 'Stop'

# Map PowerShell comparison operators to a small set of comparators we implement in JS.
$opMap = @{
    'eq'              = 'eq'
    'ne'              = 'ne'
    'in'              = 'in'
    'notin'           = 'notin'
    'gt'              = 'gt'
    'lt'              = 'lt'
    'ge'              = 'ge'
    'le'              = 'le'
    'BeGreaterOrEqual'= 'ge'
    'BeLessOrEqual'   = 'le'
    'BeGreaterThan'   = 'gt'
    'BeLessThan'      = 'lt'
}

function ParseRecommendedValue {
    param([string]$Raw, [string]$Operator)
    $r = $Raw.Trim()
    if ($r -match "^@\((.+)\)$") {
        $inner = $matches[1]
        # Tokens like 'a','b','c'
        $items = @()
        foreach ($m in [regex]::Matches($inner, "'([^']*)'")) { $items += $m.Groups[1].Value }
        if ($items.Count -eq 0) {
            foreach ($m in [regex]::Matches($inner, '"([^"]*)"')) { $items += $m.Groups[1].Value }
        }
        return ,$items
    }
    if ($r -match "^'([^']*)'$") { return $matches[1] }
    if ($r -match '^"([^"]*)"$') { return $matches[1] }
    if ($r -match '^-?\d+(\.\d+)?$') { return [double]$r }
    return $r
}

$tests = @()
$files = Get-ChildItem -Path $EidscaSource -Filter 'Test-MtEidsca*.ps1' | Sort-Object Name
Write-Host "Parsing $($files.Count) EIDSCA test files..."

foreach ($f in $files) {
    $code = Get-Content -Raw -Path $f.FullName
    $mdPath = [IO.Path]::ChangeExtension($f.FullName, '.md')
    $md = if (Test-Path $mdPath) { Get-Content -Raw -Path $mdPath } else { '' }

    $checkId = $null
    if ($f.BaseName -match '^Test-MtEidsca(.+)$') { $checkId = $matches[1] }
    if (-not $checkId) { continue }

    # Synopsis -> control title incl. recommended value.
    $synopsis = $null
    if ($code -match '\.SYNOPSIS\s*\r?\n\s*(.+?)\r?\n') { $synopsis = $matches[1].Trim() }

    # Description block (between .DESCRIPTION and "Queries").
    $description = $null
    if ($code -match '(?s)\.DESCRIPTION\s*\r?\n\s*(.+?)\r?\n\s*Queries ') { $description = ($matches[1].Trim() -replace '\r?\n\s*', ' ') }

    # RelativeUri + ApiVersion
    $relativeUri = $null; $apiVersion = 'v1.0'
    if ($code -match 'Invoke-MtGraphRequest\s+-RelativeUri\s+"([^"]+)"\s+-ApiVersion\s+(\w+)') {
        $relativeUri = $matches[1]
        $apiVersion = $matches[2]
    }

    # Property path, operator and recommended value
    $propertyPath = $null; $operator = $null; $recommended = $null
    if ($code -match '\$rawValue\s*=\s*\$result\.([\w\.\[\]\(\)]+)') { $propertyPath = $matches[1] }
    if ($code -match '\$testResult\s*=\s*\$tenantValue\s*-(\w+)\s+(.+?)\r?\n') {
        $operator = $matches[1]
        $recommended = ParseRecommendedValue -Raw $matches[2] -Operator $matches[1]
    }

    # Severity
    $severity = 'Info'
    if ($code -match 'Add-MtTestResultDetail.*-Severity\s+''([^'']+)''') { $severity = $matches[1] }

    # Skip reason (if any). When the .ps1 has an early-return guarded by a session variable
    # we record it so the JS runner can decide to skip elegantly when the data is missing.
    $skipIf = $null
    if ($code -match '(?s)if\s*\(\s*\$(\w+)\s+-notmatch\s+''([^'']+)''\s*\)\s*\{[^}]*Add-MtTestResultDetail\s+-SkippedBecause\s+''(\w+)''(?:\s+-SkippedCustomReason\s+''([^'']*)'')?[^}]*return') {
        $skipIf = @{ requireMatch = $matches[2]; reason = $matches[4] }
    }

    # Pull external links from the .md (we surface them in the drill-down view).
    $links = @()
    if ($md) {
        foreach ($m in [regex]::Matches($md, '\[([^\]]+)\]\((https?://[^\)]+)\)')) {
            $links += @{ text = $m.Groups[1].Value; url = $m.Groups[2].Value }
        }
    }

    # Short, rendering-friendly test description from the .md (everything before "Test script").
    $detailMd = $null
    if ($md -match '(?s)^(.+?)\r?\n#### Test script') { $detailMd = $matches[1].Trim() }

    $tests += [ordered]@{
        checkId      = $checkId
        title        = $synopsis
        description  = $description
        detailMd     = $detailMd
        severity     = $severity
        relativeUri  = $relativeUri
        apiVersion   = $apiVersion
        propertyPath = $propertyPath
        operator     = $operator
        recommended  = $recommended
        skipIf       = $skipIf
        links        = $links
        docUrl       = "https://maester.dev/docs/tests/EIDSCA.$checkId"
    }
}

$dir = Split-Path $Output -Parent
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
$tests | ConvertTo-Json -Depth 8 | Set-Content -Path $Output -Encoding UTF8
Write-Host "Wrote $($tests.Count) tests to $Output"
