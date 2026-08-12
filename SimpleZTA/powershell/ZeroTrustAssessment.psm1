#Requires -Version 7.0
#Requires -PSEdition Core,Desktop
#Requires -Assembly lib\DuckDB.NET.Data.dll
#Requires -Module @{'GUID'='8028b914-132b-431f-baa9-94a6952f21ff';'ModuleName'='PSFramework';'ModuleVersion'='1.13.419'}

<#
.SYNOPSIS
    ZeroTrustAssessment
.DESCRIPTION
    Perform a Zero Trust Assessment of your Microsoft 365 environment.
.NOTES
    ModuleVersion: 2.4.26-preview
    GUID: 708723ef-2420-4bcb-bfd7-988e190d7acf
    Author: Microsoft
    CompanyName: Microsoft
    Copyright: (c) Microsoft. All rights reserved.
.FUNCTIONALITY
    Clear-ZtRequiredModule, Connect-ZtAssessment, Disconnect-ZtAssessment, Get-ZtCurrentLicense, Get-ZtExportStatistics, Get-ZtGraphScope, Get-ZtTest, Get-ZtTestStatistics, Invoke-ZtAssessment, Invoke-ZtAzureRequest, Invoke-ZtAzureResourceGraphRequest, Invoke-ZtGraphRequest
.LINK
    https://github.com/microsoft/zerotrustassessment
#>

#region NestedModules Script(s)

#endregion

# To give a module-wide constant point of reference
$script:ModuleRoot = $PSScriptRoot
[string[]] $script:ConnectedService = @()
[string[]] $script:CurrentLicense = @()

# Load PowerShell Classes
foreach ($file in Get-ChildItem -Path "$script:ModuleRoot\classes" -Recurse -Filter "*.ps1") {
	try { . $file.FullName }
	catch { Write-PSFMessage -Level Error -Message "Failed to import file {0}" -StringValues $file.FullName -ErrorRecord $_ -Target $file }
}

# Load Non-Public commands
foreach ($file in Get-ChildItem -Path "$script:ModuleRoot\private" -Recurse -Filter "*.ps1") {
	try { . $file.FullName }
	catch { Write-PSFMessage -Level Error -Message "Failed to import file {0}" -StringValues $file.FullName -ErrorRecord $_ -Target $file }
}

# Load Public commands
foreach ($file in Get-ChildItem -Path "$script:ModuleRoot\public" -Recurse -Filter "*.ps1") {
	try { . $file.FullName }
	catch { Write-PSFMessage -Level Error -Message "Failed to import file {0}" -StringValues $file.FullName -ErrorRecord $_ -Target $file }
}

# Execute Startup scripts
foreach ($file in Get-ChildItem -Path "$script:ModuleRoot\scripts" -Recurse -Filter "*.ps1") {
	try { . $file.FullName }
	catch { Write-PSFMessage -Level Error -Message "Failed to import file {0}" -StringValues $file.FullName -ErrorRecord $_ -Target $file }
}

# Ready the Tests
foreach ($file in Get-ChildItem -Path "$script:ModuleRoot\tests" -Recurse -Filter "*.ps1") {
	try { . $file.FullName }
	catch { Write-PSFMessage -Level Error -Message "Failed to import file {0}" -StringValues $file.FullName -ErrorRecord $_ -Target $file }
}


