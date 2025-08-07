# Authentication Functions to Extract from Legacy Script
# Look for these function names in your Legacy/M365-UserProvisioning-Enterprise-Fixed.ps1

<#
FUNCTIONS TO EXTRACT FOR AUTHENTICATION MODULE:
======================================================

Primary Functions:
- Connect-ToMicrosoftGraph
- Disconnect-FromMicrosoftGraph  
- Start-TenantDiscovery

Exchange Online Functions:
- Connect-ExchangeOnlineAtStartup
- Connect-ExchangeOnlineIfNeeded
- Test-ExchangeOnlineModule

Helper Functions:
- Any functions that support the above

Global Variables:
- $Global:IsConnected
- $Global:TenantInfo
- $Global:ExchangeOnlineConnected
- $Global:AvailableLicenses
- $Global:AvailableGroups
- $Global:AvailableUsers
- $Global:AvailableMailboxes
- $Global:DistributionLists
- $Global:MailEnabledSecurityGroups
- $Global:SharedMailboxes
- $Global:SharePointSites
- $Global:AcceptedDomains
#>

# Create this script to help extract the authentication module
# Save as: Scripts\Extract-AuthenticationModule.ps1

param(
    [string]$LegacyScript = "Legacy\M365-UserProvisioning-Enterprise-Fixed.ps1",
    [string]$OutputPath = "Modules\M365.Authentication"
)

Write-Host "🔍 Analyzing legacy script for authentication functions..." -ForegroundColor Cyan

if (-not (Test-Path $LegacyScript)) {
    Write-Error "Legacy script not found at: $LegacyScript"
    exit 1
}

# Read the legacy script
$LegacyContent = Get-Content -Path $LegacyScript -Raw

# Define the functions we want to extract
$FunctionsToExtract = @(
    'Connect-ToMicrosoftGraph',
    'Disconnect-FromMicrosoftGraph', 
    'Start-TenantDiscovery',
    'Connect-ExchangeOnlineAtStartup',
    'Connect-ExchangeOnlineIfNeeded', 
    'Test-ExchangeOnlineModule'
)

Write-Host "📋 Functions to extract:" -ForegroundColor Yellow
$FunctionsToExtract | ForEach-Object { Write-Host "  • $_" -ForegroundColor White }

# Check which functions exist in the legacy script
$FoundFunctions = @()
$MissingFunctions = @()

foreach ($FunctionName in $FunctionsToExtract) {
    if ($LegacyContent -match "function $FunctionName") {
        $FoundFunctions += $FunctionName
        Write-Host "✅ Found: $FunctionName" -ForegroundColor Green
    } else {
        $MissingFunctions += $FunctionName  
        Write-Host "❌ Missing: $FunctionName" -ForegroundColor Red
    }
}

Write-Host "`n📊 Summary:" -ForegroundColor Cyan
Write-Host "  Found: $($FoundFunctions.Count) functions" -ForegroundColor Green
Write-Host "  Missing: $($MissingFunctions.Count) functions" -ForegroundColor Red

if ($FoundFunctions.Count -eq 0) {
    Write-Error "No authentication functions found in legacy script. Please check function names."
    exit 1
}

Write-Host "`n🎯 Next steps:" -ForegroundColor Yellow
Write-Host "1. Manually copy the found functions to create the authentication module" -ForegroundColor White
Write-Host "2. Or run the extraction process automatically" -ForegroundColor White

$Choice = Read-Host "`nWould you like to proceed with automatic extraction? (y/N)"

if ($Choice -eq 'y' -or $Choice -eq 'Y') {
    Write-Host "`n🚀 Starting automatic extraction..." -ForegroundColor Cyan
    
    # Create output directory
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-Host "📁 Created directory: $OutputPath" -ForegroundColor Green
    }
    
    # TODO: Implement automatic extraction logic here
    # For now, we'll create a template
    
    $ModuleTemplate = @"
# M365.Authentication.psm1
# Microsoft Graph and Exchange Online authentication module for M365 User Provisioning Tool

<#
.SYNOPSIS
    Authentication module for M365 services
.DESCRIPTION  
    Handles Microsoft Graph and Exchange Online connections, disconnections, and tenant discovery
.NOTES
    Extracted from M365 User Provisioning Tool legacy script
    Version: 1.0.0
#>

# Module-scoped variables
`$Script:IsConnected = `$false
`$Script:TenantInfo = `$null
`$Script:ExchangeOnlineConnected = `$false
`$Script:TenantData = @{}

# TODO: Copy your authentication functions here from the legacy script
# Functions to copy:
$(foreach ($func in $FoundFunctions) { "# - $func" })

function Get-M365AuthenticationStatus {
    <#
    .SYNOPSIS
        Gets current authentication status for M365 services
    #>
    return @{
        GraphConnected = `$Script:IsConnected
        ExchangeOnlineConnected = `$Script:ExchangeOnlineConnected
        TenantInfo = `$Script:TenantInfo
        TenantData = `$Script:TenantData
    }
}

# Export public functions
Export-ModuleMember -Function @(
$($FoundFunctions | ForEach-Object { "    '$_'" })
    'Get-M365AuthenticationStatus'
)
"@

    # Write the module template
    $ModuleTemplate | Out-File -FilePath "$OutputPath\M365.Authentication.psm1" -Encoding UTF8
    Write-Host "📄 Created module template: $OutputPath\M365.Authentication.psm1" -ForegroundColor Green
    
    # Create module manifest
    $ManifestParams = @{
        Path = "$OutputPath\M365.Authentication.psd1"
        RootModule = 'M365.Authentication.psm1'
        ModuleVersion = '1.0.0'
        GUID = (New-Guid).Guid
        Author = 'Tom Mortiboys'
        CompanyName = 'M365 Project Delivery'
        Description = 'Microsoft Graph and Exchange Online authentication module for M365 User Provisioning Tool'
        PowerShellVersion = '7.0'
        RequiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Identity.DirectoryManagement')
        FunctionsToExport = $FoundFunctions + @('Get-M365AuthenticationStatus')
        CompatiblePSEditions = @('Core')
        Tags = @('M365', 'Authentication', 'MicrosoftGraph', 'ExchangeOnline')
    }
    
    New-ModuleManifest @ManifestParams
    Write-Host "📋 Created module manifest: $OutputPath\M365.Authentication.psd1" -ForegroundColor Green
    
    Write-Host "`n✅ Module template created successfully!" -ForegroundColor Green
    Write-Host "`n🔧 Manual steps required:" -ForegroundColor Yellow
    Write-Host "1. Open $OutputPath\M365.Authentication.psm1 in VSCode" -ForegroundColor White
    Write-Host "2. Copy the actual function code from your legacy script" -ForegroundColor White  
    Write-Host "3. Update any global variables to use module scope (`$Script: instead of `$Global:)" -ForegroundColor White
    Write-Host "4. Test the module: Import-Module .\$OutputPath -Force" -ForegroundColor White
}
else {
    Write-Host "Manual extraction steps will be provided next." -ForegroundColor Yellow
}