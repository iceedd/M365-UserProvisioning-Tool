#Requires -Version 7.0
<#
.SYNOPSIS
    M365 User Provisioning Tool - Enterprise Edition 2025
    
.DESCRIPTION
    Main entry point for the modular M365 User Provisioning Tool
    
    Features:
    - Microsoft Graph and Exchange Online integration
    - Single user creation and bulk CSV import
    - Intelligent tenant discovery
    - Clean tabbed interface with pagination
    - Robust error handling and validation
    - Azure AD replication delay handling
    
.NOTES
    Version: 2.0.0 - Modular Architecture
    Author: Tom Mortiboys
    PowerShell: 7.0+ Required
    Dependencies: Microsoft Graph PowerShell SDK V2.28+, Exchange Online PowerShell
    
.EXAMPLE
    .\M365-UserProvisioning.ps1
#>

param(
    [switch]$NoGUI,
    [switch]$TestMode,
    [string]$LogPath = "Logs"
)

# Set strict mode and error preference
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Create logs directory if it doesn't exist
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

try {
    Write-Host "🚀 Starting M365 User Provisioning Tool - Enterprise Edition 2025" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "Version: 2.0.0 - Modular Architecture" -ForegroundColor Green
    Write-Host "Author: Tom Mortiboys" -ForegroundColor Green
    Write-Host "================================================================" -ForegroundColor Cyan
    
    # Import required modules
    Write-Host "📦 Loading modules..." -ForegroundColor Yellow
    
    # Check if modules exist
    $ModulesPath = Join-Path $PSScriptRoot "Modules"
    $RequiredModules = @(
        @{ Name = "M365.Authentication"; Path = Join-Path $ModulesPath "M365.Authentication" }
        @{ Name = "M365.UserManagement"; Path = Join-Path $ModulesPath "M365.UserManagement" }
        @{ Name = "M365.GUI"; Path = Join-Path $ModulesPath "M365.GUI" }
    )
    
    foreach ($Module in $RequiredModules) {
        if (-not (Test-Path $Module.Path)) {
            throw "Module directory not found: $($Module.Path)"
        }
        
        $ManifestPath = Join-Path $Module.Path "$($Module.Name).psd1"
        if (-not (Test-Path $ManifestPath)) {
            throw "Module manifest not found: $ManifestPath"
        }
        
        Write-Host "  • Importing $($Module.Name)..." -ForegroundColor White
        Import-Module $ManifestPath -Force -ErrorAction Stop
        Write-Host "    ✅ $($Module.Name) loaded successfully" -ForegroundColor Green
    }
    
    # Verify authentication module functions
    Write-Host "🔍 Verifying module functionality..." -ForegroundColor Yellow
    
    $AuthFunctions = Get-Command -Module M365.Authentication
    $UserMgmtFunctions = Get-Command -Module M365.UserManagement
    $GUIFunctions = Get-Command -Module M365.GUI
    
    Write-Host "  • Authentication functions: $($AuthFunctions.Count)" -ForegroundColor White
    Write-Host "  • User management functions: $($UserMgmtFunctions.Count)" -ForegroundColor White
    Write-Host "  • GUI functions: $($GUIFunctions.Count)" -ForegroundColor White
    
    if ($TestMode) {
        Write-Host "🧪 Test Mode - Showing available functions:" -ForegroundColor Magenta
        Write-Host "Authentication Module Functions:" -ForegroundColor Yellow
        $AuthFunctions | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor White }
        
        Write-Host "User Management Module Functions:" -ForegroundColor Yellow
        $UserMgmtFunctions | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor White }
        
        Write-Host "GUI Module Functions:" -ForegroundColor Yellow
        $GUIFunctions | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor White }
        
        return
    }
    
    # Import GUI module
    Write-Host "  • Importing M365.GUI..." -ForegroundColor White
    $GUIManifestPath = Join-Path $ModulesPath "M365.GUI\M365.GUI.psd1"
    if (-not (Test-Path $GUIManifestPath)) {
        throw "GUI module manifest not found: $GUIManifestPath"
    }
    Import-Module $GUIManifestPath -Force -ErrorAction Stop
    Write-Host "    ✅ M365.GUI loaded successfully" -ForegroundColor Green
    
    if ($NoGUI) {
        Write-Host "🔧 Console mode - Attempting to connect to Microsoft Graph..." -ForegroundColor Yellow
        $ConnectionResult = Connect-ToMicrosoftGraph
        
        if ($ConnectionResult.Success) {
            Write-Host "✅ Connected successfully!" -ForegroundColor Green
            Write-Host "Tenant: $($ConnectionResult.TenantId)" -ForegroundColor Cyan
            Write-Host "Account: $($ConnectionResult.Account)" -ForegroundColor Cyan
            
            # Show tenant data summary
            $TenantData = Get-M365TenantData
            Write-Host "`n📊 Tenant Summary:" -ForegroundColor Yellow
            Write-Host "  • Users: $($TenantData.AvailableUsers.Count)" -ForegroundColor White
            Write-Host "  • Groups: $($TenantData.AvailableGroups.Count)" -ForegroundColor White
            Write-Host "  • Licenses: $($TenantData.AvailableLicenses.Count)" -ForegroundColor White
            Write-Host "  • Domains: $($TenantData.AcceptedDomains.Count)" -ForegroundColor White
        }
        else {
            Write-Host "❌ Connection failed: $($ConnectionResult.Message)" -ForegroundColor Red
            exit 1
        }
    }
    else {
        # Start GUI interface
        Write-Host "🖥️ Starting GUI interface..." -ForegroundColor Yellow
        
        # Verify GUI functions are available
        $GUIFunctions = Get-Command -Module M365.GUI
        Write-Host "  • GUI functions available: $($GUIFunctions.Count)" -ForegroundColor White
        
        # Launch the GUI
        Write-Host "🚀 Launching M365 User Provisioning Tool GUI..." -ForegroundColor Green
        Start-M365ProvisioningTool
    }
    
    Write-Host "`n✅ M365 User Provisioning Tool ready!" -ForegroundColor Green
    
}
catch {
    Write-Host "❌ Critical Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    
    # Log error
    $ErrorLog = Join-Path $LogPath "error_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    "Error: $($_.Exception.Message)" | Out-File $ErrorLog -Encoding UTF8
    "Stack Trace: $($_.ScriptStackTrace)" | Add-Content $ErrorLog -Encoding UTF8
    
    Write-Host "Error details logged to: $ErrorLog" -ForegroundColor Yellow
    exit 1
}
finally {
    Write-Host "Session ended at $(Get-Date)" -ForegroundColor Gray
}