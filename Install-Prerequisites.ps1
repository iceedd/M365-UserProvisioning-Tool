#Requires -Version 7.0
[CmdletBinding()]
param([switch]$Force)

Write-Host "================================================" -ForegroundColor Magenta
Write-Host "M365 User Provisioning Tool - Prerequisites Setup" -ForegroundColor Magenta
Write-Host "================================================" -ForegroundColor Magenta
Write-Host ""

Write-Host "Checking PowerShell version..." -ForegroundColor Cyan
$currentVersion = $PSVersionTable.PSVersion
if ($currentVersion -lt [Version]"7.0") {
    Write-Host "   ERROR: PowerShell $currentVersion is too old" -ForegroundColor Red
    Write-Host "   This tool requires PowerShell 7.0 or later" -ForegroundColor Yellow
    Write-Host "   Download from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Yellow
    Write-Host "   Or install via: winget install Microsoft.PowerShell" -ForegroundColor Yellow
    exit 1
}
Write-Host "   OK: PowerShell $currentVersion" -ForegroundColor Green

Write-Host ""
Write-Host "Installing required modules..." -ForegroundColor Magenta

$modules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users', 
    'Microsoft.Graph.Users.Actions',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Sites',
    'ExchangeOnlineManagement'
)

$failedModules = @()

foreach ($module in $modules) {
    Write-Host "Processing $module..." -ForegroundColor Cyan
    
    $installed = Get-Module -ListAvailable -Name $module -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if ($installed -and -not $Force) {
        Write-Host "   Already installed: $($installed.Version)" -ForegroundColor Green
        continue
    }
    
    if ($Force) {
        Write-Host "   Force reinstall requested..." -ForegroundColor Yellow
    } else {
        Write-Host "   Not found, installing..." -ForegroundColor Yellow
    }
    
    $installResult = $null
    try {
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -Confirm:$false -SkipPublisherCheck -ErrorAction Stop
        Write-Host "   SUCCESS: $module installed" -ForegroundColor Green
    } catch {
        Write-Host "   FAILED: $($_.Exception.Message)" -ForegroundColor Red
        $failedModules += $module
    }
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Magenta

if ($failedModules.Count -eq 0) {
    Write-Host "SUCCESS: All modules installed!" -ForegroundColor Green
    Write-Host ""
    Write-Host "You can now run the M365 tool:" -ForegroundColor Cyan
    Write-Host "   .\M365-UserProvisioning-Enterprise.ps1" -ForegroundColor White
} else {
    Write-Host "WARNING: $($failedModules.Count) modules failed to install:" -ForegroundColor Yellow
    foreach ($failed in $failedModules) {
        Write-Host "   - $failed" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "You may need to:" -ForegroundColor Yellow
    Write-Host "   1. Run as Administrator" -ForegroundColor White
    Write-Host "   2. Check your internet connection" -ForegroundColor White
    Write-Host "   3. Install modules manually" -ForegroundColor White
}

Write-Host "================================================" -ForegroundColor Magenta