#Requires -Version 7.0

<#
.SYNOPSIS
    Force reload all modules and run the main application
.DESCRIPTION
    This script completely clears PowerShell module cache and forces fresh loading
#>

Write-Host "🔄 Forcing complete module reload..." -ForegroundColor Yellow

# Remove all M365 modules from memory
$ModulesToRemove = @("M365.GUI", "M365.Authentication", "M365.UserManagement", "M365.ExchangeOnline", "M365.Utilities")

foreach ($ModuleName in $ModulesToRemove) {
    $Module = Get-Module $ModuleName -ErrorAction SilentlyContinue
    if ($Module) {
        Write-Host "  Removing $ModuleName from memory..." -ForegroundColor Cyan
        Remove-Module $ModuleName -Force -ErrorAction SilentlyContinue
    }
}

# Clear any cached imports
Write-Host "  Clearing import cache..." -ForegroundColor Cyan
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "✅ Module cache cleared" -ForegroundColor Green
Write-Host ""
Write-Host "🚀 Starting M365 User Provisioning Tool with fresh modules..." -ForegroundColor Green

# Run the main script with fresh module loading
& .\M365-UserProvisioning-Enterprise.ps1