#Requires -Version 7.0

<#
.SYNOPSIS
    Quick test to verify Switch Tenant button appears in main script
#>

Write-Host "🧪 Testing Switch Tenant button in main script..." -ForegroundColor Yellow

# Force clear modules
Get-Module M365.* | Remove-Module -Force -ErrorAction SilentlyContinue

# Start the main application in test mode
try {
    Write-Host "🚀 Starting main application..." -ForegroundColor Cyan
    & .\M365-UserProvisioning-Enterprise.ps1 -TestMode
}
catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
}