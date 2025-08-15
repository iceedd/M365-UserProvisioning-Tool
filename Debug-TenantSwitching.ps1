#Requires -Version 7.0

<#
.SYNOPSIS
    Debug script to monitor tenant switching and data clearing
#>

Write-Host "🔍 Tenant Switching Debug Monitor" -ForegroundColor Yellow
Write-Host "This script helps monitor what happens during tenant switching" -ForegroundColor Gray
Write-Host ""

# Function to display current tenant data state
function Show-TenantDataState {
    param([string]$Label)
    
    Write-Host "📊 $Label" -ForegroundColor Cyan
    Write-Host "   AcceptedDomains: $($Global:AcceptedDomains.Count) items" -ForegroundColor Gray
    Write-Host "   AvailableUsers: $($Global:AvailableUsers.Count) items" -ForegroundColor Gray
    Write-Host "   AvailableGroups: $($Global:AvailableGroups.Count) items" -ForegroundColor Gray
    Write-Host "   SharedMailboxes: $($Global:SharedMailboxes.Count) items" -ForegroundColor Gray
    Write-Host "   DistributionLists: $($Global:DistributionLists.Count) items" -ForegroundColor Gray
    Write-Host "   IsConnected: $Global:IsConnected" -ForegroundColor Gray
    
    if ($Global:AcceptedDomains.Count -gt 0) {
        Write-Host "   Sample domains: $($Global:AcceptedDomains[0..2] | ForEach-Object { $_.Id } | Select-Object -First 3)" -ForegroundColor DarkGray
    }
    if ($Global:SharedMailboxes.Count -gt 0) {
        Write-Host "   Sample shared mailboxes: $($Global:SharedMailboxes[0..2] | ForEach-Object { $_.DisplayName } | Select-Object -First 3)" -ForegroundColor DarkGray
    }
    if ($Global:DistributionLists.Count -gt 0) {
        Write-Host "   Sample distribution lists: $($Global:DistributionLists[0..2] | ForEach-Object { $_.DisplayName } | Select-Object -First 3)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# Instructions
Write-Host "📋 INSTRUCTIONS:" -ForegroundColor Yellow
Write-Host "1. Run your M365 User Provisioning Tool in another window" -ForegroundColor Gray
Write-Host "2. Connect to the first tenant and wait for data discovery" -ForegroundColor Gray
Write-Host "3. Come back here and press ENTER to see tenant data state" -ForegroundColor Gray
Write-Host "4. Go back and click 'Switch Tenant' button" -ForegroundColor Gray
Write-Host "5. Come back here and press ENTER again to see if data was cleared" -ForegroundColor Gray
Write-Host "6. Connect to a different tenant" -ForegroundColor Gray
Write-Host "7. Come back here and press ENTER to see new tenant data" -ForegroundColor Gray
Write-Host ""

$step = 1
while ($true) {
    Write-Host "Step $step - Press ENTER when ready to check tenant data state (or 'q' to quit):" -ForegroundColor White -NoNewline
    $input = Read-Host
    
    if ($input -eq 'q') {
        Write-Host "👋 Exiting debug monitor" -ForegroundColor Green
        break
    }
    
    try {
        Show-TenantDataState "Current Tenant Data State"
        
        # Give specific feedback based on what we see
        if ($Global:IsConnected -eq $true -and $Global:AcceptedDomains.Count -gt 0) {
            Write-Host "✅ Connected to tenant with data" -ForegroundColor Green
        } elseif ($Global:IsConnected -eq $false -and $Global:AcceptedDomains.Count -eq 0) {
            Write-Host "✅ Disconnected and data cleared" -ForegroundColor Green
        } elseif ($Global:IsConnected -eq $false -and $Global:AcceptedDomains.Count -gt 0) {
            Write-Host "⚠️ WARNING: Disconnected but data still present!" -ForegroundColor Yellow
        } elseif ($Global:IsConnected -eq $true -and $Global:AcceptedDomains.Count -eq 0) {
            Write-Host "⚠️ WARNING: Connected but no data loaded" -ForegroundColor Yellow
        } else {
            Write-Host "ℹ️ Initial state or data loading..." -ForegroundColor Blue
        }
    }
    catch {
        Write-Host "❌ Error checking tenant data: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Make sure the M365 tool is running in another window" -ForegroundColor Gray
    }
    
    Write-Host ""
    $step++
}