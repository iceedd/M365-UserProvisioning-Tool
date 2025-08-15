#Requires -Version 7.0

<#
.SYNOPSIS
    Test script to verify tenant data clearing functionality
#>

Write-Host "🧪 Testing Tenant Data Clearing Functionality..." -ForegroundColor Yellow
Write-Host ""

# Test the Clear-TenantData function by loading the main script functions
try {
    # Source the main script to get access to functions
    Write-Host "📁 Loading main script functions..." -ForegroundColor Cyan
    
    # Load just the functions we need for testing
    $MainScriptContent = Get-Content ".\M365-UserProvisioning-Enterprise.ps1" -Raw
    
    # Extract and execute the Clear-TenantData function
    if ($MainScriptContent -match '(?s)function Clear-TenantData.*?^}') {
        $ClearTenantDataFunction = $matches[0]
        Write-Host "✅ Found Clear-TenantData function" -ForegroundColor Green
        
        # Create mock script variables to test clearing
        Write-Host "🏗️ Creating mock tenant data..." -ForegroundColor Cyan
        
        $Global:AcceptedDomains = @("domain1.com", "domain2.com")
        $Global:AvailableUsers = @("user1", "user2", "user3") 
        $Global:AvailableGroups = @("Group1", "Group2", "Group3")
        $Global:AvailableLicenses = @("License1", "License2")
        $Global:SharePointSites = @("Site1", "Site2")
        $Global:SharedMailboxes = @("Mailbox1", "Mailbox2")
        $Global:DistributionLists = @("DL1", "DL2")
        
        Write-Host "   Created mock data:" -ForegroundColor Gray
        Write-Host "     • AcceptedDomains: $($Global:AcceptedDomains.Count) items" -ForegroundColor Gray
        Write-Host "     • AvailableUsers: $($Global:AvailableUsers.Count) items" -ForegroundColor Gray
        Write-Host "     • AvailableGroups: $($Global:AvailableGroups.Count) items" -ForegroundColor Gray
        Write-Host "     • AvailableLicenses: $($Global:AvailableLicenses.Count) items" -ForegroundColor Gray
        Write-Host "     • SharePointSites: $($Global:SharePointSites.Count) items" -ForegroundColor Gray
        Write-Host "     • SharedMailboxes: $($Global:SharedMailboxes.Count) items" -ForegroundColor Gray
        Write-Host "     • DistributionLists: $($Global:DistributionLists.Count) items" -ForegroundColor Gray
        
        Write-Host ""
        Write-Host "🧹 Running Clear-TenantData function..." -ForegroundColor Yellow
        
        # Execute the function (this will show its own output)
        Invoke-Expression $ClearTenantDataFunction
        Clear-TenantData
        
        Write-Host ""
        Write-Host "🔍 Verifying data was cleared..." -ForegroundColor Cyan
        
        $AllCleared = $true
        
        if ($Global:AcceptedDomains.Count -ne 0) {
            Write-Host "❌ AcceptedDomains not cleared: $($Global:AcceptedDomains.Count) items remain" -ForegroundColor Red
            $AllCleared = $false
        } else {
            Write-Host "✅ AcceptedDomains cleared" -ForegroundColor Green
        }
        
        if ($Global:AvailableUsers.Count -ne 0) {
            Write-Host "❌ AvailableUsers not cleared: $($Global:AvailableUsers.Count) items remain" -ForegroundColor Red
            $AllCleared = $false
        } else {
            Write-Host "✅ AvailableUsers cleared" -ForegroundColor Green
        }
        
        if ($Global:AvailableGroups.Count -ne 0) {
            Write-Host "❌ AvailableGroups not cleared: $($Global:AvailableGroups.Count) items remain" -ForegroundColor Red
            $AllCleared = $false
        } else {
            Write-Host "✅ AvailableGroups cleared" -ForegroundColor Green
        }
        
        if ($Global:AvailableLicenses.Count -ne 0) {
            Write-Host "❌ AvailableLicenses not cleared: $($Global:AvailableLicenses.Count) items remain" -ForegroundColor Red
            $AllCleared = $false
        } else {
            Write-Host "✅ AvailableLicenses cleared" -ForegroundColor Green
        }
        
        if ($Global:SharePointSites.Count -ne 0) {
            Write-Host "❌ SharePointSites not cleared: $($Global:SharePointSites.Count) items remain" -ForegroundColor Red
            $AllCleared = $false
        } else {
            Write-Host "✅ SharePointSites cleared" -ForegroundColor Green
        }
        
        if ($Global:SharedMailboxes.Count -ne 0) {
            Write-Host "❌ SharedMailboxes not cleared: $($Global:SharedMailboxes.Count) items remain" -ForegroundColor Red
            $AllCleared = $false
        } else {
            Write-Host "✅ SharedMailboxes cleared" -ForegroundColor Green
        }
        
        if ($Global:DistributionLists.Count -ne 0) {
            Write-Host "❌ DistributionLists not cleared: $($Global:DistributionLists.Count) items remain" -ForegroundColor Red
            $AllCleared = $false
        } else {
            Write-Host "✅ DistributionLists cleared" -ForegroundColor Green
        }
        
        Write-Host ""
        if ($AllCleared) {
            Write-Host "🎉 SUCCESS: All tenant data cleared successfully!" -ForegroundColor Green
        } else {
            Write-Host "❌ FAILURE: Some tenant data was not cleared properly" -ForegroundColor Red
        }
    } else {
        Write-Host "❌ Could not find Clear-TenantData function in main script" -ForegroundColor Red
    }
}
catch {
    Write-Host "❌ Error during testing: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "🏁 Test completed" -ForegroundColor Green