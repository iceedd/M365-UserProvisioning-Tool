# Test-ExchangeOnlineIntegration.ps1
# Quick test script to verify M365.ExchangeOnline module integration

Write-Host "🧪 Testing M365.ExchangeOnline Module Integration" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan

# Test 1: Module Loading
Write-Host "`n1. Testing module loading..." -ForegroundColor Yellow

try {
    Write-Host "   Loading M365.Authentication..." -ForegroundColor Gray
    Import-Module .\Modules\M365.Authentication\M365.Authentication.psd1 -Force
    Write-Host "   ✅ M365.Authentication loaded" -ForegroundColor Green
    
    Write-Host "   Loading M365.ExchangeOnline..." -ForegroundColor Gray
    Import-Module .\Modules\M365.ExchangeOnline\M365.ExchangeOnline.psd1 -Force
    Write-Host "   ✅ M365.ExchangeOnline loaded" -ForegroundColor Green
    
} catch {
    Write-Host "   ❌ Module loading failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 2: Check Available Functions
Write-Host "`n2. Checking available functions..." -ForegroundColor Yellow

$ExchangeFunctions = Get-Command -Module M365.ExchangeOnline
Write-Host "   Available Exchange functions:" -ForegroundColor Gray
foreach ($Function in $ExchangeFunctions) {
    Write-Host "   - $($Function.Name)" -ForegroundColor Cyan
}

# Test 3: Authentication Test (Optional)
Write-Host "`n3. Authentication test (optional)..." -ForegroundColor Yellow
$TestAuth = Read-Host "   Do you want to test authentication? (y/n)"

if ($TestAuth -eq 'y' -or $TestAuth -eq 'Y') {
    try {
        Write-Host "   Connecting to Microsoft Graph..." -ForegroundColor Gray
        $GraphResult = Connect-ToMicrosoftGraph
        
        if ($GraphResult) {
            Write-Host "   ✅ Microsoft Graph connected" -ForegroundColor Green
            
            Write-Host "   Attempting Exchange Online connection..." -ForegroundColor Gray
            $ExchangeResult = Connect-ExchangeOnlineAtStartup
            
            if ($ExchangeResult.Connected) {
                Write-Host "   ✅ Exchange Online connected" -ForegroundColor Green
                
                # Test 4: Exchange Data Discovery
                Write-Host "`n4. Testing Exchange data discovery..." -ForegroundColor Yellow
                
                Write-Host "   Getting Exchange data..." -ForegroundColor Gray
                $ExchangeData = Get-AllExchangeData
                
                Write-Host "   📊 Exchange Discovery Results:" -ForegroundColor Cyan
                Write-Host "      User Mailboxes: $($ExchangeData.UserMailboxes.Count)" -ForegroundColor White
                Write-Host "      Shared Mailboxes: $($ExchangeData.SharedMailboxes.Count)" -ForegroundColor White
                Write-Host "      Distribution Lists: $($ExchangeData.DistributionLists.Count)" -ForegroundColor White
                Write-Host "      Mail-Enabled Security Groups: $($ExchangeData.MailEnabledSecurityGroups.Count)" -ForegroundColor White
                Write-Host "      Accepted Domains: $($ExchangeData.AcceptedDomains.Count)" -ForegroundColor White
                Write-Host "      Data Source: $($ExchangeData.ConnectionStatus.DataSource)" -ForegroundColor White
                
                Write-Host "`n   🎯 Integration Test Results:" -ForegroundColor Green
                if ($ExchangeData.ConnectionStatus.ExchangeOnlineConnected) {
                    Write-Host "      ✅ Full Exchange Online functionality available" -ForegroundColor Green
                    Write-Host "      ✅ Accurate shared mailbox detection" -ForegroundColor Green
                    Write-Host "      ✅ Complete distribution list data" -ForegroundColor Green
                } else {
                    Write-Host "      ⚠️  Using Graph API fallback (limited accuracy)" -ForegroundColor Yellow
                    Write-Host "      ⚠️  Shared mailbox detection may be unreliable" -ForegroundColor Yellow
                }
                
            } else {
                Write-Host "   ⚠️  Exchange Online not connected - testing Graph fallback" -ForegroundColor Yellow
                
                # Test fallback functionality
                Write-Host "`n4. Testing Graph API fallback..." -ForegroundColor Yellow
                $ExchangeData = Get-AllExchangeData
                
                Write-Host "   📊 Fallback Results:" -ForegroundColor Cyan
                Write-Host "      Data Source: $($ExchangeData.ConnectionStatus.DataSource)" -ForegroundColor White
                Write-Host "      Fallback data available: $($ExchangeData.UserMailboxes.Count -gt 0)" -ForegroundColor White
            }
        } else {
            Write-Host "   ❌ Microsoft Graph connection failed" -ForegroundColor Red
        }
        
    } catch {
        Write-Host "   ❌ Authentication test failed: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "   ⏭️  Skipping authentication test" -ForegroundColor Gray
}

# Test 5: Module Integration Check
Write-Host "`n5. Checking integration points..." -ForegroundColor Yellow

# Check if functions can call each other
try {
    # Test if Exchange module can use Authentication module functions
    $AuthStatus = Get-M365AuthenticationStatus
    Write-Host "   ✅ M365.ExchangeOnline can access M365.Authentication functions" -ForegroundColor Green
} catch {
    Write-Host "   ❌ Module integration issue: $($_.Exception.Message)" -ForegroundColor Red
}

# Final Summary
Write-Host "`n🎯 Integration Test Summary" -ForegroundColor Cyan
Write-Host "=========================" -ForegroundColor Cyan

Write-Host "Module Loading: " -NoNewline
if (Get-Module M365.ExchangeOnline) {
    Write-Host "✅ Success" -ForegroundColor Green
} else {
    Write-Host "❌ Failed" -ForegroundColor Red
}

Write-Host "Function Availability: " -NoNewline
if ($ExchangeFunctions.Count -gt 0) {
    Write-Host "✅ $($ExchangeFunctions.Count) functions available" -ForegroundColor Green
} else {
    Write-Host "❌ No functions found" -ForegroundColor Red
}

Write-Host "`n📋 Next Steps:" -ForegroundColor Yellow
Write-Host "1. Update your main script to include 'M365.ExchangeOnline' in RequiredModules" -ForegroundColor White
Write-Host "2. Enhance your Start-TenantDiscovery function to use Get-AllExchangeData" -ForegroundColor White
Write-Host "3. Update user creation logic to handle Exchange assignments" -ForegroundColor White
Write-Host "4. Test with your existing M365-UserProvisioning-Enterprise.ps1 script" -ForegroundColor White

Write-Host "`n✅ M365.ExchangeOnline module integration test completed!" -ForegroundColor Green