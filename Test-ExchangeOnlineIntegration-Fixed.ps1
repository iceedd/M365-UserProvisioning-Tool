# Test-ExchangeOnlineIntegration-Fixed.ps1
# Corrected test script that adapts to your actual directory structure

Write-Host "🧪 Testing M365.ExchangeOnline Module Integration (Fixed)" -ForegroundColor Cyan
Write-Host "=======================================================" -ForegroundColor Cyan

# Get current directory for proper pathing
$ScriptDir = Get-Location
Write-Host "`n📍 Working from: $ScriptDir" -ForegroundColor Gray

# Test 1: Directory Structure Check
Write-Host "`n1. Checking directory structure..." -ForegroundColor Yellow

$HasModulesDir = Test-Path ".\Modules"
$HasExchangeModule = Test-Path ".\Modules\M365.ExchangeOnline"

if ($HasModulesDir) {
    Write-Host "   ✅ Modules directory found" -ForegroundColor Green
    
    if ($HasExchangeModule) {
        Write-Host "   ✅ M365.ExchangeOnline directory found" -ForegroundColor Green
    } else {
        Write-Host "   ❌ M365.ExchangeOnline directory missing" -ForegroundColor Red
        Write-Host "   💡 You need to create: .\Modules\M365.ExchangeOnline\" -ForegroundColor Yellow
        Write-Host "   💡 And add the .psm1 and .psd1 files there" -ForegroundColor Yellow
        exit 1
    }
} else {
    Write-Host "   ❌ Modules directory not found" -ForegroundColor Red
    Write-Host "   💡 You might have a single-script architecture" -ForegroundColor Yellow
    Write-Host "   💡 Run the diagnostic script first: .\Diagnostic-DirectoryStructure.ps1" -ForegroundColor Yellow
    exit 1
}

# Test 2: Module File Check
Write-Host "`n2. Checking module files..." -ForegroundColor Yellow

$ExchangePsm = ".\Modules\M365.ExchangeOnline\M365.ExchangeOnline.psm1"
$ExchangePsd = ".\Modules\M365.ExchangeOnline\M365.ExchangeOnline.psd1"

if (Test-Path $ExchangePsm) {
    Write-Host "   ✅ M365.ExchangeOnline.psm1 found" -ForegroundColor Green
    $PsmSize = (Get-Item $ExchangePsm).Length
    Write-Host "      📊 Size: $PsmSize bytes" -ForegroundColor Gray
} else {
    Write-Host "   ❌ M365.ExchangeOnline.psm1 missing" -ForegroundColor Red
    exit 1
}

if (Test-Path $ExchangePsd) {
    Write-Host "   ✅ M365.ExchangeOnline.psd1 found" -ForegroundColor Green
    
    # Test manifest validity
    try {
        $Manifest = Test-ModuleManifest $ExchangePsd -ErrorAction Stop
        Write-Host "      ✅ Module manifest is valid" -ForegroundColor Green
        Write-Host "      📋 Version: $($Manifest.Version)" -ForegroundColor Gray
        Write-Host "      📋 Author: $($Manifest.Author)" -ForegroundColor Gray
    } catch {
        Write-Host "      ❌ Module manifest error: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "   ❌ M365.ExchangeOnline.psd1 missing" -ForegroundColor Red
    exit 1
}

# Test 3: Module Loading
Write-Host "`n3. Testing module loading..." -ForegroundColor Yellow

try {
    # Clear any existing modules first
    Remove-Module M365.ExchangeOnline -Force -ErrorAction SilentlyContinue
    
    Write-Host "   Loading M365.ExchangeOnline..." -ForegroundColor Gray
    Import-Module $ExchangePsd -Force -ErrorAction Stop
    Write-Host "   ✅ M365.ExchangeOnline loaded successfully" -ForegroundColor Green
    
} catch {
    Write-Host "   ❌ Module loading failed: $($_.Exception.Message)" -ForegroundColor Red
    
    # Provide specific troubleshooting based on error
    if ($_.Exception.Message -like "*dependency*" -or $_.Exception.Message -like "*required*") {
        Write-Host "   💡 This looks like a dependency issue" -ForegroundColor Yellow
        Write-Host "   💡 Make sure you have ExchangeOnlineManagement module installed:" -ForegroundColor Yellow
        Write-Host "      Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor White
    }
    
    exit 1
}

# Test 4: Check Available Functions
Write-Host "`n4. Checking available functions..." -ForegroundColor Yellow

$ExchangeFunctions = Get-Command -Module M365.ExchangeOnline -ErrorAction SilentlyContinue

if ($ExchangeFunctions.Count -gt 0) {
    Write-Host "   ✅ Found $($ExchangeFunctions.Count) Exchange functions:" -ForegroundColor Green
    foreach ($Function in $ExchangeFunctions) {
        Write-Host "      - $($Function.Name)" -ForegroundColor Cyan
    }
} else {
    Write-Host "   ❌ No functions found in M365.ExchangeOnline module" -ForegroundColor Red
    Write-Host "   💡 Check that the .psm1 file contains the Export-ModuleMember statement" -ForegroundColor Yellow
    exit 1
}

# Test 5: Check Dependencies
Write-Host "`n5. Checking dependencies..." -ForegroundColor Yellow

$RequiredModules = @('ExchangeOnlineManagement', 'Microsoft.Graph.Users', 'Microsoft.Graph.Groups')

foreach ($ModuleName in $RequiredModules) {
    $Module = Get-Module -Name $ModuleName -ListAvailable
    if ($Module) {
        Write-Host "   ✅ $ModuleName available (Version: $($Module.Version | Select-Object -First 1))" -ForegroundColor Green
    } else {
        Write-Host "   ⚠️  $ModuleName not installed" -ForegroundColor Yellow
        Write-Host "      Install with: Install-Module $ModuleName -Scope CurrentUser" -ForegroundColor Gray
    }
}

# Test 6: Test Authentication Module (if it exists)
Write-Host "`n6. Testing M365.Authentication integration..." -ForegroundColor Yellow

$AuthModulePath = ".\Modules\M365.Authentication\M365.Authentication.psd1"
if (Test-Path $AuthModulePath) {
    try {
        Import-Module $AuthModulePath -Force -ErrorAction Stop
        Write-Host "   ✅ M365.Authentication module loaded" -ForegroundColor Green
        
        # Test if functions are available
        $AuthFunctions = Get-Command -Module M365.Authentication -ErrorAction SilentlyContinue
        if ($AuthFunctions) {
            Write-Host "   ✅ Authentication functions available: $($AuthFunctions.Count)" -ForegroundColor Green
        } else {
            Write-Host "   ⚠️  Authentication module loaded but no functions found" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "   ⚠️  M365.Authentication module has issues: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "   ⚠️  M365.Authentication module not found" -ForegroundColor Yellow
    Write-Host "   💡 You may need to create this module or use direct Graph API calls" -ForegroundColor Yellow
}

# Test 7: Optional Live Test
Write-Host "`n7. Optional live authentication test..." -ForegroundColor Yellow
$TestLive = Read-Host "   Do you want to test with live authentication? (y/n)"

if ($TestLive -eq 'y' -or $TestLive -eq 'Y') {
    Write-Host "   Starting live authentication test..." -ForegroundColor Gray
    
    # Check if we have authentication functions available
    $ConnectGraphCmd = Get-Command Connect-ToMicrosoftGraph -ErrorAction SilentlyContinue
    $ConnectExchangeCmd = Get-Command Connect-ExchangeOnlineAtStartup -ErrorAction SilentlyContinue
    
    if ($ConnectGraphCmd) {
        try {
            Write-Host "   Connecting to Microsoft Graph..." -ForegroundColor Gray
            $GraphResult = Connect-ToMicrosoftGraph
            if ($GraphResult) {
                Write-Host "   ✅ Microsoft Graph connected" -ForegroundColor Green
            }
        } catch {
            Write-Host "   ❌ Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "   ⚠️  Connect-ToMicrosoftGraph function not available" -ForegroundColor Yellow
        Write-Host "   💡 Trying direct Graph connection..." -ForegroundColor Gray
        
        try {
            Connect-MgGraph -Scopes "User.Read.All","Group.Read.All" -NoWelcome
            Write-Host "   ✅ Direct Graph connection successful" -ForegroundColor Green
        } catch {
            Write-Host "   ❌ Direct Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    if ($ConnectExchangeCmd) {
        try {
            Write-Host "   Connecting to Exchange Online..." -ForegroundColor Gray
            $ExchangeResult = Connect-ExchangeOnlineAtStartup
            if ($ExchangeResult.Connected) {
                Write-Host "   ✅ Exchange Online connected" -ForegroundColor Green
                
                # Test Exchange data retrieval
                Write-Host "   Testing Exchange data retrieval..." -ForegroundColor Gray
                $ExchangeData = Get-AllExchangeData
                
                Write-Host "   📊 Exchange Test Results:" -ForegroundColor Cyan
                Write-Host "      User Mailboxes: $($ExchangeData.UserMailboxes.Count)" -ForegroundColor White
                Write-Host "      Shared Mailboxes: $($ExchangeData.SharedMailboxes.Count)" -ForegroundColor White
                Write-Host "      Distribution Lists: $($ExchangeData.DistributionLists.Count)" -ForegroundColor White
                Write-Host "      Data Source: $($ExchangeData.ConnectionStatus.DataSource)" -ForegroundColor White
                
            } else {
                Write-Host "   ⚠️  Exchange Online connection failed - testing fallback" -ForegroundColor Yellow
                $ExchangeData = Get-AllExchangeData
                Write-Host "   📊 Fallback Results: $($ExchangeData.ConnectionStatus.DataSource)" -ForegroundColor Cyan
            }
        } catch {
            Write-Host "   ❌ Exchange test failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "   ⚠️  Exchange authentication function not available" -ForegroundColor Yellow
    }
} else {
    Write-Host "   ⏭️  Skipping live authentication test" -ForegroundColor Gray
}

# Final Summary
Write-Host "`n🎯 INTEGRATION TEST SUMMARY" -ForegroundColor Cyan
Write-Host "===========================" -ForegroundColor Cyan

$ModuleLoaded = Get-Module M365.ExchangeOnline -ErrorAction SilentlyContinue
$FunctionsAvailable = (Get-Command -Module M365.ExchangeOnline -ErrorAction SilentlyContinue).Count

Write-Host "Module Structure: " -NoNewline
if ($HasExchangeModule) { Write-Host "✅ Good" -ForegroundColor Green } else { Write-Host "❌ Issues" -ForegroundColor Red }

Write-Host "Module Loading: " -NoNewline
if ($ModuleLoaded) { Write-Host "✅ Success" -ForegroundColor Green } else { Write-Host "❌ Failed" -ForegroundColor Red }

Write-Host "Functions Available: " -NoNewline
if ($FunctionsAvailable -gt 0) { Write-Host "✅ $FunctionsAvailable functions" -ForegroundColor Green } else { Write-Host "❌ None found" -ForegroundColor Red }

if ($ModuleLoaded -and $FunctionsAvailable -gt 0) {
    Write-Host "`n🎉 SUCCESS! M365.ExchangeOnline module is ready for integration!" -ForegroundColor Green
    
    Write-Host "`n📋 NEXT STEPS:" -ForegroundColor Yellow
    Write-Host "1. Add 'M365.ExchangeOnline' to your main script's RequiredModules" -ForegroundColor White
    Write-Host "2. Update your tenant discovery to use Get-AllExchangeData" -ForegroundColor White
    Write-Host "3. Add Exchange assignments to your user creation process" -ForegroundColor White
    Write-Host "4. Test the integration with your full application" -ForegroundColor White
    
} else {
    Write-Host "`n⚠️  Integration not ready - please fix the issues above first" -ForegroundColor Yellow
}

Write-Host "`n✅ Test completed!" -ForegroundColor Green