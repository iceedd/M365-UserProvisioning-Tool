# Test-ModuleLoading-Fixed.ps1 - CORRECTED VERSION
# Run this from ROOT directory: .\Test-ModuleLoading-Fixed.ps1

Write-Host "🧪 M365 User Provisioning Tool - Module Validation (FIXED)" -ForegroundColor Cyan
Write-Host "===========================================================" -ForegroundColor Cyan

$CurrentLocation = Get-Location
Write-Host "📍 Running from: $CurrentLocation" -ForegroundColor Gray

# Test 0: Verify directory structure
Write-Host "`n0️⃣ Verifying Directory Structure..." -ForegroundColor Yellow
$ExpectedPaths = @(
    "Modules\M365.Authentication\M365.Authentication.psm1",
    "Modules\M365.Authentication\M365.Authentication.psd1",
    "Modules\M365.UserManagement\M365.UserManagement.psm1", 
    "Modules\M365.UserManagement\M365.UserManagement.psd1"
)

$AllPathsExist = $true
foreach ($Path in $ExpectedPaths) {
    if (Test-Path $Path) {
        Write-Host "   ✅ Found: $Path" -ForegroundColor Green
    } else {
        Write-Host "   ❌ Missing: $Path" -ForegroundColor Red
        $AllPathsExist = $false
    }
}

if (-not $AllPathsExist) {
    Write-Host "   🚨 CRITICAL: Module files are missing! Create them first." -ForegroundColor Red
    Write-Host "   📝 Follow the artifacts provided to create missing files." -ForegroundColor Yellow
    exit 1
}

# Test 1: Authentication Module
Write-Host "`n1️⃣ Testing Authentication Module..." -ForegroundColor Yellow
try {
    # Use absolute paths to avoid confusion
    $AuthModulePath = Join-Path $PWD "Modules\M365.Authentication\M365.Authentication.psd1"
    Write-Host "   📁 Loading from: $AuthModulePath" -ForegroundColor Gray
    
    Import-Module $AuthModulePath -Force -Scope Global
    $AuthFunctions = Get-Command -Module M365.Authentication
    
    Write-Host "   ✅ Authentication module loaded successfully" -ForegroundColor Green
    Write-Host "   📊 Functions exported: $($AuthFunctions.Count)" -ForegroundColor Cyan
    
    if ($AuthFunctions.Count -eq 9) {
        Write-Host "   🎯 Correct function count (9/9)" -ForegroundColor Green
    } else {
        Write-Host "   ⚠️ Expected 9 functions, got $($AuthFunctions.Count)" -ForegroundColor Yellow
    }
    
    # List functions
    Write-Host "   📋 Available functions:" -ForegroundColor White
    $AuthFunctions | ForEach-Object { Write-Host "      • $($_.Name)" -ForegroundColor Gray }
    
    # Test a basic function
    Write-Host "   🔍 Testing Get-M365AuthenticationStatus..." -ForegroundColor White
    $Status = Get-M365AuthenticationStatus
    if ($Status) {
        Write-Host "      ✅ Status object returned successfully" -ForegroundColor Green
        Write-Host "      📊 Graph Connected: $($Status.GraphConnected)" -ForegroundColor Cyan
        Write-Host "      📊 Exchange Connected: $($Status.ExchangeOnlineConnected)" -ForegroundColor Cyan
    } else {
        Write-Host "      ❌ Status function returned null" -ForegroundColor Red
    }
    
} catch {
    Write-Host "   ❌ Authentication module failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   📍 Error location: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Gray
}

# Test 2: User Management Module
Write-Host "`n2️⃣ Testing User Management Module..." -ForegroundColor Yellow
try {
    $UserMgmtModulePath = Join-Path $PWD "Modules\M365.UserManagement\M365.UserManagement.psd1"
    Write-Host "   📁 Loading from: $UserMgmtModulePath" -ForegroundColor Gray
    
    Import-Module $UserMgmtModulePath -Force -Scope Global
    $UserMgmtFunctions = Get-Command -Module M365.UserManagement
    
    Write-Host "   ✅ User Management module loaded successfully" -ForegroundColor Green
    Write-Host "   📊 Functions exported: $($UserMgmtFunctions.Count)" -ForegroundColor Cyan
    
    # List functions
    Write-Host "   📋 Available functions:" -ForegroundColor White
    $UserMgmtFunctions | ForEach-Object { Write-Host "      • $($_.Name)" -ForegroundColor Gray }
    
    # Test password generation
    Write-Host "   🔍 Testing New-SecurePassword..." -ForegroundColor White
    $TestPassword = New-SecurePassword
    if ($TestPassword -and $TestPassword.Length -eq 16) {
        Write-Host "      ✅ Generated password (length: $($TestPassword.Length)): $TestPassword" -ForegroundColor Green
    } else {
        Write-Host "      ❌ Password generation failed" -ForegroundColor Red
    }
    
    # Test user creation function (dry run)
    Write-Host "   🔍 Testing New-M365User (dry run)..." -ForegroundColor White
    $TestUser = New-M365User -DisplayName "Test User" -UserPrincipalName "test@domain.com" -Password "TestPass123!"
    if ($TestUser -and $TestUser.DisplayName -eq "Test User") {
        Write-Host "      ✅ User creation function working (dry run)" -ForegroundColor Green
    } else {
        Write-Host "      ❌ User creation function failed" -ForegroundColor Red
    }
    
} catch {
    Write-Host "   ❌ User Management module failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   📍 Error location: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Gray
}

# Test 3: Main Entry Point
Write-Host "`n3️⃣ Testing Main Entry Point..." -ForegroundColor Yellow
if (Test-Path ".\M365-UserProvisioning.ps1") {
    Write-Host "   ✅ Main script exists" -ForegroundColor Green
    
    Write-Host "   🔍 Testing main script in test mode..." -ForegroundColor White
    try {
        # Test without actually running to avoid authentication prompts
        $MainScript = Get-Content ".\M365-UserProvisioning.ps1" -Raw
        if ($MainScript -match "Connect-ToMicrosoftGraph") {
            Write-Host "   ✅ Main script contains expected function calls" -ForegroundColor Green
        } else {
            Write-Host "   ⚠️ Main script may need updates" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "   ⚠️ Could not analyze main script: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "   ❌ Main script not found: .\M365-UserProvisioning.ps1" -ForegroundColor Red
}

# Test 4: Module Dependencies
Write-Host "`n4️⃣ Testing Module Dependencies..." -ForegroundColor Yellow
$RequiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Groups')

foreach ($ModuleName in $RequiredModules) {
    $Module = Get-Module -ListAvailable -Name $ModuleName | Select-Object -First 1
    if ($Module) {
        Write-Host "   ✅ $ModuleName (v$($Module.Version))" -ForegroundColor Green
    } else {
        Write-Host "   ❌ $ModuleName not found - install with: Install-Module $ModuleName" -ForegroundColor Red
    }
}

Write-Host "`n🎯 Summary:" -ForegroundColor Cyan
Write-Host "============" -ForegroundColor Cyan

# Final verification
$AuthModule = Get-Module M365.Authentication
$UserModule = Get-Module M365.UserManagement

if ($AuthModule -and $UserModule) {
    Write-Host "✅ Both modules loaded successfully!" -ForegroundColor Green
    Write-Host "📊 Authentication Module: $($AuthModule.ExportedFunctions.Count) functions" -ForegroundColor Cyan
    Write-Host "📊 User Management Module: $($UserModule.ExportedFunctions.Count) functions" -ForegroundColor Cyan
    Write-Host "`n🚀 Ready for next phase: Test with actual M365 connection" -ForegroundColor Magenta
    Write-Host "   Run: .\M365-UserProvisioning.ps1 -NoGUI -TestMode" -ForegroundColor Yellow
} else {
    Write-Host "❌ One or more modules failed to load properly" -ForegroundColor Red
    Write-Host "📝 Review the errors above and fix the module files" -ForegroundColor Yellow
}

Write-Host "`n📍 Next Steps:" -ForegroundColor Yellow
Write-Host "1. If tests pass: Test with real connection" -ForegroundColor White  
Write-Host "2. If tests fail: Fix the specific errors shown above" -ForegroundColor White
Write-Host "3. Extract GUI functions from legacy script" -ForegroundColor White