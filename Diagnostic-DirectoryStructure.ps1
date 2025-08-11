# Diagnostic-DirectoryStructure.ps1
# Let's see what you actually have in your project

Write-Host "🔍 M365 User Provisioning Tool - Directory Structure Diagnostic" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan

# Get current location
$CurrentLocation = Get-Location
Write-Host "`n📍 Current Directory: $CurrentLocation" -ForegroundColor Yellow

# Check if we're in the right place
Write-Host "`n🗂️  Directory Contents:" -ForegroundColor Yellow
Get-ChildItem -Path . | ForEach-Object {
    if ($_.PSIsContainer) {
        Write-Host "   📁 $($_.Name)/" -ForegroundColor Cyan
    } else {
        Write-Host "   📄 $($_.Name)" -ForegroundColor White
        
        # Check for main scripts
        if ($_.Name -like "*UserProvisioning*" -and $_.Extension -eq ".ps1") {
            Write-Host "      ⭐ MAIN SCRIPT FOUND!" -ForegroundColor Green
        }
    }
}

# Check for Modules directory
Write-Host "`n📦 Checking for Modules directory..." -ForegroundColor Yellow
if (Test-Path ".\Modules") {
    Write-Host "   ✅ Modules directory exists" -ForegroundColor Green
    
    Write-Host "`n   📁 Modules contents:" -ForegroundColor Cyan
    Get-ChildItem ".\Modules" -Directory | ForEach-Object {
        Write-Host "      📁 $($_.Name)/" -ForegroundColor Gray
        
        # Check each module directory
        $ModulePath = $_.FullName
        $PsmFile = Join-Path $ModulePath "$($_.Name).psm1"
        $PsdFile = Join-Path $ModulePath "$($_.Name).psd1"
        
        if (Test-Path $PsmFile) {
            Write-Host "         ✅ $($_.Name).psm1" -ForegroundColor Green
        } else {
            Write-Host "         ❌ $($_.Name).psm1 (missing)" -ForegroundColor Red
        }
        
        if (Test-Path $PsdFile) {
            Write-Host "         ✅ $($_.Name).psd1" -ForegroundColor Green
        } else {
            Write-Host "         ❌ $($_.Name).psd1 (missing)" -ForegroundColor Red
        }
    }
} else {
    Write-Host "   ❌ Modules directory not found" -ForegroundColor Red
    Write-Host "   💡 You may have a single-script architecture" -ForegroundColor Yellow
}

# Check for main scripts
Write-Host "`n🎯 Looking for main scripts..." -ForegroundColor Yellow
$MainScripts = Get-ChildItem -Path . -Filter "*UserProvisioning*.ps1"

if ($MainScripts.Count -gt 0) {
    Write-Host "   ✅ Found main scripts:" -ForegroundColor Green
    foreach ($Script in $MainScripts) {
        Write-Host "      📄 $($Script.Name)" -ForegroundColor Cyan
        
        # Check first few lines for module imports
        $FirstLines = Get-Content $Script.FullName -First 50 | Where-Object { $_ -match "Import-Module|RequiredModules|\\\$RequiredModules" }
        if ($FirstLines) {
            Write-Host "         🔍 Found module references:" -ForegroundColor Gray
            foreach ($Line in $FirstLines) {
                Write-Host "            $($Line.Trim())" -ForegroundColor DarkGray
            }
        }
    }
} else {
    Write-Host "   ❌ No main scripts found" -ForegroundColor Red
}

# Check for M365.ExchangeOnline specifically
Write-Host "`n🔍 Checking M365.ExchangeOnline module..." -ForegroundColor Yellow
$ExchangeModulePath = ".\Modules\M365.ExchangeOnline"

if (Test-Path $ExchangeModulePath) {
    Write-Host "   ✅ M365.ExchangeOnline directory exists" -ForegroundColor Green
    
    $ExchangePsm = Join-Path $ExchangeModulePath "M365.ExchangeOnline.psm1"
    $ExchangePsd = Join-Path $ExchangeModulePath "M365.ExchangeOnline.psd1"
    
    if (Test-Path $ExchangePsm) {
        Write-Host "   ✅ M365.ExchangeOnline.psm1 exists" -ForegroundColor Green
        $FileSize = (Get-Item $ExchangePsm).Length
        Write-Host "      📊 File size: $FileSize bytes" -ForegroundColor Gray
    } else {
        Write-Host "   ❌ M365.ExchangeOnline.psm1 missing" -ForegroundColor Red
    }
    
    if (Test-Path $ExchangePsd) {
        Write-Host "   ✅ M365.ExchangeOnline.psd1 exists" -ForegroundColor Green
        
        # Try to test the manifest
        try {
            $Manifest = Test-ModuleManifest $ExchangePsd -ErrorAction Stop
            Write-Host "      ✅ Module manifest is valid" -ForegroundColor Green
            Write-Host "      📋 Version: $($Manifest.Version)" -ForegroundColor Gray
            Write-Host "      📋 Functions: $($Manifest.ExportedFunctions.Count)" -ForegroundColor Gray
        } catch {
            Write-Host "      ❌ Module manifest has issues: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "   ❌ M365.ExchangeOnline.psd1 missing" -ForegroundColor Red
    }
} else {
    Write-Host "   ❌ M365.ExchangeOnline directory not found" -ForegroundColor Red
}

# Check other expected modules
Write-Host "`n🔍 Checking for other expected modules..." -ForegroundColor Yellow
$ExpectedModules = @("M365.Authentication", "M365.UserManagement", "M365.GUI", "M365.Utilities")

foreach ($ModuleName in $ExpectedModules) {
    $ModulePath = ".\Modules\$ModuleName"
    if (Test-Path $ModulePath) {
        Write-Host "   ✅ $ModuleName directory exists" -ForegroundColor Green
    } else {
        Write-Host "   ❌ $ModuleName directory missing" -ForegroundColor Red
    }
}

# Summary and recommendations
Write-Host "`n🎯 DIAGNOSTIC SUMMARY" -ForegroundColor Cyan
Write-Host "=====================" -ForegroundColor Cyan

if (Test-Path ".\Modules") {
    Write-Host "✅ You have a modular architecture setup" -ForegroundColor Green
    
    if (Test-Path ".\Modules\M365.ExchangeOnline") {
        Write-Host "✅ M365.ExchangeOnline module is in place" -ForegroundColor Green
        Write-Host "`n📋 Next Steps:" -ForegroundColor Yellow
        Write-Host "1. Fix the test script paths (I'll provide a corrected version)" -ForegroundColor White
        Write-Host "2. Update your main script to include M365.ExchangeOnline" -ForegroundColor White
        Write-Host "3. Test the integration" -ForegroundColor White
    } else {
        Write-Host "⚠️  M365.ExchangeOnline module needs to be added" -ForegroundColor Yellow
    }
} else {
    Write-Host "📝 You appear to have a single-script architecture" -ForegroundColor Yellow
    Write-Host "`n📋 Options:" -ForegroundColor Yellow
    Write-Host "1. Convert to modular architecture (recommended)" -ForegroundColor White
    Write-Host "2. Add Exchange functionality directly to your main script" -ForegroundColor White
    Write-Host "3. Create a hybrid approach" -ForegroundColor White
}

Write-Host "`n🔧 QUICK FIXES AVAILABLE:" -ForegroundColor Green
Write-Host "- I can provide a corrected test script for your setup" -ForegroundColor White
Write-Host "- I can show you how to integrate with your existing script" -ForegroundColor White
Write-Host "- I can help set up the missing pieces" -ForegroundColor White