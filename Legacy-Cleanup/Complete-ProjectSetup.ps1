# Complete-ProjectSetup.ps1 - Final Project Setup
# Creates all necessary directories and verifies project structure

Write-Host "🏗️ M365 User Provisioning Tool - Final Project Setup" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan

# Create all required directories
$Directories = @(
    "Modules\M365.Authentication",
    "Modules\M365.UserManagement", 
    "Modules\M365.GUI",
    "Tests\Unit",
    "Tests\Integration",
    "Logs",
    "Templates"
)

Write-Host "`n📁 Creating Directory Structure..." -ForegroundColor Yellow
foreach ($Dir in $Directories) {
    if (-not (Test-Path $Dir)) {
        New-Item -ItemType Directory -Path $Dir -Force | Out-Null
        Write-Host "   ✅ Created: $Dir" -ForegroundColor Green
    } else {
        Write-Host "   📂 Exists: $Dir" -ForegroundColor Cyan
    }
}

# Verify all required files
Write-Host "`n📋 Required Files Checklist:" -ForegroundColor Yellow

$RequiredFiles = @(
    @{ Path = "Modules\M365.Authentication\M365.Authentication.psm1"; Name = "Authentication Module" },
    @{ Path = "Modules\M365.Authentication\M365.Authentication.psd1"; Name = "Authentication Manifest" },
    @{ Path = "Modules\M365.UserManagement\M365.UserManagement.psm1"; Name = "User Management Module" },
    @{ Path = "Modules\M365.UserManagement\M365.UserManagement.psd1"; Name = "User Management Manifest" },
    @{ Path = "Modules\M365.GUI\M365.GUI.psm1"; Name = "GUI Module" },
    @{ Path = "Modules\M365.GUI\M365.GUI.psd1"; Name = "GUI Manifest" },
    @{ Path = "M365-UserProvisioning.ps1"; Name = "Main Entry Point" },
    @{ Path = "Test-LiveConnection.ps1"; Name = "Live Connection Test" },
    @{ Path = "Test-CompleteApplication.ps1"; Name = "Complete App Test" }
)

$MissingFiles = @()
foreach ($File in $RequiredFiles) {
    if (Test-Path $File.Path) {
        Write-Host "   ✅ $($File.Name): $($File.Path)" -ForegroundColor Green
    } else {
        Write-Host "   ❌ $($File.Name): $($File.Path) - MISSING" -ForegroundColor Red
        $MissingFiles += $File
    }
}

if ($MissingFiles.Count -eq 0) {
    Write-Host "`n🎉 ALL REQUIRED FILES PRESENT!" -ForegroundColor Green
    Write-Host "`n🚀 Your M365 User Provisioning Tool is ready for use!" -ForegroundColor Magenta
    
    Write-Host "`n📖 Usage Instructions:" -ForegroundColor Yellow
    Write-Host "1. Launch GUI:  .\M365-UserProvisioning.ps1" -ForegroundColor White
    Write-Host "2. Console mode: .\M365-UserProvisioning.ps1 -NoGUI" -ForegroundColor White
    Write-Host "3. Test mode: .\M365-UserProvisioning.ps1 -TestMode" -ForegroundColor White
    Write-Host "4. Test all modules: .\Test-CompleteApplication.ps1" -ForegroundColor White
    
} else {
    Write-Host "`n⚠️ MISSING FILES DETECTED!" -ForegroundColor Yellow
    Write-Host "The following files need to be created:" -ForegroundColor Yellow
    
    foreach ($File in $MissingFiles) {
        Write-Host "   📝 Create: $($File.Path)" -ForegroundColor White
    }
    
    Write-Host "`n💡 Use the artifacts provided earlier to create these files." -ForegroundColor Cyan
}

# Show current project status
Write-Host "`n📊 Project Status:" -ForegroundColor Cyan
$CompletionPercentage = [Math]::Round((($RequiredFiles.Count - $MissingFiles.Count) / $RequiredFiles.Count) * 100, 0)
Write-Host "Progress: $CompletionPercentage% Complete" -ForegroundColor $(if ($CompletionPercentage -eq 100) { 'Green' } else { 'Yellow' })

# Show file sizes for verification
Write-Host "`n📏 Module File Sizes (for verification):" -ForegroundColor Gray
Get-ChildItem "Modules" -Recurse -File | Where-Object { $_.Extension -in @('.psm1', '.psd1') } | 
    ForEach-Object { 
        $SizeKB = [Math]::Round($_.Length / 1KB, 1)
        Write-Host "   $($_.FullName.Replace($PWD, '.')): ${SizeKB}KB" -ForegroundColor Gray 
    }

Write-Host "`n🎯 Next Steps:" -ForegroundColor Yellow
if ($MissingFiles.Count -eq 0) {
    Write-Host "✅ Run: .\Test-CompleteApplication.ps1" -ForegroundColor Green
    Write-Host "✅ Then: .\M365-UserProvisioning.ps1" -ForegroundColor Green
} else {
    Write-Host "1. Create the missing files listed above" -ForegroundColor White
    Write-Host "2. Run this setup script again to verify" -ForegroundColor White  
    Write-Host "3. Test with: .\Test-CompleteApplication.ps1" -ForegroundColor White
}

Write-Host "`n🏆 You're building an enterprise-grade M365 tool! 🏆" -ForegroundColor Magenta