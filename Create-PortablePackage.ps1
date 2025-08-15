# Create a portable package for non-admin users
# This script packages everything needed into a single folder

Write-Host "Creating portable package for non-admin deployment..." -ForegroundColor Cyan

# Create package directory
$packageDir = ".\M365-Tool-Portable"
New-Item -ItemType Directory -Path $packageDir -Force | Out-Null

# Copy main tool files
Copy-Item ".\M365-UserProvisioning-Enterprise.ps1" -Destination $packageDir
Copy-Item ".\Start-Tool.bat" -Destination $packageDir
Copy-Item ".\Install-Prerequisites.ps1" -Destination $packageDir
Copy-Item ".\USER-GUIDE.md" -Destination $packageDir
Copy-Item ".\README.md" -Destination $packageDir

# Copy folders
Copy-Item ".\Modules" -Destination $packageDir -Recurse
Copy-Item ".\Templates" -Destination $packageDir -Recurse
Copy-Item ".\Tests" -Destination $packageDir -Recurse

# Create directories for runtime
New-Item -ItemType Directory -Path "$packageDir\Logs" -Force | Out-Null
New-Item -ItemType Directory -Path "$packageDir\Config" -Force | Out-Null

# Create non-admin startup script
$nonAdminScript = @"
@echo off
echo M365 User Provisioning Tool - Non-Admin Mode
echo ============================================
echo.

REM Check if PowerShell 7 is available
pwsh -Version >nul 2>&1
if errorlevel 1 (
    echo ERROR: PowerShell 7 not found
    echo.
    echo Please ask your IT team to install PowerShell 7
    echo Download: https://github.com/PowerShell/PowerShell/releases
    echo Or via Microsoft Store: ms-windows-store://pdp/?productid=9MZ1SNWT0N5D
    echo.
    pause
    exit /b 1
)

echo Starting M365 User Provisioning Tool...
pwsh -ExecutionPolicy Bypass -File "M365-UserProvisioning-Enterprise.ps1"
pause
"@

$nonAdminScript | Out-File -FilePath "$packageDir\Start-Tool-NonAdmin.bat" -Encoding ASCII

# Create installation guide for non-admin users
$nonAdminGuide = @"
# M365 Tool - Non-Admin Installation Guide

## For Users Without Administrator Rights

### Step 1: PowerShell 7 Installation
You need PowerShell 7 to run this tool. Try these options:

**Option A: Microsoft Store (Recommended)**
1. Open Microsoft Store
2. Search for "PowerShell"
3. Install "PowerShell" by Microsoft Corporation

**Option B: Portable Version**
1. Download portable PowerShell 7 from: https://github.com/PowerShell/PowerShell/releases
2. Extract to a folder (e.g., C:\Tools\PowerShell7)
3. Add to your PATH or run directly

**Option C: Ask IT Team**
Request IT to install PowerShell 7 via Group Policy

### Step 2: Module Installation
1. Double-click `Start-Tool-NonAdmin.bat`
2. If modules are missing, run: `.\Install-Prerequisites.ps1`
3. Modules will install to your user profile (no admin needed)

### Step 3: Execution Policy (if needed)
If you get "script execution disabled" error:
1. Open PowerShell 7
2. Run: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
3. Type 'Y' to confirm

### Step 4: Run the Tool
- Use `Start-Tool-NonAdmin.bat` for easiest startup
- All data stored in your user profile
- No system-wide changes made

## Troubleshooting

**"PowerShell 7 not found"**
- Install PowerShell 7 using one of the options above

**"Module installation failed"**
- Check internet connection
- Verify corporate firewall allows PowerShell Gallery access
- Contact IT if proxy authentication is required

**"Execution of scripts is disabled"**
- Run the execution policy command from Step 3 above

## Need Help?
Contact your IT team or system administrator.
"@

$nonAdminGuide | Out-File -FilePath "$packageDir\NON-ADMIN-GUIDE.md" -Encoding UTF8

Write-Host "Portable package created in: $packageDir" -ForegroundColor Green
Write-Host "Files included:" -ForegroundColor Cyan
Write-Host "  - All tool files and modules" -ForegroundColor White
Write-Host "  - Start-Tool-NonAdmin.bat (special launcher)" -ForegroundColor White
Write-Host "  - NON-ADMIN-GUIDE.md (installation instructions)" -ForegroundColor White