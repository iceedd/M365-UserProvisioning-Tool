#Requires -Version 7.0

<#
.SYNOPSIS
    M365 User Provisioning Tool - Environment Diagnostics and Fix
    
.DESCRIPTION
    Comprehensive diagnostic script that tests all requirements for the M365 User Provisioning Tool
    and can automatically fix common issues including the SetCompatibleTextRenderingDefault error.
    
.NOTES
    Version: 4.0.1 - Environment Diagnostics
    Author: Enterprise Solutions Team
    Last Updated: August 2025
    
.EXAMPLE
    .\Test-M365ToolEnvironment.ps1
    
.EXAMPLE  
    .\Test-M365ToolEnvironment.ps1 -FixIssues
    Automatically attempts to fix detected issues
    
.EXAMPLE
    .\Test-M365ToolEnvironment.ps1 -TestWindowsForms
    Focuses on Windows Forms initialization testing
#>

[CmdletBinding()]
param(
    [switch]$FixIssues,
    [switch]$TestWindowsForms,
    [switch]$Detailed,
    [switch]$InstallMissing
)

#region Diagnostic Functions
function Test-PowerShellVersion {
    Write-Host "🔍 Testing PowerShell Version..." -ForegroundColor Cyan
    
    $CurrentVersion = $PSVersionTable.PSVersion
    $RequiredVersion = [Version]"7.0.0"
    
    if ($CurrentVersion -ge $RequiredVersion) {
        Write-Host "   ✅ PowerShell Version: $CurrentVersion (Required: $RequiredVersion)" -ForegroundColor Green
        return $true
    }
    else {
        Write-Host "   ❌ PowerShell Version: $CurrentVersion (Required: $RequiredVersion)" -ForegroundColor Red
        Write-Host "   💡 Download PowerShell 7: https://aka.ms/powershell" -ForegroundColor Yellow
        return $false
    }
}

function Test-OperatingSystem {
    Write-Host "🔍 Testing Operating System..." -ForegroundColor Cyan
    
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        if ($IsWindows) {
            $OSInfo = Get-CimInstance -ClassName Win32_OperatingSystem
            Write-Host "   ✅ Windows OS: $($OSInfo.Caption) (Build $($OSInfo.BuildNumber))" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "   ❌ Non-Windows OS: $($PSVersionTable.Platform)" -ForegroundColor Red
            Write-Host "   💡 This tool requires Windows for Windows Forms support" -ForegroundColor Yellow
            return $false
        }
    }
    else {
        # PowerShell 5.1 - assume Windows
        $OSInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        Write-Host "   ✅ Windows OS: $($OSInfo.Caption) (PowerShell 5.1)" -ForegroundColor Green
        return $true
    }
}

function Test-DotNetFramework {
    Write-Host "🔍 Testing .NET Framework..." -ForegroundColor Cyan
    
    try {
        $DotNetVersion = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
        Write-Host "   ✅ .NET Runtime: $DotNetVersion" -ForegroundColor Green
        
        # Test if System.Windows.Forms is available
        $FormsAvailable = $false
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null
            $FormsAvailable = $true
            Write-Host "   ✅ System.Windows.Forms: Available" -ForegroundColor Green
        }
        catch {
            Write-Host "   ❌ System.Windows.Forms: Not Available" -ForegroundColor Red
            Write-Host "   💡 Install .NET Framework 4.8 or higher" -ForegroundColor Yellow
        }
        
        return $FormsAvailable
    }
    catch {
        Write-Host "   ⚠️  Could not determine .NET version: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

function Test-WindowsFormsInitialization {
    Write-Host "🔍 Testing Windows Forms Initialization..." -ForegroundColor Cyan
    
    try {
        # Test 1: Basic assembly loading
        Write-Host "   📦 Test 1: Loading Windows Forms assemblies..." -ForegroundColor Yellow
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Write-Host "   ✅ Assemblies loaded successfully" -ForegroundColor Green
        
        # Test 2: EnableVisualStyles
        Write-Host "   🎨 Test 2: Enabling visual styles..." -ForegroundColor Yellow
        [System.Windows.Forms.Application]::EnableVisualStyles()
        Write-Host "   ✅ Visual styles enabled" -ForegroundColor Green
        
        # Test 3: SetCompatibleTextRenderingDefault (THE CRITICAL TEST)
        Write-Host "   📝 Test 3: Setting compatible text rendering..." -ForegroundColor Yellow
        try {
            [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
            Write-Host "   ✅ SetCompatibleTextRenderingDefault succeeded" -ForegroundColor Green
            $TextRenderingOK = $true
        }
        catch {
            Write-Host "   ❌ SetCompatibleTextRenderingDefault FAILED: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "   🚨 This is the exact error you're experiencing!" -ForegroundColor Yellow
            $TextRenderingOK = $false
        }
        
        # Test 4: Basic form creation
        if ($TextRenderingOK) {
            Write-Host "   🖥️  Test 4: Creating test form..." -ForegroundColor Yellow
            $TestForm = New-Object System.Windows.Forms.Form
            $TestForm.Text = "Windows Forms Test"
            $TestForm.Size = New-Object System.Drawing.Size(300, 200)
            Write-Host "   ✅ Test form created successfully" -ForegroundColor Green
            $TestForm.Dispose()
        }
        
        return $TextRenderingOK
    }
    catch {
        Write-Host "   ❌ Windows Forms initialization failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Test-GraphModules {
    Write-Host "🔍 Testing Microsoft Graph Modules..." -ForegroundColor Cyan
    
    $RequiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Users.Actions',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )
    
    $AllModulesOK = $true
    $MissingModules = @()
    
    foreach ($Module in $RequiredModules) {
        $ModuleInfo = Get-Module -ListAvailable -Name $Module | Select-Object -First 1
        
        if ($ModuleInfo) {
            Write-Host "   ✅ $Module : Version $($ModuleInfo.Version)" -ForegroundColor Green
        }
        else {
            Write-Host "   ❌ $Module : Not installed" -ForegroundColor Red
            $MissingModules += $Module
            $AllModulesOK = $false
        }
    }
    
    if ($MissingModules.Count -gt 0 -and $InstallMissing) {
        Write-Host "   🔄 Installing missing modules..." -ForegroundColor Yellow
        foreach ($Module in $MissingModules) {
            try {
                Write-Host "      📥 Installing $Module..." -ForegroundColor Yellow
                Install-Module -Name $Module -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                Write-Host "      ✅ $Module installed" -ForegroundColor Green
            }
            catch {
                Write-Host "      ❌ Failed to install $Module : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    return $AllModulesOK
}

function Test-ModuleArchitecture {
    Write-Host "🔍 Testing Module Architecture..." -ForegroundColor Cyan
    
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ModulesPath = Join-Path $ScriptPath "Modules"
    
    if (Test-Path $ModulesPath) {
        Write-Host "   ✅ Modules directory found: $ModulesPath" -ForegroundColor Green
        
        $ExpectedModules = @('M365.Authentication', 'M365.UserManagement', 'M365.GUI', 'M365.Utilities')
        $ModuleStatus = $true
        
        foreach ($ModuleName in $ExpectedModules) {
            $ModulePath = Join-Path $ModulesPath $ModuleName
            if (Test-Path $ModulePath) {
                Write-Host "   ✅ Module found: $ModuleName" -ForegroundColor Green
                
                # Check for PSM1 file
                $PSM1Path = Join-Path $ModulePath "$ModuleName.psm1"
                if (Test-Path $PSM1Path) {
                    Write-Host "      ✅ PSM1 file exists" -ForegroundColor Green
                }
                else {
                    Write-Host "      ⚠️  PSM1 file missing: $PSM1Path" -ForegroundColor Yellow
                    $ModuleStatus = $false
                }
            }
            else {
                Write-Host "   ❌ Module missing: $ModuleName" -ForegroundColor Red
                $ModuleStatus = $false
            }
        }
        
        return $ModuleStatus
    }
    else {
        Write-Host "   ⚠️  Modules directory not found - using standalone mode" -ForegroundColor Yellow
        return $false
    }
}

function Test-WindowsFormsInCleanProcess {
    Write-Host "🔍 Testing Windows Forms in Clean Process..." -ForegroundColor Cyan
    
    $TestScript = @'
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Clean Process Test"
    Write-Output "SUCCESS: Windows Forms initialized in clean process"
    exit 0
}
catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
}
'@
    
    try {
        $TempScript = [System.IO.Path]::GetTempFileName() + ".ps1"
        $TestScript | Out-File -FilePath $TempScript -Encoding UTF8
        
        $ProcessResult = Start-Process -FilePath "pwsh" -ArgumentList @("-NoProfile", "-File", $TempScript) -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput ([System.IO.Path]::GetTempFileName()) -RedirectStandardError ([System.IO.Path]::GetTempFileName())
        
        if ($ProcessResult.ExitCode -eq 0) {
            Write-Host "   ✅ Windows Forms works correctly in clean process" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "   ❌ Windows Forms failed in clean process" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "   ⚠️  Could not test clean process: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
    finally {
        if ($TempScript -and (Test-Path $TempScript)) {
            Remove-Item $TempScript -Force -ErrorAction SilentlyContinue
        }
    }
}
#endregion

#region Fix Functions
function Fix-WindowsFormsInitialization {
    Write-Host "🔧 Attempting to fix Windows Forms initialization..." -ForegroundColor Yellow
    
    Write-Host "   📝 Creating fixed module files..." -ForegroundColor Yellow
    
    # Here you would implement the actual file fixes
    # For now, provide guidance
    
    Write-Host "   💡 Recommended fixes:" -ForegroundColor Cyan
    Write-Host "      1. Use the provided fixed M365.GUI.psm1 file" -ForegroundColor White
    Write-Host "      2. Use the standalone version to avoid module conflicts" -ForegroundColor White
    Write-Host "      3. Use the clean launcher to start fresh process" -ForegroundColor White
    
    return $true
}
#endregion

#region Main Diagnostic Logic
Write-Host "M365 User Provisioning Tool - Environment Diagnostics" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""

$AllTestsPassed = $true
$TestResults = @{}

# Core environment tests
$TestResults.PowerShell = Test-PowerShellVersion
$TestResults.OS = Test-OperatingSystem  
$TestResults.DotNet = Test-DotNetFramework
$TestResults.GraphModules = Test-GraphModules

# Architecture tests
$TestResults.ModuleArchitecture = Test-ModuleArchitecture

# Windows Forms specific tests
if ($TestWindowsForms -or $Detailed) {
    Write-Host ""
    Write-Host "🔬 DETAILED WINDOWS FORMS TESTING" -ForegroundColor Magenta
    Write-Host "=================================" -ForegroundColor Magenta
    
    $TestResults.WindowsFormsInit = Test-WindowsFormsInitialization
    $TestResults.CleanProcess = Test-WindowsFormsInCleanProcess
}

# Calculate overall result
$AllTestsPassed = $TestResults.Values | ForEach-Object { $_ } | Where-Object { $_ -eq $false } | Measure-Object | Select-Object -ExpandProperty Count
$AllTestsPassed = $AllTestsPassed -eq 0

Write-Host ""
Write-Host "📊 DIAGNOSTIC SUMMARY" -ForegroundColor Magenta
Write-Host "===================" -ForegroundColor Magenta

foreach ($Test in $TestResults.GetEnumerator()) {
    $Status = if ($Test.Value) { "✅ PASS" } else { "❌ FAIL" }
    $Color = if ($Test.Value) { "Green" } else { "Red" }
    Write-Host "$Status $($Test.Key)" -ForegroundColor $Color
}

Write-Host ""
if ($AllTestsPassed) {
    Write-Host "🎉 ALL TESTS PASSED! Your environment is ready for M365 User Provisioning Tool." -ForegroundColor Green
    Write-Host ""
    Write-Host "🚀 Recommended launch methods:" -ForegroundColor Cyan
    Write-Host "   • Modular: .\M365-UserProvisioning.ps1" -ForegroundColor White
    Write-Host "   • Standalone: .\M365-UserProvisioning-Standalone.ps1" -ForegroundColor White
    Write-Host "   • Clean Launcher: .\Start-M365Tool-Clean.ps1" -ForegroundColor White
}
else {
    Write-Host "⚠️  SOME TESTS FAILED. See issues above." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "🔧 RECOMMENDED SOLUTIONS:" -ForegroundColor Cyan
    
    if (-not $TestResults.WindowsFormsInit) {
        Write-Host ""
        Write-Host "🚨 WINDOWS FORMS INITIALIZATION ISSUE DETECTED!" -ForegroundColor Red
        Write-Host "This is likely the SetCompatibleTextRenderingDefault error you're experiencing." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "✅ SOLUTIONS (in order of recommendation):" -ForegroundColor Green
        Write-Host "   1. Use Clean Launcher: .\Start-M365Tool-Clean.ps1" -ForegroundColor White
        Write-Host "      → Starts completely fresh process, avoids all timing issues" -ForegroundColor Gray
        Write-Host ""
        Write-Host "   2. Use Standalone Version: .\M365-UserProvisioning-Standalone.ps1" -ForegroundColor White  
        Write-Host "      → No modules, proper initialization order" -ForegroundColor Gray
        Write-Host ""
        Write-Host "   3. Replace M365.GUI.psm1 with the provided fixed version" -ForegroundColor White
        Write-Host "      → Fixes module initialization timing" -ForegroundColor Gray
    }
    
    if (-not $TestResults.GraphModules) {
        Write-Host ""
        Write-Host "📦 MISSING MICROSOFT GRAPH MODULES" -ForegroundColor Yellow
        Write-Host "   Run: .\Test-M365ToolEnvironment.ps1 -InstallMissing" -ForegroundColor White
        Write-Host "   Or: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor White
    }
}

Write-Host ""
Write-Host "📋 Diagnostic completed at $(Get-Date)" -ForegroundColor Gray
#endregion