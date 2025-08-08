#Requires -Version 7.0

<#
.SYNOPSIS
    Quick Fix Test - Tests module loading for M365 UserProvisioning Tool
    
.DESCRIPTION
    Simple test to see if your modules load correctly and identifies the exact issue
    
.EXAMPLE
    .\Quick-Fix-Test.ps1
#>

Write-Host "M365 UserProvisioning Tool - Quick Fix Test" -ForegroundColor Green
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host ""

# Get script directory and setup paths
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulesPath = Join-Path $ScriptPath "Modules"

Write-Host "📂 Script path: $ScriptPath" -ForegroundColor Gray
Write-Host "📦 Modules path: $ModulesPath" -ForegroundColor Gray
Write-Host ""

# Test 1: Check if modules directory exists
Write-Host "🔍 Test 1: Checking module directory structure..." -ForegroundColor Cyan

if (Test-Path $ModulesPath) {
    Write-Host "   ✅ Modules directory found" -ForegroundColor Green
    
    # List available modules
    $AvailableModules = Get-ChildItem $ModulesPath -Directory
    Write-Host "   📋 Available modules:" -ForegroundColor Yellow
    foreach ($Module in $AvailableModules) {
        $PSM1File = Join-Path $Module.FullName "$($Module.Name).psm1"
        $PSD1File = Join-Path $Module.FullName "$($Module.Name).psd1"
        
        $PSM1Exists = Test-Path $PSM1File
        $PSD1Exists = Test-Path $PSD1File
        
        $Status = if ($PSM1Exists -and $PSD1Exists) { "✅" } elseif ($PSM1Exists) { "⚠️" } else { "❌" }
        Write-Host "      $Status $($Module.Name)" -ForegroundColor White
        
        if (-not $PSM1Exists) {
            Write-Host "         Missing: $($Module.Name).psm1" -ForegroundColor Red
        }
        if (-not $PSD1Exists) {
            Write-Host "         Missing: $($Module.Name).psd1" -ForegroundColor Yellow
        }
    }
}
else {
    Write-Host "   ❌ Modules directory not found: $ModulesPath" -ForegroundColor Red
    Write-Host "   💡 Make sure you're running this from the correct directory" -ForegroundColor Yellow
    exit 1
}

Write-Host ""

# Test 2: Try loading each module individually
Write-Host "🔍 Test 2: Testing individual module loading..." -ForegroundColor Cyan

$RequiredModules = @(
    'M365.Authentication',
    'M365.UserManagement', 
    'M365.GUI'
)

$ModuleResults = @{}

foreach ($ModuleName in $RequiredModules) {
    $ModulePath = Join-Path $ModulesPath $ModuleName
    
    Write-Host "   📦 Testing $ModuleName..." -ForegroundColor Yellow
    
    if (Test-Path $ModulePath) {
        try {
            # Remove module if already loaded
            if (Get-Module $ModuleName) {
                Remove-Module $ModuleName -Force
            }
            
            # Try to import the module
            Import-Module $ModulePath -Force -ErrorAction Stop
            
            # Count exported functions
            $Functions = Get-Command -Module $ModuleName -ErrorAction SilentlyContinue
            $FunctionCount = if ($Functions) { $Functions.Count } else { 0 }
            
            Write-Host "      ✅ $ModuleName loaded successfully ($FunctionCount functions)" -ForegroundColor Green
            $ModuleResults[$ModuleName] = @{ Success = $true; FunctionCount = $FunctionCount; Error = $null }
            
            # Special test for GUI module
            if ($ModuleName -eq "M365.GUI") {
                Write-Host "      🧪 Testing Windows Forms initialization..." -ForegroundColor Cyan
                try {
                    # Test if it can initialize Windows Forms without error
                    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
                    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
                    [System.Windows.Forms.Application]::EnableVisualStyles()
                    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
                    
                    Write-Host "         ✅ Windows Forms test passed" -ForegroundColor Green
                }
                catch {
                    Write-Host "         ❌ Windows Forms test failed: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "         🔧 This is the issue causing your main script to fail!" -ForegroundColor Yellow
                    $ModuleResults[$ModuleName].WindowsFormsError = $_.Exception.Message
                }
            }
        }
        catch {
            Write-Host "      ❌ $ModuleName failed to load: $($_.Exception.Message)" -ForegroundColor Red
            $ModuleResults[$ModuleName] = @{ Success = $false; FunctionCount = 0; Error = $_.Exception.Message }
        }
    }
    else {
        Write-Host "      ❌ $ModuleName directory not found: $ModulePath" -ForegroundColor Red
        $ModuleResults[$ModuleName] = @{ Success = $false; FunctionCount = 0; Error = "Directory not found" }
    }
}

Write-Host ""

# Test 3: Try the main script logic
Write-Host "🔍 Test 3: Simulating main script module verification..." -ForegroundColor Cyan

try {
    # This simulates what your main script does at line 84
    $AuthFunctions = if ($ModuleResults['M365.Authentication'].Success) { $ModuleResults['M365.Authentication'].FunctionCount } else { 0 }
    $UserMgmtFunctions = if ($ModuleResults['M365.UserManagement'].Success) { $ModuleResults['M365.UserManagement'].FunctionCount } else { 0 }
    $GUIFunctions = if ($ModuleResults['M365.GUI'].Success) { $ModuleResults['M365.GUI'].FunctionCount } else { 0 }
    
    Write-Host "   📊 Function counts:" -ForegroundColor Yellow
    Write-Host "      Authentication: $AuthFunctions" -ForegroundColor White
    Write-Host "      User Management: $UserMgmtFunctions" -ForegroundColor White
    Write-Host "      GUI: $GUIFunctions" -ForegroundColor White
    
    # This would be line 84 equivalent
    if ($GUIFunctions -eq 0) {
        Write-Host "   ❌ GUI module has 0 functions - this causes the Count error!" -ForegroundColor Red
    }
    else {
        Write-Host "   ✅ All modules have functions - Count error should not occur" -ForegroundColor Green
    }
}
catch {
    Write-Host "   ❌ Module verification failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   🔧 This is likely the same error as line 84 in your main script" -ForegroundColor Yellow
}

Write-Host ""

# Summary and recommendations
Write-Host "📋 SUMMARY & RECOMMENDATIONS" -ForegroundColor Magenta
Write-Host "============================" -ForegroundColor Magenta

$AllModulesWorking = $ModuleResults.Values | ForEach-Object { $_.Success } | Where-Object { $_ -eq $false } | Measure-Object | Select-Object -ExpandProperty Count
$AllModulesWorking = $AllModulesWorking -eq 0

if ($AllModulesWorking) {
    if ($ModuleResults['M365.GUI'].WindowsFormsError) {
        Write-Host "🔧 ISSUE IDENTIFIED: Windows Forms initialization error in GUI module" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "💡 SOLUTION:" -ForegroundColor Cyan
        Write-Host "   1. Replace your M365.GUI.psm1 with the fixed version provided above" -ForegroundColor White
        Write-Host "   2. The fixed version removes SetCompatibleTextRenderingDefault" -ForegroundColor White
        Write-Host "   3. This will resolve your line 84 Count error" -ForegroundColor White
    }
    else {
        Write-Host "✅ All modules are working correctly!" -ForegroundColor Green
        Write-Host "💡 Your M365-UserProvisioning-Enterprise.ps1 should work now" -ForegroundColor Cyan
    }
}
else {
    Write-Host "❌ Some modules failed to load" -ForegroundColor Red
    Write-Host ""
    Write-Host "💡 SOLUTIONS:" -ForegroundColor Cyan
    
    foreach ($Module in $ModuleResults.Keys) {
        if (-not $ModuleResults[$Module].Success) {
            Write-Host "   • Fix $Module : $($ModuleResults[$Module].Error)" -ForegroundColor White
        }
    }
}

Write-Host ""
Write-Host "🚀 NEXT STEPS:" -ForegroundColor Green
Write-Host "   1. If Windows Forms error detected, use the fixed M365.GUI.psm1 above" -ForegroundColor White
Write-Host "   2. Copy the fixed module content to: $ModulesPath\M365.GUI\M365.GUI.psm1" -ForegroundColor White
Write-Host "   3. Run your M365-UserProvisioning-Enterprise.ps1 again" -ForegroundColor White

Write-Host ""
Write-Host "📋 Quick fix test completed" -ForegroundColor Gray

# Cleanup loaded modules
foreach ($ModuleName in $RequiredModules) {
    if (Get-Module $ModuleName) {
        Remove-Module $ModuleName -Force -ErrorAction SilentlyContinue
    }
}