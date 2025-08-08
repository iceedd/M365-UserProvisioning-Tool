#Requires -Version 7.0

<#
.SYNOPSIS
    Debug version to test M365 Tool directly in current session
    
.DESCRIPTION
    This script runs a minimal version of your tool directly in the current PowerShell session
    to help identify exactly where the problem is occurring.
    
.EXAMPLE
    .\Debug-M365Tool.ps1
#>

[CmdletBinding()]
param()

Write-Host "M365 User Provisioning Tool - Debug Mode" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

try {
    Write-Host "🔍 Step 1: Testing Windows Forms initialization..." -ForegroundColor Cyan
    
    # Test if Windows Forms assemblies can be loaded
    Write-Host "   Loading System.Windows.Forms..." -ForegroundColor Yellow
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Write-Host "   ✅ System.Windows.Forms loaded" -ForegroundColor Green
    
    Write-Host "   Loading System.Drawing..." -ForegroundColor Yellow
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Write-Host "   ✅ System.Drawing loaded" -ForegroundColor Green
    
    Write-Host "   Enabling visual styles..." -ForegroundColor Yellow
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Write-Host "   ✅ Visual styles enabled" -ForegroundColor Green
    
    Write-Host "   Setting compatible text rendering..." -ForegroundColor Yellow
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
    Write-Host "   ✅ SetCompatibleTextRenderingDefault succeeded!" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "🔍 Step 2: Testing basic form creation..." -ForegroundColor Cyan
    
    $TestForm = New-Object System.Windows.Forms.Form
    $TestForm.Text = "M365 Tool Debug Test"
    $TestForm.Size = New-Object System.Drawing.Size(400, 300)
    $TestForm.StartPosition = "CenterScreen"
    
    $TestLabel = New-Object System.Windows.Forms.Label
    $TestLabel.Text = "Windows Forms is working correctly!`n`nIf you can see this, your environment supports Windows Forms."
    $TestLabel.Location = New-Object System.Drawing.Point(20, 20)
    $TestLabel.Size = New-Object System.Drawing.Size(360, 100)
    $TestLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    
    $TestButton = New-Object System.Windows.Forms.Button
    $TestButton.Text = "Close Test"
    $TestButton.Location = New-Object System.Drawing.Point(150, 150)
    $TestButton.Size = New-Object System.Drawing.Size(100, 30)
    $TestButton.Add_Click({ $TestForm.Close() })
    
    $TestForm.Controls.Add($TestLabel)
    $TestForm.Controls.Add($TestButton)
    
    Write-Host "   ✅ Test form created successfully" -ForegroundColor Green
    Write-Host ""
    Write-Host "🖥️  Showing test form..." -ForegroundColor Green
    Write-Host "   (Close the form to continue)" -ForegroundColor Gray
    
    $TestForm.ShowDialog() | Out-Null
    $TestForm.Dispose()
    
    Write-Host "   ✅ Test form closed successfully" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "🔍 Step 3: Testing module imports..." -ForegroundColor Cyan
    
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ModulesPath = Join-Path $ScriptPath "Modules"
    
    if (Test-Path $ModulesPath) {
        Write-Host "   ✅ Modules directory found: $ModulesPath" -ForegroundColor Green
        
        $ExpectedModules = @('M365.Authentication', 'M365.UserManagement', 'M365.GUI', 'M365.Utilities')
        
        foreach ($ModuleName in $ExpectedModules) {
            $ModulePath = Join-Path $ModulesPath $ModuleName
            if (Test-Path $ModulePath) {
                Write-Host "   📦 Testing $ModuleName..." -ForegroundColor Yellow
                try {
                    Import-Module $ModulePath -Force -ErrorAction Stop
                    Write-Host "   ✅ $ModuleName imported successfully" -ForegroundColor Green
                }
                catch {
                    Write-Host "   ❌ $ModuleName failed to import: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            else {
                Write-Host "   ⚠️  $ModuleName not found at $ModulePath" -ForegroundColor Yellow
            }
        }
    }
    else {
        Write-Host "   ⚠️  Modules directory not found - this is okay for standalone mode" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "🔍 Step 4: Testing Microsoft Graph modules..." -ForegroundColor Cyan
    
    $GraphModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users'
    )
    
    foreach ($Module in $GraphModules) {
        if (Get-Module -ListAvailable -Name $Module) {
            Write-Host "   ✅ $Module is available" -ForegroundColor Green
        }
        else {
            Write-Host "   ⚠️  $Module is not installed" -ForegroundColor Yellow
        }
    }
    
    Write-Host ""
    Write-Host "🎉 ALL BASIC TESTS PASSED!" -ForegroundColor Green
    Write-Host ""
    Write-Host "✅ Your environment can run Windows Forms applications" -ForegroundColor Green
    Write-Host "✅ The SetCompatibleTextRenderingDefault error should not occur" -ForegroundColor Green
    Write-Host ""
    Write-Host "🔧 Next steps:" -ForegroundColor Cyan
    Write-Host "   1. Try running your main script directly: .\M365-UserProvisioning.ps1" -ForegroundColor White
    Write-Host "   2. If that fails, check for specific error messages" -ForegroundColor White
    Write-Host "   3. Consider using the standalone version" -ForegroundColor White
    
}
catch {
    Write-Host ""
    Write-Host "❌ ERROR DETECTED:" -ForegroundColor Red
    Write-Host "   $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "📍 Error occurred at:" -ForegroundColor Yellow
    Write-Host "   Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor White
    Write-Host "   Position: $($_.InvocationInfo.PositionMessage)" -ForegroundColor White
    
    if ($_.Exception.InnerException) {
        Write-Host ""
        Write-Host "🔍 Inner exception:" -ForegroundColor Yellow
        Write-Host "   $($_.Exception.InnerException.Message)" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "🚨 This is likely the same error causing your main script to fail!" -ForegroundColor Red
}

Write-Host ""
Write-Host "📋 Debug session completed" -ForegroundColor Gray
Write-Host "Press any key to exit..." -ForegroundColor Gray
Read-Host