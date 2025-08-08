#Requires -Version 7.0

<#
.SYNOPSIS
    Clean Launcher for M365 User Provisioning Tool
    
.DESCRIPTION
    This launcher starts the M365 User Provisioning Tool in a completely clean PowerShell process,
    avoiding any Windows Forms initialization conflicts that might exist in the current session.
    
    This approach completely bypasses the SetCompatibleTextRenderingDefault error by ensuring
    a fresh process where Windows Forms has never been touched.
    
.NOTES
    Version: 4.0.1 - Clean Process Launcher
    Author: Enterprise Solutions Team
    Last Updated: August 2025
    
    This launcher resolves ALL Windows Forms timing issues by:
    1. Starting a completely fresh PowerShell process
    2. Ensuring no previous Windows Forms state exists
    3. Providing perfect isolation from current session
    
.EXAMPLE
    .\Start-M365Tool-Clean.ps1
    
.EXAMPLE
    .\Start-M365Tool-Clean.ps1 -UseStandalone
    Uses the standalone version instead of modular
    
.EXAMPLE
    .\Start-M365Tool-Clean.ps1 -VerboseLogging
    Runs with verbose output enabled
#>

[CmdletBinding()]
param(
    [switch]$UseStandalone,
    [switch]$VerboseLogging,
    [switch]$DebugMode
)

#region Launcher Configuration
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Determine which script to launch
if ($UseStandalone) {
    $TargetScript = Join-Path $ScriptPath "M365-UserProvisioning-Standalone.ps1"
    $LaunchMode = "Standalone Edition"
}
else {
    $TargetScript = Join-Path $ScriptPath "M365-UserProvisioning.ps1"
    $LaunchMode = "Modular Edition"
}
#endregion

#region Pre-Launch Validation
Write-Host "M365 User Provisioning Tool - Clean Launcher" -ForegroundColor Green
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "🚀 Launch Mode: $LaunchMode" -ForegroundColor Yellow
Write-Host ""

# Check if target script exists
if (-not (Test-Path $TargetScript)) {
    Write-Host "❌ Target script not found: $TargetScript" -ForegroundColor Red
    Write-Host ""
    Write-Host "📁 Available scripts in directory:" -ForegroundColor Yellow
    Get-ChildItem $ScriptPath -Filter "*.ps1" | ForEach-Object {
        Write-Host "   📄 $($_.Name)" -ForegroundColor White
    }
    exit 1
}

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "❌ This tool requires PowerShell 7.0 or higher" -ForegroundColor Red
    Write-Host "   Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host "   Download PowerShell 7: https://aka.ms/powershell" -ForegroundColor White
    exit 1
}

# Check if running on Windows
if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    Write-Host "❌ This tool requires Windows operating system" -ForegroundColor Red
    Write-Host "   Current platform: $($PSVersionTable.Platform)" -ForegroundColor Yellow
    exit 1
}

Write-Host "✅ Pre-launch validation completed" -ForegroundColor Green
Write-Host ""
#endregion

#region Clean Process Launch
try {
    Write-Host "🔄 Starting clean PowerShell process..." -ForegroundColor Cyan
    Write-Host "📂 Script location: $TargetScript" -ForegroundColor Gray
    Write-Host ""
    
    # Build PowerShell arguments
    $PSArgs = @(
        '-NoProfile'           # Don't load user profile
        '-WindowStyle', 'Normal'  # Normal window
        '-ExecutionPolicy', 'Bypass'  # Bypass execution policy
        '-NoExit'              # Keep window open after script runs
        '-Command', "& '$TargetScript'; Write-Host ''; Write-Host 'Script completed. Press any key to close...' -ForegroundColor Yellow; Read-Host"
    )
    
    # Add debug/verbose flags if specified
    $CommandString = "& '$TargetScript'"
    if ($VerboseLogging -or $VerbosePreference -ne 'SilentlyContinue') { 
        $CommandString += " -Verbose" 
    }
    if ($DebugMode -or $DebugPreference -ne 'SilentlyContinue') { 
        $CommandString += " -Debug" 
    }
    $CommandString += "; Write-Host ''; Write-Host 'Script completed. Press any key to close...' -ForegroundColor Yellow; Read-Host"
    
    $PSArgs[4] = $CommandString  # Update the command string
    
    Write-Host "🚀 Launching M365 User Provisioning Tool in clean process..." -ForegroundColor Green
    Write-Host "   PowerShell executable: $((Get-Process -Id $PID).Path)" -ForegroundColor Gray
    Write-Host "   Arguments: $($PSArgs -join ' ')" -ForegroundColor Gray
    Write-Host ""
    Write-Host "ℹ️  The tool will open in a new window. You can close this launcher." -ForegroundColor Cyan
    Write-Host ""
    
    # Start the clean process
    $ProcessInfo = @{
        FilePath = (Get-Process -Id $PID).Path  # Use same PowerShell executable
        ArgumentList = $PSArgs
        WindowStyle = 'Normal'
        PassThru = $true
    }
    
    $Process = Start-Process @ProcessInfo
    
    if ($Process) {
        Write-Host "✅ M365 User Provisioning Tool launched successfully" -ForegroundColor Green
        Write-Host "   Process ID: $($Process.Id)" -ForegroundColor Gray
        Write-Host "   Process Name: $($Process.ProcessName)" -ForegroundColor Gray
        Write-Host ""
        Write-Host "🎯 The application is now running in a clean environment." -ForegroundColor Green
        Write-Host "   This completely avoids any Windows Forms initialization conflicts." -ForegroundColor White
        Write-Host ""
        Write-Host "👋 Launcher task completed. You may close this window." -ForegroundColor Yellow
    }
    else {
        throw "Failed to start PowerShell process"
    }
}
catch {
    Write-Host ""
    Write-Host "❌ Failed to launch M365 User Provisioning Tool" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "🔧 Troubleshooting:" -ForegroundColor Yellow
    Write-Host "   1. Ensure PowerShell 7+ is properly installed" -ForegroundColor White
    Write-Host "   2. Check that the target script exists and is accessible" -ForegroundColor White
    Write-Host "   3. Verify you have permission to execute PowerShell scripts" -ForegroundColor White
    Write-Host "   4. Try running as Administrator" -ForegroundColor White
    Write-Host ""
    Write-Host "💡 Alternative launch methods:" -ForegroundColor Cyan
    Write-Host "   • Right-click the .ps1 file and select 'Run with PowerShell'" -ForegroundColor White
    Write-Host "   • Open PowerShell 7 manually and run: & '$TargetScript'" -ForegroundColor White
    Write-Host "   • Use the standalone version: .\Start-M365Tool-Clean.ps1 -UseStandalone" -ForegroundColor White
    Write-Host "   • Enable verbose logging: .\Start-M365Tool-Clean.ps1 -VerboseLogging" -ForegroundColor White
    
    exit 1
}
#endregion

#region Optional - Wait for Process (Uncomment if needed)
<#
# Uncomment this section if you want the launcher to wait for the application to close

Write-Host "⏳ Waiting for application to close..." -ForegroundColor Yellow
Write-Host "   (Press Ctrl+C to stop waiting and close launcher)" -ForegroundColor Gray

try {
    $Process.WaitForExit()
    Write-Host ""
    Write-Host "✅ M365 User Provisioning Tool has closed" -ForegroundColor Green
    Write-Host "   Exit code: $($Process.ExitCode)" -ForegroundColor Gray
}
catch {
    Write-Host ""
    Write-Host "⚠️  Stopped waiting for application" -ForegroundColor Yellow
}
#>
#endregion