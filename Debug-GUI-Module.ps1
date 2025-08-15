#Requires -Version 7.0

<#
.SYNOPSIS
    Debug GUI module loading to check for Switch Tenant button
#>

Write-Host "🔍 Debugging GUI Module Loading..." -ForegroundColor Yellow

# Force remove any existing modules
Get-Module M365.* | Remove-Module -Force -ErrorAction SilentlyContinue

# Import GUI module directly with force
$GuiModulePath = ".\Modules\M365.GUI\M365.GUI.psm1"
Write-Host "📁 Loading GUI module from: $GuiModulePath" -ForegroundColor Cyan

Import-Module $GuiModulePath -Force -Verbose

# Check what functions are available
Write-Host "📋 Available GUI functions:" -ForegroundColor Green
$GUIFunctions = Get-Command -Module M365.GUI
$GUIFunctions | ForEach-Object { Write-Host "  • $($_.Name)" -ForegroundColor White }

Write-Host ""
Write-Host "🧪 Testing New-MainForm function..." -ForegroundColor Yellow

try {
    # Initialize Windows Forms
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    [System.Windows.Forms.Application]::EnableVisualStyles()
    
    Write-Host "✅ Windows Forms initialized" -ForegroundColor Green
    
    # Test creating the main form
    Write-Host "🏗️ Creating main form..." -ForegroundColor Cyan
    $TestForm = New-MainForm
    
    if ($TestForm) {
        Write-Host "✅ Main form created successfully" -ForegroundColor Green
        
        # Check for Switch Tenant button in form controls
        Write-Host "🔍 Searching for Switch Tenant button..." -ForegroundColor Cyan
        
        $foundButton = $false
        function Search-Controls($control, $depth = 0) {
            $indent = "  " * $depth
            
            if ($control -is [System.Windows.Forms.Button]) {
                Write-Host "$indent🔘 Button: '$($control.Text)' at ($($control.Location.X), $($control.Location.Y))" -ForegroundColor White
                
                if ($control.Text -like "*Switch Tenant*") {
                    Write-Host "$indent✅ FOUND Switch Tenant button!" -ForegroundColor Green
                    Write-Host "$indent   Size: $($control.Size)" -ForegroundColor Yellow
                    Write-Host "$indent   Location: $($control.Location)" -ForegroundColor Yellow
                    Write-Host "$indent   Enabled: $($control.Enabled)" -ForegroundColor Yellow
                    Write-Host "$indent   BackColor: $($control.BackColor)" -ForegroundColor Yellow
                    $script:foundButton = $true
                }
            }
            
            foreach ($child in $control.Controls) {
                Search-Controls $child ($depth + 1)
            }
        }
        
        Search-Controls $TestForm
        
        if (-not $foundButton) {
            Write-Host "❌ Switch Tenant button NOT found in form controls!" -ForegroundColor Red
        }
        
        Write-Host ""
        Write-Host "📊 Form summary:" -ForegroundColor Yellow
        Write-Host "  Title: $($TestForm.Text)" -ForegroundColor White
        Write-Host "  Size: $($TestForm.Size)" -ForegroundColor White
        Write-Host "  Total Controls: $($TestForm.Controls.Count)" -ForegroundColor White
        
        # Clean up
        $TestForm.Dispose()
    }
    else {
        Write-Host "❌ Failed to create main form" -ForegroundColor Red
    }
}
catch {
    Write-Host "❌ Error during testing: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "🏁 Debug completed" -ForegroundColor Green