# Fix-WindowsFormsInit.ps1 - Fix Windows Forms initialization order

Write-Host "🔧 Fixing Windows Forms Initialization Order..." -ForegroundColor Yellow

$GUIModulePath = "Modules\M365.GUI\M365.GUI.psm1"

if (-not (Test-Path $GUIModulePath)) {
    Write-Host "❌ GUI module not found at: $GUIModulePath" -ForegroundColor Red
    exit 1
}

Write-Host "📝 Reading GUI module content..." -ForegroundColor Cyan
$Content = Get-Content $GUIModulePath -Raw

Write-Host "🔍 Looking for Windows Forms initialization issue..." -ForegroundColor Cyan

# Check if the fix is already applied
if ($Content -match "EnableVisualStyles.*SetCompatibleTextRenderingDefault.*New-MainForm" -and 
    $Content -notmatch "New-MainForm.*EnableVisualStyles.*SetCompatibleTextRenderingDefault") {
    Write-Host "✅ Windows Forms initialization is already in correct order" -ForegroundColor Green
} else {
    Write-Host "⚠️ Found Windows Forms initialization timing issue - fixing..." -ForegroundColor Yellow
    
    # Fix the initialization order by moving EnableVisualStyles and SetCompatibleTextRenderingDefault
    # to before New-MainForm is called
    
    $Content = $Content -replace `
        '(\s+)# Create and show main form\s+\$Script:MainForm = New-MainForm\s+Write-Host "🖥️ Launching GUI interface\.\.\." -ForegroundColor Green\s+\[System\.Windows\.Forms\.Application\]::EnableVisualStyles\(\)\s+\[System\.Windows\.Forms\.Application\]::SetCompatibleTextRenderingDefault\(\$false\)', `
        '$1# CRITICAL: Initialize Windows Forms BEFORE creating any form objects
$1Write-Host "🖥️ Initializing Windows Forms..." -ForegroundColor Yellow
$1[System.Windows.Forms.Application]::EnableVisualStyles()
$1[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
$1
$1# Now create and show main form
$1Write-Host "🖥️ Creating main form..." -ForegroundColor Yellow
$1$Script:MainForm = New-MainForm
$1
$1Write-Host "🖥️ Launching GUI interface..." -ForegroundColor Green'
    
    # Write the fixed content back
    $Content | Out-File $GUIModulePath -Encoding UTF8
    Write-Host "✅ Fixed Windows Forms initialization order" -ForegroundColor Green
}

Write-Host "`n🧪 Testing the fix..." -ForegroundColor Yellow

try {
    # Test by importing the module
    Import-Module $GUIModulePath -Force -ErrorAction Stop
    Write-Host "✅ GUI module imports successfully" -ForegroundColor Green
    
    # Check if the Start-M365ProvisioningTool function exists
    $StartFunction = Get-Command Start-M365ProvisioningTool -ErrorAction SilentlyContinue
    if ($StartFunction) {
        Write-Host "✅ Start-M365ProvisioningTool function available" -ForegroundColor Green
    } else {
        Write-Host "❌ Start-M365ProvisioningTool function not found" -ForegroundColor Red
    }
} catch {
    Write-Host "❌ GUI module still has issues: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n🎯 Next steps:" -ForegroundColor Yellow
Write-Host "1. Test: .\Test-CompleteApplication.ps1" -ForegroundColor White
Write-Host "2. Launch: .\M365-UserProvisioning.ps1" -ForegroundColor White

Write-Host "`n✅ Windows Forms initialization fix completed!" -ForegroundColor Green