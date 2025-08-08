#Requires -Version 7.0

<#
.SYNOPSIS
    Diagnostic script to check New-M365User function parameters
    
.DESCRIPTION
    This script checks what parameters your New-M365User function accepts and provides
    suggestions for fixing parameter mismatches.
    
.EXAMPLE
    .\Check-UserFunction.ps1
#>

Write-Host "New-M365User Function Diagnostic" -ForegroundColor Green
Write-Host "===============================" -ForegroundColor Cyan
Write-Host ""

# Get script directory and setup paths
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulesPath = Join-Path $ScriptPath "Modules"

# Import M365.UserManagement module
$UserMgmtPath = Join-Path $ModulesPath "M365.UserManagement"

if (Test-Path $UserMgmtPath) {
    try {
        Write-Host "📦 Loading M365.UserManagement module..." -ForegroundColor Cyan
        Import-Module $UserMgmtPath -Force -ErrorAction Stop
        Write-Host "✅ Module loaded successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Failed to load module: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host "❌ M365.UserManagement module not found at: $UserMgmtPath" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Check if New-M365User function exists
Write-Host "🔍 Checking New-M365User function..." -ForegroundColor Cyan

$NewUserCommand = Get-Command "New-M365User" -ErrorAction SilentlyContinue

if ($NewUserCommand) {
    Write-Host "✅ New-M365User function found" -ForegroundColor Green
    Write-Host "   Source: $($NewUserCommand.Source)" -ForegroundColor Gray
    Write-Host "   Module: $($NewUserCommand.ModuleName)" -ForegroundColor Gray
    
    Write-Host ""
    Write-Host "📋 Function Parameters:" -ForegroundColor Yellow
    
    $Parameters = $NewUserCommand.Parameters
    $ParameterCount = $Parameters.Count
    
    Write-Host "   Total parameters: $ParameterCount" -ForegroundColor White
    Write-Host ""
    
    # List all parameters
    foreach ($ParamName in ($Parameters.Keys | Sort-Object)) {
        $Param = $Parameters[$ParamName]
        $Type = $Param.ParameterType.Name
        $IsMandatory = $Param.Attributes | Where-Object { $_.TypeId.Name -eq "ParameterAttribute" } | Select-Object -ExpandProperty Mandatory -First 1
        $MandatoryText = if ($IsMandatory) { " (Mandatory)" } else { " (Optional)" }
        
        Write-Host "      ✓ $ParamName [$Type]$MandatoryText" -ForegroundColor White
    }
    
    Write-Host ""
    
    # Check for common parameter variations
    Write-Host "🔍 Parameter Compatibility Check:" -ForegroundColor Cyan
    
    $ExpectedParams = @{
        "FirstName" = @("FirstName", "GivenName")
        "LastName" = @("LastName", "Surname") 
        "DisplayName" = @("DisplayName")
        "Username/UPN" = @("UserPrincipalName", "Username")
        "Password" = @("Password")
        "Department" = @("Department")
        "JobTitle" = @("JobTitle")
        "Office" = @("Office", "OfficeLocation")
        "Manager" = @("Manager")
        "LicenseType" = @("LicenseType")
        "Groups" = @("Groups")
        "Domain" = @("Domain")
    }
    
    foreach ($Category in $ExpectedParams.Keys) {
        $PossibleParams = $ExpectedParams[$Category]
        $FoundParam = $PossibleParams | Where-Object { $Parameters.ContainsKey($_) } | Select-Object -First 1
        
        if ($FoundParam) {
            Write-Host "   ✅ $Category : $FoundParam" -ForegroundColor Green
        }
        else {
            Write-Host "   ❌ $Category : Not found (looking for: $($PossibleParams -join ', '))" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    
    # Check function definition
    Write-Host "🔍 Function Definition:" -ForegroundColor Cyan
    try {
        $FunctionDef = Get-Content function:\New-M365User -ErrorAction Stop
        $ParamBlock = ($FunctionDef | Where-Object { $_ -match "param\s*\(" }) -join "`n"
        
        if ($ParamBlock) {
            Write-Host "   Parameter block found in function definition" -ForegroundColor Green
        }
        else {
            Write-Host "   ⚠️  Could not parse parameter block" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "   ⚠️  Could not read function definition: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    Write-Host ""
    
    # Provide recommendations
    Write-Host "💡 RECOMMENDATIONS:" -ForegroundColor Magenta
    
    # Check for missing JobTitle parameter specifically
    if (-not $Parameters.ContainsKey("JobTitle")) {
        Write-Host ""
        Write-Host "🔧 ISSUE: JobTitle parameter not found" -ForegroundColor Red
        Write-Host "   This is causing your current error!" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "   SOLUTIONS:" -ForegroundColor Cyan
        Write-Host "   1. Update GUI to use parameter-safe wrapper (recommended)" -ForegroundColor White
        Write-Host "   2. Add JobTitle parameter to your New-M365User function" -ForegroundColor White
        Write-Host "   3. Use the fixed Invoke-CreateUser function provided above" -ForegroundColor White
    }
    
    # Check for other missing common parameters
    $MissingCommonParams = @()
    $CommonParams = @("FirstName", "LastName", "Department", "JobTitle", "Office", "Manager", "LicenseType")
    
    foreach ($CommonParam in $CommonParams) {
        if (-not $Parameters.ContainsKey($CommonParam)) {
            $MissingCommonParams += $CommonParam
        }
    }
    
    if ($MissingCommonParams.Count -gt 0) {
        Write-Host ""
        Write-Host "⚠️  Missing common parameters: $($MissingCommonParams -join ', ')" -ForegroundColor Yellow
        Write-Host "   Consider updating your function to accept these optional parameters" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "🚀 NEXT STEPS:" -ForegroundColor Green
    Write-Host "   1. Replace the Invoke-CreateUser function in your GUI module with the fixed version above" -ForegroundColor White
    Write-Host "   2. The fixed version only passes parameters your function accepts" -ForegroundColor White
    Write-Host "   3. Optional fields will be ignored if your function doesn't support them" -ForegroundColor White
    
}
else {
    Write-Host "❌ New-M365User function not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Possible causes:" -ForegroundColor Yellow
    Write-Host "   • M365.UserManagement module not loaded properly" -ForegroundColor White
    Write-Host "   • Function not exported from module" -ForegroundColor White
    Write-Host "   • Function has a different name" -ForegroundColor White
    
    Write-Host ""
    Write-Host "   Available functions in M365.UserManagement:" -ForegroundColor Cyan
    $AvailableFunctions = Get-Command -Module M365.UserManagement 2>$null
    if ($AvailableFunctions) {
        foreach ($Func in $AvailableFunctions) {
            Write-Host "      • $($Func.Name)" -ForegroundColor White
        }
    }
    else {
        Write-Host "      No functions found or module not loaded" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "📋 Diagnostic completed" -ForegroundColor Gray