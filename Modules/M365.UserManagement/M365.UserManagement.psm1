#Requires -Version 7.0
<#
.SYNOPSIS
    M365 User Management Module - Simple Version
.DESCRIPTION
    Basic user management functionality for M365 User Provisioning Tool
.NOTES
    Version: 1.0.0 - Simple Version for Testing
    Author: Tom Mortiboys
#>

function New-M365User {
    <#
    .SYNOPSIS
        Creates a new M365 user
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DisplayName,
        
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory)]
        [string]$Password
    )
    
    Write-Host "Creating user: $DisplayName ($UserPrincipalName)" -ForegroundColor Green
    return @{
        DisplayName = $DisplayName
        UserPrincipalName = $UserPrincipalName
        Id = "test-user-123"
        Status = "Created"
    }
}

function Import-UsersFromCSV {
    <#
    .SYNOPSIS
        Imports users from CSV file
    #>
    param(
        [Parameter(Mandatory)]
        [string]$CSVPath
    )
    
    if (-not (Test-Path $CSVPath)) {
        throw "CSV file not found: $CSVPath"
    }
    
    $Users = Import-Csv -Path $CSVPath
    Write-Host "Found $($Users.Count) users to import" -ForegroundColor Yellow
    
    return @{
        TotalUsers = $Users.Count
        Status = "Ready for import"
    }
}

function New-SecurePassword {
    <#
    .SYNOPSIS
        Generates a secure password
    #>
    param([int]$Length = 16)
    
    $Characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*"
    $Password = ""
    
    for ($i = 0; $i -lt $Length; $i++) {
        $Password += $Characters[(Get-Random -Maximum $Characters.Length)]
    }
    
    return $Password
}

# Export functions
Export-ModuleMember -Function @(
    'New-M365User',
    'Import-UsersFromCSV', 
    'New-SecurePassword'
)