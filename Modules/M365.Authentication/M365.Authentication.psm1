#Requires -Version 7.0
<#
.SYNOPSIS
    M365 Authentication Module - Complete Version
.DESCRIPTION  
    Handles Microsoft Graph and Exchange Online connections, disconnections, and tenant discovery
.NOTES
    Version: 1.0.0 - Complete Working Version
    Author: Tom Mortiboys
#>

# Module-scoped variables
$Script:IsConnected = $false
$Script:TenantInfo = $null
$Script:ExchangeOnlineConnected = $false
$Script:TenantData = @{
    AvailableLicenses = @()
    AvailableGroups = @()
    AvailableUsers = @()
    AvailableMailboxes = @()
    DistributionLists = @()
    MailEnabledSecurityGroups = @()
    SharedMailboxes = @()
    SharePointSites = @()
    AcceptedDomains = @()
}

function Test-ExchangeOnlineModule {
    <#
    .SYNOPSIS
        Tests if Exchange Online PowerShell module is available
    .OUTPUTS
        Returns hashtable with status and message
    #>
    try {
        Write-Verbose "Checking for Exchange Online PowerShell module..."
        $ExOModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement
        
        if (-not $ExOModule) {
            Write-Warning "Exchange Online PowerShell module not found"
            return @{
                Available = $false
                Message = "Exchange Online PowerShell module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
                Version = $null
            }
        }
        else {
            Write-Verbose "Exchange Online PowerShell module found (Version: $($ExOModule.Version))"
            return @{
                Available = $true
                Message = "Exchange Online PowerShell module available"
                Version = $ExOModule.Version
            }
        }
    }
    catch {
        Write-Warning "Error checking Exchange Online module: $($_.Exception.Message)"
        return @{
            Available = $false
            Message = "Error checking Exchange Online module: $($_.Exception.Message)"
            Version = $null
        }
    }
}

function Connect-ExchangeOnlineAtStartup {
    <#
    .SYNOPSIS
        Attempts to connect to Exchange Online
    .PARAMETER Force
        Skip user prompts and attempt connection directly
    .OUTPUTS
        Returns connection result
    #>
    param(
        [switch]$Force
    )
    
    $ModuleStatus = Test-ExchangeOnlineModule
    
    if (-not $ModuleStatus.Available) {
        Write-Warning $ModuleStatus.Message
        return @{
            Connected = $false
            Message = $ModuleStatus.Message
        }
    }

    try {
        Write-Verbose "Attempting to connect to Exchange Online..."
        Import-Module ExchangeOnlineManagement -Force
        
        Write-Verbose "Opening browser for Exchange Online authentication..."
        
        # Use device code authentication (browser-based)
        Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
        
        $Script:ExchangeOnlineConnected = $true
        Write-Verbose "Successfully connected to Exchange Online"
        
        return @{
            Connected = $true
            Message = "Successfully connected to Exchange Online"
        }
    }
    catch {
        Write-Warning "Exchange Online connection failed: $($_.Exception.Message)"
        $Script:ExchangeOnlineConnected = $false
        
        return @{
            Connected = $false
            Message = "Exchange Online connection failed: $($_.Exception.Message)"
        }
    }
}

function Connect-ExchangeOnlineIfNeeded {
    <#
    .SYNOPSIS
        Connects to Exchange Online if not already connected
    #>
    if ($Script:ExchangeOnlineConnected) {
        return $true
    }
    else {
        Write-Verbose "Exchange Online not connected, attempting to connect..."
        $Result = Connect-ExchangeOnlineAtStartup -Force
        return $Result.Connected
    }
}

function Connect-ToMicrosoftGraph {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph and performs tenant discovery
    .PARAMETER Scopes
        Array of Graph scopes to request
    .OUTPUTS
        Returns connection result with tenant information
    #>
    param(
        [string[]]$Scopes = @(
            "User.ReadWrite.All",
            "Directory.ReadWrite.All", 
            "Group.ReadWrite.All",
            "Organization.Read.All",
            "Domain.Read.All",
            "Sites.Read.All",
            "Mail.Read",
            "MailboxSettings.ReadWrite"
        )
    )
    
    try {
        Write-Verbose "Initiating connection to Microsoft Graph..."
        Write-Verbose "Requesting scopes: $($Scopes -join ', ')"
        
        Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop
        
        $Context = Get-MgContext -ErrorAction Stop
        
        if ($Context -and $Context.TenantId) {
            $Script:IsConnected = $true
            $Script:TenantInfo = $Context
            
            Write-Verbose "Successfully connected to Microsoft Graph"
            Write-Verbose "Tenant ID: $($Context.TenantId)"
            Write-Verbose "Account: $($Context.Account)"
            Write-Verbose "Environment: $($Context.Environment)"
            
            # Start tenant discovery
            Write-Verbose "Starting tenant discovery..."
            $DiscoveryResult = Start-TenantDiscovery
            
            # Note: Exchange Online connection will be prompted separately in main application
            
            return @{
                Success = $true
                TenantId = $Context.TenantId
                Account = $Context.Account
                Environment = $Context.Environment
                TenantData = $Script:TenantData
                ExchangeOnlineConnected = $false  # Will be set to true after user chooses to connect
                Message = "Successfully connected to Microsoft Graph and completed tenant discovery"
            }
        }
        else {
            throw "Connection established but context is invalid"
        }
    }
    catch {
        $ErrorMsg = "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        Write-Error $ErrorMsg
        
        return @{
            Success = $false
            Message = $ErrorMsg
            TenantId = $null
            Account = $null
            Environment = $null
            TenantData = $null
            ExchangeOnlineConnected = $false
        }
    }
}

function Start-TenantDiscovery {
    <#
    .SYNOPSIS
        Discovers tenant resources (users, groups, licenses, etc.)
    .OUTPUTS
        Returns discovery results
    #>
    try {
        Write-Verbose "Starting comprehensive tenant discovery..."
        
        # Discover available licenses
        Write-Verbose "Discovering available licenses..."
        try {
            $Licenses = Get-MgSubscribedSku -ErrorAction Stop
            $Script:TenantData.AvailableLicenses = $Licenses | Select-Object SkuId, SkuPartNumber, DisplayName, @{
                Name = "Available"
                Expression = { 
                    $Available = $_.PrepaidUnits.Enabled - $_.ConsumedUnits
                    if ($Available -lt 0) { 0 } else { $Available }
                }
            }, @{
                Name = "Total"
                Expression = { $_.PrepaidUnits.Enabled }
            }, @{
                Name = "Consumed" 
                Expression = { $_.ConsumedUnits }
            }
            Write-Verbose "Found $($Script:TenantData.AvailableLicenses.Count) license types"
        }
        catch {
            Write-Warning "License discovery failed: $($_.Exception.Message)"
            $Script:TenantData.AvailableLicenses = @()
        }
        
        # Discover all groups
        Write-Verbose "Discovering all groups..."
        $AllGroups = Get-MgGroup -All -ErrorAction Stop
        $Script:TenantData.AvailableGroups = $AllGroups | Select-Object Id, DisplayName, Description, @{
            Name = "GroupType"
            Expression = {
                if ($_.GroupTypes -contains "Unified") { 
                    "Microsoft 365" 
                }
                elseif ($_.SecurityEnabled -eq $true -and $_.MailEnabled -eq $true) { 
                    "Mail-Enabled Security" 
                }
                elseif ($_.SecurityEnabled -eq $true) { 
                    "Security" 
                }
                elseif ($_.MailEnabled -eq $true) { 
                    "Distribution List" 
                }
                else { 
                    "Other" 
                }
            }
        }, GroupTypes, SecurityEnabled, MailEnabled, Mail
        Write-Verbose "Found $($Script:TenantData.AvailableGroups.Count) groups total"
        
        # Separate distribution lists
        $Script:TenantData.DistributionLists = $AllGroups | Where-Object { 
            $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $false 
        } | Select-Object Id, DisplayName, Mail, Description
        Write-Verbose "Found $($Script:TenantData.DistributionLists.Count) distribution lists"
        
        $Script:TenantData.MailEnabledSecurityGroups = $AllGroups | Where-Object { 
            $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $true 
        } | Select-Object Id, DisplayName, Mail, Description
        Write-Verbose "Found $($Script:TenantData.MailEnabledSecurityGroups.Count) mail-enabled security groups"
        
        # Discover users
        Write-Verbose "Discovering users..."
        $Users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Mail,JobTitle,Department -ErrorAction Stop
        $Script:TenantData.AvailableUsers = $Users | Select-Object Id, DisplayName, UserPrincipalName, Mail, JobTitle, Department
        Write-Verbose "Found $($Script:TenantData.AvailableUsers.Count) users"
        
        # Discover domains
        Write-Verbose "Discovering accepted domains..."
        $Domains = Get-MgDomain -ErrorAction Stop
        $Script:TenantData.AcceptedDomains = $Domains | Where-Object { $_.IsVerified -eq $true } | Select-Object Id, @{
            Name = "DomainName"
            Expression = { $_.Id }
        }, IsDefault, IsVerified
        Write-Verbose "Found $($Script:TenantData.AcceptedDomains.Count) verified domains"
        
        Write-Verbose "Tenant discovery completed successfully"
        
        return @{
            Success = $true
            Summary = @{
                Users = $Script:TenantData.AvailableUsers.Count
                Groups = $Script:TenantData.AvailableGroups.Count
                DistributionLists = $Script:TenantData.DistributionLists.Count
                MailEnabledSecurityGroups = $Script:TenantData.MailEnabledSecurityGroups.Count
                SharedMailboxes = $Script:TenantData.SharedMailboxes.Count
                Licenses = $Script:TenantData.AvailableLicenses.Count
                Domains = $Script:TenantData.AcceptedDomains.Count
            }
        }
    }
    catch {
        Write-Error "Error during tenant discovery: $($_.Exception.Message)"
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

function Disconnect-FromMicrosoftGraph {
    <#
    .SYNOPSIS
        Completely disconnects from Microsoft Graph and Exchange Online for tenant switching
    .DESCRIPTION
        Performs a complete disconnection from all M365 services and clears all cached data.
        This function is specifically designed to support tenant switching without restarting the application.
    .OUTPUTS
        Returns disconnection result with detailed status
    #>
    try {
        $Results = @()
        Write-Verbose "Starting complete disconnection for tenant switching..."
        
        # AGGRESSIVE Exchange Online disconnection for tenant switching
        Write-Verbose "Checking Exchange Online connection status..."
        try {
            # Try to get Exchange connection info to see if we're actually connected
            $ExchangeConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($ExchangeConnection) {
                Write-Verbose "Exchange Online connection detected - forcing disconnection"
                Write-Host "🔄 Force disconnecting from Exchange Online..." -ForegroundColor Yellow
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                $Results += "Exchange Online force disconnected"
                Write-Host "✅ Exchange Online disconnected" -ForegroundColor Green
            } else {
                Write-Verbose "No active Exchange Online connection detected"
            }
        }
        catch {
            Write-Verbose "Exchange connection check failed, attempting disconnect anyway"
        }
        
        # Also try disconnection regardless of flags (for tenant switching)
        try {
            Write-Verbose "Attempting additional Exchange Online disconnect..."
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            $Results += "Exchange Online additional disconnect attempted"
        }
        catch {
            Write-Verbose "Additional Exchange disconnect completed with warnings"
        }
        
        # Reset connection flag
        $Script:ExchangeOnlineConnected = $false
        
        # Disconnect Microsoft Graph if connected
        if ($Script:IsConnected) {
            try {
                Write-Verbose "Disconnecting from Microsoft Graph..."
                Disconnect-MgGraph -ErrorAction Stop
                $Script:IsConnected = $false
                Write-Verbose "Successfully disconnected from Microsoft Graph"
                $Results += "Microsoft Graph disconnected successfully"
            }
            catch {
                Write-Warning "Microsoft Graph disconnection may have failed: $($_.Exception.Message)"
                # Force the disconnection flag to false anyway for tenant switching
                $Script:IsConnected = $false
                $Results += "Microsoft Graph disconnection forced (may have warnings)"
            }
        }
        
        # Clear all cached tenant information
        Write-Verbose "Clearing all cached tenant data..."
        $Script:TenantInfo = $null
        $Script:TenantData = @{
            AvailableLicenses = @()
            AvailableGroups = @()
            AvailableUsers = @()
            AvailableMailboxes = @()
            DistributionLists = @()
            MailEnabledSecurityGroups = @()
            SharedMailboxes = @()
            SharePointSites = @()
            AcceptedDomains = @()
        }
        
        # Clear any cached authentication tokens more aggressively
        try {
            Write-Verbose "Clearing Microsoft Graph authentication cache..."
            
            # Clear any cached Graph context
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
            
            # Force clear any remaining context
            if (Get-Command "Clear-MgContext" -ErrorAction SilentlyContinue) {
                Clear-MgContext -ErrorAction SilentlyContinue
            }
            
            # Clear browser authentication cache if possible
            try {
                # Clear Microsoft Graph application cache folder if it exists
                $GraphCacheFolder = Join-Path $env:USERPROFILE ".mg"
                if (Test-Path $GraphCacheFolder) {
                    Write-Verbose "Removing Graph cache folder: $GraphCacheFolder"
                    Remove-Item $GraphCacheFolder -Recurse -Force -ErrorAction SilentlyContinue
                }
                
                # Clear additional cache locations - SUPER AGGRESSIVE for tenant switching
                $TokenCachePaths = @(
                    "$env:LOCALAPPDATA\Microsoft\Graph",
                    "$env:APPDATA\Microsoft\Graph", 
                    "$env:TEMP\Microsoft Graph PowerShell",
                    "$env:LOCALAPPDATA\Microsoft\ExchangeOnlineManagement",
                    "$env:APPDATA\Microsoft\ExchangeOnlineManagement",
                    "$env:TEMP\ExchangeOnlineManagement",
                    "$env:LOCALAPPDATA\Microsoft\MSAL*",
                    "$env:APPDATA\Microsoft\MSAL*"
                )
                
                foreach ($CachePath in $TokenCachePaths) {
                    if (Test-Path $CachePath) {
                        Write-Verbose "Clearing cache path: $CachePath"
                        Remove-Item $CachePath -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            catch {
                Write-Verbose "Cache clearing completed with warnings: $($_.Exception.Message)"
            }
            
            # Force garbage collection
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()  # Second collection to ensure cleanup
            
            Write-Verbose "Aggressive token clearing completed"
        }
        catch {
            Write-Verbose "Token clearing completed with warnings: $($_.Exception.Message)"
        }
        
        Write-Verbose "Complete disconnection finished successfully"
        
        return @{
            Success = $true
            Message = $Results -join ", "
            Details = @{
                GraphDisconnected = -not $Script:IsConnected
                ExchangeDisconnected = -not $Script:ExchangeOnlineConnected
                TenantDataCleared = ($Script:TenantInfo -eq $null)
                ReadyForNewTenant = $true
            }
        }
    }
    catch {
        $ErrorMessage = "Critical error during disconnection: $($_.Exception.Message)"
        Write-Error $ErrorMessage
        
        # Force all disconnection flags to false to allow reconnection attempt
        $Script:IsConnected = $false
        $Script:ExchangeOnlineConnected = $false
        $Script:TenantInfo = $null
        
        return @{
            Success = $false
            Message = $ErrorMessage
            Details = @{
                GraphDisconnected = $true  # Forced
                ExchangeDisconnected = $true  # Forced
                TenantDataCleared = $true  # Forced
                ReadyForNewTenant = $true  # Forced for retry
            }
        }
    }
}

function Get-M365AuthenticationStatus {
    <#
    .SYNOPSIS
        Gets current authentication status for M365 services
    #>
    return @{
        GraphConnected = $Script:IsConnected
        ExchangeOnlineConnected = $Script:ExchangeOnlineConnected
        TenantInfo = $Script:TenantInfo
        TenantData = $Script:TenantData
    }
}

function Get-M365TenantData {
    <#
    .SYNOPSIS
        Gets cached tenant data from discovery
    #>
    return $Script:TenantData
}

function Get-M365ConnectionInfo {
    <#
    .SYNOPSIS
        Gets detailed connection information
    #>
    $Status = Get-M365AuthenticationStatus
    
    if ($Status.GraphConnected) {
        $TenantData = Get-M365TenantData
        return @{
            Connected = $true
            TenantId = $Status.TenantInfo.TenantId
            Account = $Status.TenantInfo.Account
            Environment = $Status.TenantInfo.Environment
            ExchangeOnlineConnected = $Status.ExchangeOnlineConnected
            Summary = @{
                Users = $TenantData.AvailableUsers.Count
                Groups = $TenantData.AvailableGroups.Count
                DistributionLists = $TenantData.DistributionLists.Count
                SharedMailboxes = $TenantData.SharedMailboxes.Count
                Licenses = $TenantData.AvailableLicenses.Count
                Domains = $TenantData.AcceptedDomains.Count
            }
        }
    }
    else {
        return @{
            Connected = $false
            Message = "Not connected to Microsoft Graph"
        }
    }
}

# Export all 9 functions
Export-ModuleMember -Function @(
    'Connect-ToMicrosoftGraph',
    'Disconnect-FromMicrosoftGraph',
    'Start-TenantDiscovery',
    'Connect-ExchangeOnlineAtStartup',
    'Connect-ExchangeOnlineIfNeeded',
    'Test-ExchangeOnlineModule',
    'Get-M365AuthenticationStatus',
    'Get-M365TenantData',
    'Get-M365ConnectionInfo'
)