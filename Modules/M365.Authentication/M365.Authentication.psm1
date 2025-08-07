# M365.Authentication.psm1 - COMPLETE FIXED VERSION
# Microsoft Graph and Exchange Online authentication module for M365 User Provisioning Tool

<#
.SYNOPSIS
    Authentication module for M365 services
.DESCRIPTION  
    Handles Microsoft Graph and Exchange Online connections, disconnections, and tenant discovery
.NOTES
    Extracted from M365 User Provisioning Tool legacy script
    Version: 1.0.1 - UI Dependencies Removed, All Functions Added
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
        
        # Use device code authentication (browser-based, same method as Graph)
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
    # Check if Exchange Online is already connected
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
            
            # Attempt Exchange Online connection
            Write-Verbose "Checking Exchange Online connectivity..."
            $ExchangeResult = Connect-ExchangeOnlineAtStartup -Force
            
            return @{
                Success = $true
                TenantId = $Context.TenantId
                Account = $Context.Account
                Environment = $Context.Environment
                TenantData = $Script:TenantData
                ExchangeOnlineConnected = $Script:ExchangeOnlineConnected
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
        
        # Discover mailboxes
        Write-Verbose "Discovering mailboxes..."
        try {
            $UserMailboxes = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Mail,MailNickname | Where-Object { $null -ne $_.Mail }
            
            $Script:TenantData.AvailableMailboxes = $UserMailboxes | Select-Object @{
                Name = "Id"; Expression = { $_.Id }
            }, @{
                Name = "DisplayName"; Expression = { $_.DisplayName }
            }, @{
                Name = "EmailAddress"; Expression = { $_.Mail }
            }, @{
                Name = "MailboxType"; Expression = { "User" }
            }
            Write-Verbose "Found $($Script:TenantData.AvailableMailboxes.Count) user mailboxes"
            
            # Discover shared mailboxes
            try {
                $SharedMailboxQuery = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Mail,AccountEnabled | Where-Object { 
                    $_.AccountEnabled -eq $false -and $null -ne $_.Mail 
                }
                
                $Script:TenantData.SharedMailboxes = $SharedMailboxQuery | Select-Object @{
                    Name = "Id"; Expression = { $_.Id }
                }, @{
                    Name = "DisplayName"; Expression = { $_.DisplayName }
                }, @{
                    Name = "EmailAddress"; Expression = { $_.Mail }
                }, @{
                    Name = "MailboxType"; Expression = { "Shared" }
                }
                Write-Verbose "Found $($Script:TenantData.SharedMailboxes.Count) potential shared mailboxes"
            }
            catch {
                Write-Warning "Could not discover shared mailboxes: $($_.Exception.Message)"
                $Script:TenantData.SharedMailboxes = @()
            }
        }
        catch {
            Write-Warning "Mailbox discovery failed: $($_.Exception.Message)"
            $Script:TenantData.AvailableMailboxes = @()
            $Script:TenantData.SharedMailboxes = @()
        }
        
        # Discover SharePoint sites
        Write-Verbose "Discovering SharePoint sites..."
        try {
            $Sites = Get-MgSite -All -ErrorAction Stop
            $Script:TenantData.SharePointSites = $Sites | Select-Object Id, DisplayName, WebUrl, Description
            Write-Verbose "Found $($Script:TenantData.SharePointSites.Count) SharePoint sites"
        }
        catch {
            Write-Warning "SharePoint site discovery failed: $($_.Exception.Message)"
            $Script:TenantData.SharePointSites = @()
        }
        
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
                SharePointSites = $Script:TenantData.SharePointSites.Count
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
        Disconnects from Microsoft Graph and Exchange Online
    .OUTPUTS
        Returns disconnection result
    #>
    try {
        $Results = @()
        
        # Disconnect Exchange Online if connected
        if ($Script:ExchangeOnlineConnected) {
            try {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                $Script:ExchangeOnlineConnected = $false
                Write-Verbose "Disconnected from Exchange Online"
                $Results += "Exchange Online disconnected"
            }
            catch {
                Write-Warning "Exchange Online disconnection may have failed: $($_.Exception.Message)"
                $Results += "Exchange Online disconnection uncertain"
            }
        }
        
        if ($Script:IsConnected) {
            Disconnect-MgGraph -ErrorAction Stop
            $Script:IsConnected = $false
            $Script:TenantInfo = $null
            
            # Clear cached data
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
            
            Write-Verbose "Disconnected from Microsoft Graph"
            $Results += "Microsoft Graph disconnected"
        }
        
        return @{
            Success = $true
            Message = $Results -join ", "
        }
    }
    catch {
        Write-Error "Error during disconnection: $($_.Exception.Message)"
        return @{
            Success = $false
            Message = "Disconnection failed: $($_.Exception.Message)"
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

# Export public functions - ALL 9 FUNCTIONS
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