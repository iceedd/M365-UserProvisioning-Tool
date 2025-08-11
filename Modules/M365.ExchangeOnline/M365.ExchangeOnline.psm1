# M365.ExchangeOnline.psm1 - UPDATED WITH DEVICE CODE AUTHENTICATION
# Exchange Online operations module for M365 User Provisioning Tool
# Works with your existing M365.Authentication module

<#
.SYNOPSIS
    Exchange Online operations module
.DESCRIPTION
    Handles Exchange Online specific operations like shared mailboxes, distribution lists,
    and mail-enabled security groups. Uses Device Code Authentication for easy help desk access.
.NOTES
    Version: 1.0.3 - Device Code Authentication Added
    Author: Tom Mortiboys
    
    UPDATED: Device Code Authentication for help desk users
    ENHANCED: User-friendly authentication process
    MAINTAINED: All existing Exchange functionality
#>

# Module-scoped variables for Exchange data
$Script:ExchangeData = @{
    SharedMailboxes = @()
    DistributionLists = @()
    MailEnabledSecurityGroups = @()
    AcceptedDomains = @()
    LastRefresh = $null
}

# ================================
# CONNECTION HELPER FUNCTIONS
# ================================

function Test-ExchangeOnlineConnection {
    <#
    .SYNOPSIS
        Tests if Exchange Online is connected and cmdlets are available
    .OUTPUTS
        Boolean indicating if Exchange Online is connected and functional
    #>
    [CmdletBinding()]
    param()
    
    try {
        # Method 1: Check if ExchangeOnlineManagement module is loaded
        $ExOModule = Get-Module -Name ExchangeOnlineManagement
        if (-not $ExOModule) {
            Write-Verbose "ExchangeOnlineManagement module not loaded"
            return $false
        }
        
        # Method 2: Try to get connection information
        $ConnectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($ConnectionInfo -and $ConnectionInfo.State -eq 'Connected') {
            Write-Verbose "Exchange Online connected via Get-ConnectionInformation"
            return $true
        }
        
        # Method 3: Test a simple Exchange cmdlet
        $TestResult = Get-AcceptedDomain -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($TestResult) {
            Write-Verbose "Exchange Online connected - Get-AcceptedDomain successful"
            return $true
        }
        
        Write-Verbose "Exchange Online not connected - no active session"
        return $false
    }
    catch {
        Write-Verbose "Exchange Online connection test failed: $($_.Exception.Message)"
        return $false
    }
}

function Connect-ExchangeOnlineIfNeeded {
    <#
    .SYNOPSIS
        Connects to Exchange Online using Device Code Authentication
    .DESCRIPTION
        Uses device code flow with clear instructions for help desk users
        Perfect for help desk environments - no complex setup required!
    .OUTPUTS
        Boolean indicating successful connection
    #>
    [CmdletBinding()]
    param()
    
    # Check if already connected
    if (Test-ExchangeOnlineConnection) {
        Write-Verbose "Exchange Online already connected"
        return $true
    }
    
    try {
        Write-Host ""
        Write-Host "🔐 EXCHANGE ONLINE AUTHENTICATION REQUIRED" -ForegroundColor Yellow -BackgroundColor DarkBlue
        Write-Host "=============================================" -ForegroundColor Yellow -BackgroundColor DarkBlue
        Write-Host ""
        Write-Host "📱 Device Code Authentication Process:" -ForegroundColor Cyan
        Write-Host "   1. A code will appear in this window" -ForegroundColor White
        Write-Host "   2. A browser will open to: https://microsoft.com/devicelogin" -ForegroundColor White
        Write-Host "   3. Enter the code in the browser" -ForegroundColor White
        Write-Host "   4. Sign in with your M365 account" -ForegroundColor White
        Write-Host "   5. Complete any MFA prompts" -ForegroundColor White
        Write-Host ""
        Write-Host "🔄 Initiating connection..." -ForegroundColor Yellow
        
        # Import Exchange Online Management module if not already loaded
        if (-not (Get-Module -Name ExchangeOnlineManagement)) {
            Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        }
        
        # Connect using Device Code Authentication
        Connect-ExchangeOnline -Device -ShowProgress:$true -ShowBanner:$false -ErrorAction Stop
        
        # Verify connection
        if (Test-ExchangeOnlineConnection) {
            Write-Host ""
            Write-Host "✅ SUCCESS: Exchange Online connected!" -ForegroundColor Green -BackgroundColor DarkGreen
            Write-Host "🎯 Ready to discover Exchange data..." -ForegroundColor Green
            Write-Host ""
            return $true
        }
        else {
            Write-Warning "Exchange Online connection established but cmdlets not available"
            return $false
        }
    }
    catch {
        Write-Host ""
        Write-Host "❌ AUTHENTICATION FAILED" -ForegroundColor Red -BackgroundColor DarkRed
        Write-Warning "Failed to connect to Exchange Online: $($_.Exception.Message)"
        Write-Host ""
        Write-Host "💡 Troubleshooting Tips:" -ForegroundColor Yellow
        Write-Host "   • Make sure you have Exchange Online permissions" -ForegroundColor White
        Write-Host "   • Check your internet connection" -ForegroundColor White
        Write-Host "   • Verify your M365 account is active" -ForegroundColor White
        Write-Host "   • Try again in a few minutes" -ForegroundColor White
        Write-Host ""
        return $false
    }
}

# ================================
# EXCHANGE ONLINE DATA FUNCTIONS
# ================================

function Get-ExchangeMailboxData {
    <#
    .SYNOPSIS
        Gets shared mailboxes from Exchange Online
    .DESCRIPTION
        Retrieves shared mailboxes using Exchange Online PowerShell with fallback to Microsoft Graph
    .OUTPUTS
        Array of shared mailbox objects
    #>
    [CmdletBinding()]
    param()
    
    Write-Verbose "Getting Exchange mailbox data..."
    $SharedMailboxes = @()
    
    try {
        # Try to connect if needed
        $IsConnected = Connect-ExchangeOnlineIfNeeded
        
        if ($IsConnected) {
            Write-Verbose "Exchange Online connected - using Get-EXOMailbox"
            
            # Use Exchange Online PowerShell for accurate data
            $SharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -Properties DisplayName,PrimarySmtpAddress,ArchiveStatus,MailboxPlan -ErrorAction Stop | 
                Select-Object @{
                    Name = 'Name'
                    Expression = { $_.DisplayName }
                }, @{
                    Name = 'EmailAddress'
                    Expression = { $_.PrimarySmtpAddress }
                }, @{
                    Name = 'ArchiveEnabled'
                    Expression = { $_.ArchiveStatus -eq 'Active' }
                }, @{
                    Name = 'MailboxPlan'
                    Expression = { $_.MailboxPlan }
                }, @{
                    Name = 'Type'
                    Expression = { 'SharedMailbox' }
                }, @{
                    Name = 'Source'
                    Expression = { 'ExchangeOnline' }
                }
            
            Write-Verbose "Found $($SharedMailboxes.Count) shared mailboxes via Exchange Online"
        }
        else {
            Write-Warning "Exchange Online not available - falling back to Microsoft Graph"
            
            # Fallback to Microsoft Graph (less accurate)
            if (Get-Command "Get-MgUser" -ErrorAction SilentlyContinue) {
                $GraphUsers = Get-MgUser -Filter "accountEnabled eq false and mail ne null" -Property DisplayName,Mail,Id -All
                
                $SharedMailboxes = $GraphUsers | Where-Object { 
                    $_.Mail -and $_.DisplayName 
                } | Select-Object @{
                    Name = 'Name'
                    Expression = { $_.DisplayName }
                }, @{
                    Name = 'EmailAddress' 
                    Expression = { $_.Mail }
                }, @{
                    Name = 'ArchiveEnabled'
                    Expression = { $false }  # Cannot determine from Graph
                }, @{
                    Name = 'MailboxPlan'
                    Expression = { 'Unknown' }
                }, @{
                    Name = 'Type'
                    Expression = { 'PossibleSharedMailbox' }
                }, @{
                    Name = 'Source'
                    Expression = { 'MicrosoftGraph' }
                }
                
                Write-Warning "Found $($SharedMailboxes.Count) possible shared mailboxes via Microsoft Graph (less accurate)"
            }
            else {
                Write-Warning "Neither Exchange Online nor Microsoft Graph available for mailbox data"
            }
        }
    }
    catch {
        Write-Error "Failed to get shared mailboxes: $($_.Exception.Message)"
        return @()
    }
    
    return $SharedMailboxes
}

function Get-ExchangeDistributionGroupData {
    <#
    .SYNOPSIS
        Gets distribution lists and mail-enabled security groups from Exchange Online
    .DESCRIPTION
        Retrieves distribution groups using Exchange Online PowerShell with fallback to Microsoft Graph
    .OUTPUTS
        Hashtable with DistributionLists and MailEnabledSecurityGroups arrays
    #>
    [CmdletBinding()]
    param()
    
    Write-Verbose "Getting Exchange distribution group data..."
    $Result = @{
        DistributionLists = @()
        MailEnabledSecurityGroups = @()
    }
    
    try {
        # Try to connect if needed
        $IsConnected = Connect-ExchangeOnlineIfNeeded
        
        if ($IsConnected) {
            Write-Verbose "Exchange Online connected - using Get-DistributionGroup"
            
            # Get Distribution Lists
            $DistributionLists = Get-DistributionGroup -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize Unlimited -ErrorAction Stop |
                Select-Object @{
                    Name = 'Name'
                    Expression = { $_.DisplayName }
                }, @{
                    Name = 'EmailAddress'
                    Expression = { $_.PrimarySmtpAddress }
                }, @{
                    Name = 'MemberCount'
                    Expression = { 
                        try { (Get-DistributionGroupMember $_.Identity -ErrorAction SilentlyContinue).Count } 
                        catch { 0 }
                    }
                }, @{
                    Name = 'Type'
                    Expression = { 'DistributionList' }
                }, @{
                    Name = 'Source'
                    Expression = { 'ExchangeOnline' }
                }
            
            # Get Mail-Enabled Security Groups
            $MailEnabledSecurityGroups = Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited -ErrorAction Stop |
                Select-Object @{
                    Name = 'Name'
                    Expression = { $_.DisplayName }
                }, @{
                    Name = 'EmailAddress'
                    Expression = { $_.PrimarySmtpAddress }
                }, @{
                    Name = 'MemberCount'
                    Expression = { 
                        try { (Get-DistributionGroupMember $_.Identity -ErrorAction SilentloContinue).Count } 
                        catch { 0 }
                    }
                }, @{
                    Name = 'Type'
                    Expression = { 'MailEnabledSecurityGroup' }
                }, @{
                    Name = 'Source'
                    Expression = { 'ExchangeOnline' }
                }
            
            $Result.DistributionLists = $DistributionLists
            $Result.MailEnabledSecurityGroups = $MailEnabledSecurityGroups
            
            Write-Verbose "Found $($DistributionLists.Count) distribution lists and $($MailEnabledSecurityGroups.Count) mail-enabled security groups via Exchange Online"
        }
        else {
            Write-Warning "Exchange Online not available - falling back to Microsoft Graph"
            
            # Fallback to Microsoft Graph
            if (Get-Command "Get-MgGroup" -ErrorAction SilentlyContinue) {
                $GraphGroups = Get-MgGroup -Filter "mailEnabled eq true" -Property DisplayName,Mail,Id,GroupTypes -All
                
                foreach ($Group in $GraphGroups) {
                    $GroupObj = @{
                        Name = $Group.DisplayName
                        EmailAddress = $Group.Mail
                        MemberCount = 0  # Cannot easily get from Graph without additional calls
                        Source = 'MicrosoftGraph'
                    }
                    
                    # Distinguish between distribution lists and mail-enabled security groups
                    if ($Group.GroupTypes -contains "Unified") {
                        $GroupObj.Type = 'Microsoft365Group'
                        $Result.DistributionLists += $GroupObj
                    }
                    else {
                        $GroupObj.Type = 'DistributionList'
                        $Result.DistributionLists += $GroupObj
                    }
                }
                
                Write-Warning "Found $($Result.DistributionLists.Count) mail-enabled groups via Microsoft Graph (less detailed)"
            }
            else {
                Write-Warning "Neither Exchange Online nor Microsoft Graph available for group data"
            }
        }
    }
    catch {
        Write-Error "Failed to get distribution groups: $($_.Exception.Message)"
    }
    
    return $Result
}

function Get-ExchangeAcceptedDomains {
    <#
    .SYNOPSIS
        Gets accepted domains from Exchange Online
    .DESCRIPTION
        Retrieves accepted domains using Exchange Online PowerShell
    .OUTPUTS
        Array of accepted domain objects
    #>
    [CmdletBinding()]
    param()
    
    Write-Verbose "Getting Exchange accepted domains..."
    $AcceptedDomains = @()
    
    try {
        # Try to connect if needed
        $IsConnected = Connect-ExchangeOnlineIfNeeded
        
        if ($IsConnected) {
            $AcceptedDomains = Get-AcceptedDomain -ErrorAction Stop | Select-Object @{
                Name = 'DomainName'
                Expression = { $_.DomainName }
            }, @{
                Name = 'DomainType'
                Expression = { $_.DomainType }
            }, @{
                Name = 'Default'
                Expression = { $_.Default }
            }, @{
                Name = 'Source'
                Expression = { 'ExchangeOnline' }
            }
            
            Write-Verbose "Found $($AcceptedDomains.Count) accepted domains"
        }
        else {
            Write-Warning "Exchange Online not connected - cannot get accepted domains"
        }
    }
    catch {
        Write-Error "Failed to get accepted domains: $($_.Exception.Message)"
    }
    
    return $AcceptedDomains
}

function Get-AllExchangeData {
    <#
    .SYNOPSIS
        Gets all Exchange Online data and updates module cache
    .DESCRIPTION
        Retrieves all Exchange data (shared mailboxes, distribution lists, domains) and caches it
        Automatically connects to Exchange Online if needed using Device Code Authentication
    .OUTPUTS
        Hashtable with all Exchange data
    #>
    [CmdletBinding()]
    param()
    
    Write-Host "🔄 Discovering Exchange Online data..." -ForegroundColor Cyan
    
    try {
        # Check Exchange Online connection
        $IsConnected = Test-ExchangeOnlineConnection
        if (-not $IsConnected) {
            Write-Host "   🔗 Exchange Online not connected, attempting connection..." -ForegroundColor Yellow
            $IsConnected = Connect-ExchangeOnlineIfNeeded
        }
        
        if ($IsConnected) {
            Write-Host "   ✅ Exchange Online connected - using enhanced cmdlets" -ForegroundColor Green
        }
        else {
            Write-Host "   ⚠️  Exchange Online not available - using fallback methods" -ForegroundColor Yellow
        }
        
        # Get all Exchange data
        Write-Host "   📧 Getting shared mailboxes..." -ForegroundColor Gray
        $SharedMailboxes = Get-ExchangeMailboxData
        
        Write-Host "   📋 Getting distribution groups..." -ForegroundColor Gray
        $DistributionGroups = Get-ExchangeDistributionGroupData
        
        Write-Host "   🌐 Getting accepted domains..." -ForegroundColor Gray
        $AcceptedDomains = Get-ExchangeAcceptedDomains
        
        # Update module cache
        $Script:ExchangeData.SharedMailboxes = $SharedMailboxes
        $Script:ExchangeData.DistributionLists = $DistributionGroups.DistributionLists
        $Script:ExchangeData.MailEnabledSecurityGroups = $DistributionGroups.MailEnabledSecurityGroups
        $Script:ExchangeData.AcceptedDomains = $AcceptedDomains
        $Script:ExchangeData.LastRefresh = Get-Date
        
        # Return consolidated data
        $AllData = @{
            SharedMailboxes = $SharedMailboxes
            DistributionLists = $DistributionGroups.DistributionLists
            MailEnabledSecurityGroups = $DistributionGroups.MailEnabledSecurityGroups
            AcceptedDomains = $AcceptedDomains
            LastRefresh = $Script:ExchangeData.LastRefresh
            ConnectionStatus = if($IsConnected) { 'Connected' } else { 'Fallback' }
            Summary = @{
                SharedMailboxCount = $SharedMailboxes.Count
                DistributionListCount = $DistributionGroups.DistributionLists.Count
                MailEnabledSecurityGroupCount = $DistributionGroups.MailEnabledSecurityGroups.Count
                AcceptedDomainCount = $AcceptedDomains.Count
            }
        }
        
        Write-Host "   ✅ Exchange data discovery completed!" -ForegroundColor Green
        $StatusMsg = if($IsConnected) { "Enhanced Exchange Online" } else { "Fallback methods" }
        Write-Host "      📊 Found via $StatusMsg : $($AllData.Summary.SharedMailboxCount) shared mailboxes, $($AllData.Summary.DistributionListCount) distribution lists, $($AllData.Summary.MailEnabledSecurityGroupCount) mail-enabled security groups" -ForegroundColor Gray
        
        return $AllData
    }
    catch {
        Write-Error "Failed to get all Exchange data: $($_.Exception.Message)"
        return @{
            SharedMailboxes = @()
            DistributionLists = @()
            MailEnabledSecurityGroups = @()
            AcceptedDomains = @()
            LastRefresh = $null
            ConnectionStatus = 'Failed'
            Summary = @{
                SharedMailboxCount = 0
                DistributionListCount = 0
                MailEnabledSecurityGroupCount = 0
                AcceptedDomainCount = 0
            }
        }
    }
}

# ================================
# USER PROVISIONING FUNCTIONS
# ================================

function Add-UserToSharedMailbox {
    <#
    .SYNOPSIS
        Adds a user to a shared mailbox with specified permissions
    .PARAMETER NewUser
        User object or email address to add
    .PARAMETER SharedMailboxName
        Name or email of the shared mailbox
    .PARAMETER Permission
        Permission level (FullAccess, SendAs, SendOnBehalf)
    .OUTPUTS
        Boolean indicating success
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$NewUser,
        
        [Parameter(Mandatory = $true)]
        [string]$SharedMailboxName,
        
        [Parameter()]
        [ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf')]
        [string]$Permission = 'FullAccess'
    )
    
    try {
        # Get user email
        $UserEmail = if ($NewUser -is [string]) { $NewUser } else { $NewUser.UserPrincipalName }
        
        Write-Verbose "Adding $UserEmail to shared mailbox $SharedMailboxName with $Permission permission"
        
        # Check/connect to Exchange Online
        $IsConnected = Connect-ExchangeOnlineIfNeeded
        
        if ($IsConnected) {
            switch ($Permission) {
                'FullAccess' {
                    Add-MailboxPermission -Identity $SharedMailboxName -User $UserEmail -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
                }
                'SendAs' {
                    Add-RecipientPermission -Identity $SharedMailboxName -Trustee $UserEmail -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                }
                'SendOnBehalf' {
                    Set-Mailbox -Identity $SharedMailboxName -GrantSendOnBehalfTo @{Add=$UserEmail} -ErrorAction Stop
                }
            }
            
            Write-Host "   ✅ Added $UserEmail to shared mailbox $SharedMailboxName ($Permission)" -ForegroundColor Green
            return $true
        }
        else {
            Write-Warning "Exchange Online not connected - cannot add user to shared mailbox"
            return $false
        }
    }
    catch {
        Write-Error "Failed to add user to shared mailbox: $($_.Exception.Message)"
        return $false
    }
}

function Add-UserToDistributionList {
    <#
    .SYNOPSIS
        Adds a user to a distribution list
    .PARAMETER NewUser
        User object or email address to add
    .PARAMETER DistributionListName
        Name or email of the distribution list
    .OUTPUTS
        Boolean indicating success
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$NewUser,
        
        [Parameter(Mandatory = $true)]
        [string]$DistributionListName
    )
    
    try {
        # Get user email
        $UserEmail = if ($NewUser -is [string]) { $NewUser } else { $NewUser.UserPrincipalName }
        
        Write-Verbose "Adding $UserEmail to distribution list $DistributionListName"
        
        # Check/connect to Exchange Online
        $IsConnected = Connect-ExchangeOnlineIfNeeded
        
        if ($IsConnected) {
            Add-DistributionGroupMember -Identity $DistributionListName -Member $UserEmail -Confirm:$false -ErrorAction Stop
            
            Write-Host "   ✅ Added $UserEmail to distribution list $DistributionListName" -ForegroundColor Green
            return $true
        }
        else {
            Write-Warning "Exchange Online not connected - cannot add user to distribution list"
            return $false
        }
    }
    catch {
        Write-Error "Failed to add user to distribution list: $($_.Exception.Message)"
        return $false
    }
}

function Add-UserToMailEnabledSecurityGroup {
    <#
    .SYNOPSIS
        Adds a user to a mail-enabled security group
    .PARAMETER NewUser
        User object or email address to add
    .PARAMETER GroupName
        Name or email of the mail-enabled security group
    .OUTPUTS
        Boolean indicating success
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$NewUser,
        
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )
    
    try {
        # Get user email
        $UserEmail = if ($NewUser -is [string]) { $NewUser } else { $NewUser.UserPrincipalName }
        
        Write-Verbose "Adding $UserEmail to mail-enabled security group $GroupName"
        
        # Check/connect to Exchange Online
        $IsConnected = Connect-ExchangeOnlineIfNeeded
        
        if ($IsConnected) {
            Add-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Confirm:$false -ErrorAction Stop
            
            Write-Host "   ✅ Added $UserEmail to mail-enabled security group $GroupName" -ForegroundColor Green
            return $true
        }
        else {
            Write-Warning "Exchange Online not connected - cannot add user to mail-enabled security group"
            return $false
        }
    }
    catch {
        Write-Error "Failed to add user to mail-enabled security group: $($_.Exception.Message)"
        return $false
    }
}

function Invoke-ExchangeUserProvisioning {
    <#
    .SYNOPSIS
        Provisions a user with Exchange Online resources based on configuration
    .PARAMETER NewUser
        User object containing user details
    .PARAMETER SharedMailboxes
        Array of shared mailbox names to add user to
    .PARAMETER DistributionLists
        Array of distribution list names to add user to
    .PARAMETER MailEnabledSecurityGroups
        Array of mail-enabled security group names to add user to
    .OUTPUTS
        Hashtable with provisioning results
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$NewUser,
        
        [Parameter()]
        [string[]]$SharedMailboxes = @(),
        
        [Parameter()]
        [string[]]$DistributionLists = @(),
        
        [Parameter()]
        [string[]]$MailEnabledSecurityGroups = @()
    )
    
    Write-Host "🔧 Provisioning Exchange resources for $($NewUser.UserPrincipalName)..." -ForegroundColor Cyan
    
    $Results = @{
        SharedMailboxResults = @()
        DistributionListResults = @()
        MailEnabledSecurityGroupResults = @()
        OverallSuccess = $true
    }
    
    # Check Exchange Online connection
    $IsConnected = Connect-ExchangeOnlineIfNeeded
    if (-not $IsConnected) {
        Write-Warning "Exchange Online not available - skipping Exchange provisioning"
        $Results.OverallSuccess = $false
        return $Results
    }
    
    # Add to shared mailboxes
    foreach ($SharedMailbox in $SharedMailboxes) {
        $Result = Add-UserToSharedMailbox -NewUser $NewUser -SharedMailboxName $SharedMailbox
        $Results.SharedMailboxResults += @{
            SharedMailbox = $SharedMailbox
            Success = $Result
        }
        if (-not $Result) { $Results.OverallSuccess = $false }
    }
    
    # Add to distribution lists
    foreach ($DistributionList in $DistributionLists) {
        $Result = Add-UserToDistributionList -NewUser $NewUser -DistributionListName $DistributionList
        $Results.DistributionListResults += @{
            DistributionList = $DistributionList
            Success = $Result
        }
        if (-not $Result) { $Results.OverallSuccess = $false }
    }
    
    # Add to mail-enabled security groups
    foreach ($SecurityGroup in $MailEnabledSecurityGroups) {
        $Result = Add-UserToMailEnabledSecurityGroup -NewUser $NewUser -GroupName $SecurityGroup
        $Results.MailEnabledSecurityGroupResults += @{
            SecurityGroup = $SecurityGroup
            Success = $Result
        }
        if (-not $Result) { $Results.OverallSuccess = $false }
    }
    
    if ($Results.OverallSuccess) {
        Write-Host "   ✅ Exchange provisioning completed successfully!" -ForegroundColor Green
    }
    else {
        Write-Warning "Exchange provisioning completed with some failures - check individual results"
    }
    
    return $Results
}

# ================================
# MODULE EXPORTS
# ================================

# Export all public functions
Export-ModuleMember -Function @(
    'Get-ExchangeMailboxData',
    'Get-ExchangeDistributionGroupData', 
    'Get-ExchangeAcceptedDomains',
    'Add-UserToSharedMailbox',
    'Add-UserToDistributionList',
    'Add-UserToMailEnabledSecurityGroup',
    'Get-AllExchangeData',
    'Invoke-ExchangeUserProvisioning',
    'Test-ExchangeOnlineConnection',
    'Connect-ExchangeOnlineIfNeeded'
)