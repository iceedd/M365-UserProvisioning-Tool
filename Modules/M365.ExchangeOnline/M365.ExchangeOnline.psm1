# M365.ExchangeOnline.psm1 - FIXED VERSION
# Exchange Online operations module for M365 User Provisioning Tool
# Works with your existing M365.Authentication module

<#
.SYNOPSIS
    Exchange Online operations module
.DESCRIPTION
    Handles Exchange Online specific operations like shared mailboxes, distribution lists,
    and mail-enabled security groups. Uses authentication from M365.Authentication module.
.NOTES
    Version: 1.0.1 - Dependency Issue Fixed
    Author: Tom Mortiboys
    
    REMOVED: using module M365.Authentication (was causing loading failures)
    ADDED: Dynamic function calls to M365.Authentication functions
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
        # Check if Exchange Online is connected by testing a cmdlet
        $TestConnection = Get-AcceptedDomain -ResultSize 1 -ErrorAction SilentlyContinue
        
        if ($TestConnection) {
            Write-Verbose "Exchange Online connected - using Get-EXOMailbox"
            
            # Use Exchange Online PowerShell for accurate data
            $SharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -Properties DisplayName,PrimarySmtpAddress,ArchiveStatus,MailboxPlan | 
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
            Write-Warning "Exchange Online not connected - falling back to Microsoft Graph"
            
            # Fallback to Microsoft Graph (less accurate)
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
        # Check if Exchange Online is connected
        $TestConnection = Get-AcceptedDomain -ResultSize 1 -ErrorAction SilentlyContinue
        
        if ($TestConnection) {
            Write-Verbose "Exchange Online connected - using Get-DistributionGroup"
            
            # Get Distribution Lists
            $DistributionLists = Get-DistributionGroup -RecipientTypeDetails DistributionGroup -ResultSize Unlimited |
                Select-Object @{
                    Name = 'Name'
                    Expression = { $_.DisplayName }
                }, @{
                    Name = 'EmailAddress'
                    Expression = { $_.PrimarySmtpAddress }
                }, @{
                    Name = 'MemberCount'
                    Expression = { (Get-DistributionGroupMember $_.Identity).Count }
                }, @{
                    Name = 'Type'
                    Expression = { 'DistributionList' }
                }, @{
                    Name = 'Source'
                    Expression = { 'ExchangeOnline' }
                }
            
            # Get Mail-Enabled Security Groups
            $MailEnabledSecurityGroups = Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited |
                Select-Object @{
                    Name = 'Name'
                    Expression = { $_.DisplayName }
                }, @{
                    Name = 'EmailAddress'
                    Expression = { $_.PrimarySmtpAddress }
                }, @{
                    Name = 'MemberCount'
                    Expression = { (Get-DistributionGroupMember $_.Identity).Count }
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
            Write-Warning "Exchange Online not connected - falling back to Microsoft Graph"
            
            # Fallback to Microsoft Graph
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
        # Check if Exchange Online is connected
        $TestConnection = Get-AcceptedDomain -ResultSize 1 -ErrorAction SilentlyContinue
        
        if ($TestConnection) {
            $AcceptedDomains = Get-AcceptedDomain | Select-Object @{
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
    .OUTPUTS
        Hashtable with all Exchange data
    #>
    [CmdletBinding()]
    param()
    
    Write-Host "🔄 Discovering Exchange Online data..." -ForegroundColor Cyan
    
    try {
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
            Summary = @{
                SharedMailboxCount = $SharedMailboxes.Count
                DistributionListCount = $DistributionGroups.DistributionLists.Count
                MailEnabledSecurityGroupCount = $DistributionGroups.MailEnabledSecurityGroups.Count
                AcceptedDomainCount = $AcceptedDomains.Count
            }
        }
        
        Write-Host "   ✅ Exchange data discovery completed!" -ForegroundColor Green
        Write-Host "      📊 Found: $($AllData.Summary.SharedMailboxCount) shared mailboxes, $($AllData.Summary.DistributionListCount) distribution lists, $($AllData.Summary.MailEnabledSecurityGroupCount) mail-enabled security groups" -ForegroundColor Gray
        
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
        
        # Check if Exchange Online is connected
        $TestConnection = Get-AcceptedDomain -ResultSize 1 -ErrorAction SilentlyContinue
        
        if ($TestConnection) {
            switch ($Permission) {
                'FullAccess' {
                    Add-MailboxPermission -Identity $SharedMailboxName -User $UserEmail -AccessRights FullAccess -InheritanceType All -Confirm:$false
                }
                'SendAs' {
                    Add-RecipientPermission -Identity $SharedMailboxName -Trustee $UserEmail -AccessRights SendAs -Confirm:$false
                }
                'SendOnBehalf' {
                    Set-Mailbox -Identity $SharedMailboxName -GrantSendOnBehalfTo @{Add=$UserEmail}
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
        
        # Check if Exchange Online is connected
        $TestConnection = Get-AcceptedDomain -ResultSize 1 -ErrorAction SilentlyContinue
        
        if ($TestConnection) {
            Add-DistributionGroupMember -Identity $DistributionListName -Member $UserEmail -Confirm:$false
            
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
        
        # Check if Exchange Online is connected
        $TestConnection = Get-AcceptedDomain -ResultSize 1 -ErrorAction SilentlyContinue
        
        if ($TestConnection) {
            Add-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Confirm:$false
            
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
    'Invoke-ExchangeUserProvisioning'
)