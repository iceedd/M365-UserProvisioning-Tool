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
    Requires M365.Authentication module for connection management
    Version: 1.0.0
    Fixed: Removed problematic 'using module' statement
#>

# REMOVED: using module M365.Authentication (this was causing the error)

#region Private Helper Functions

function Test-ExchangeOnlineConnection {
    <#
    .SYNOPSIS
        Tests if Exchange Online is connected and available
    #>
    try {
        # Check connection using multiple methods
        $ConnectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        
        if ($ConnectionInfo -and $ConnectionInfo.State -eq "Connected") {
            Write-Verbose "Exchange Online connection verified via Get-ConnectionInformation"
            return $true
        }
        
        # Alternative test - try a simple Exchange cmdlet
        $null = Get-AcceptedDomain -ResultSize 1 -ErrorAction Stop
        Write-Verbose "Exchange Online connection verified via Get-AcceptedDomain"
        return $true
        
    } catch {
        Write-Verbose "Exchange Online not connected: $($_.Exception.Message)"
        return $false
    }
}

function Write-ExchangeLog {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error")]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $colors = @{
        Info = "Cyan"
        Success = "Green" 
        Warning = "Yellow"
        Error = "Red"
    }
    
    Write-Host "[$timestamp] [Exchange] " -NoNewline -ForegroundColor Gray
    Write-Host $Message -ForegroundColor $colors[$Level]
}

function Connect-ExchangeIfNeeded {
    <#
    .SYNOPSIS
        Ensures Exchange Online connection using M365.Authentication functions
    #>
    
    # Check if already connected
    if (Test-ExchangeOnlineConnection) {
        return $true
    }
    
    # Try to use M365.Authentication functions if available
    $ConnectFunction = Get-Command Connect-ExchangeOnlineIfNeeded -ErrorAction SilentlyContinue
    
    if ($ConnectFunction) {
        Write-ExchangeLog "Using M365.Authentication to connect to Exchange Online..." "Info"
        try {
            $Result = Connect-ExchangeOnlineIfNeeded
            return $Result
        } catch {
            Write-ExchangeLog "M365.Authentication connection failed: $($_.Exception.Message)" "Warning"
            return $false
        }
    } else {
        Write-ExchangeLog "M365.Authentication functions not available - Exchange features will be logged for manual processing" "Warning"
        return $false
    }
}

#endregion

#region Discovery Functions

function Get-ExchangeMailboxData {
    <#
    .SYNOPSIS
        Discovers mailbox data using Exchange Online PowerShell with authentication check
    .OUTPUTS
        Returns hashtable with mailbox data
    #>
    [CmdletBinding()]
    param()
    
    # Ensure Exchange Online connection
    if (-not (Connect-ExchangeIfNeeded)) {
        Write-ExchangeLog "Exchange Online not available. Using fallback method." "Warning"
        return Get-GraphMailboxFallback
    }
    
    Write-ExchangeLog "Discovering mailbox data via Exchange Online PowerShell..." "Info"
    
    try {
        # Get user mailboxes using high-performance EXO cmdlet
        Write-ExchangeLog "Retrieving user mailboxes..." "Info"
        $UserMailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties PrimarySmtpAddress, RecipientTypeDetails
        
        $MailboxData = @{
            UserMailboxes = $UserMailboxes | Select-Object @{
                Name = "Id"
                Expression = { $_.ExternalDirectoryObjectId }
            }, @{
                Name = "DisplayName"
                Expression = { $_.DisplayName }
            }, @{
                Name = "EmailAddress" 
                Expression = { $_.PrimarySmtpAddress }
            }, @{
                Name = "MailboxType"
                Expression = { "User" }
            }, @{
                Name = "RecipientTypeDetails"
                Expression = { $_.RecipientTypeDetails }
            }
        }
        
        Write-ExchangeLog "Found $($MailboxData.UserMailboxes.Count) user mailboxes" "Success"
        
        # Get shared mailboxes
        Write-ExchangeLog "Retrieving shared mailboxes..." "Info"
        $SharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties PrimarySmtpAddress, RecipientTypeDetails
        
        $MailboxData.SharedMailboxes = $SharedMailboxes | Select-Object @{
            Name = "Id"
            Expression = { $_.ExternalDirectoryObjectId }
        }, @{
            Name = "DisplayName"
            Expression = { $_.DisplayName }
        }, @{
            Name = "EmailAddress"
            Expression = { $_.PrimarySmtpAddress }
        }, @{
            Name = "MailboxType"
            Expression = { "Shared" }
        }, @{
            Name = "RecipientTypeDetails"
            Expression = { $_.RecipientTypeDetails }
        }
        
        Write-ExchangeLog "Found $($MailboxData.SharedMailboxes.Count) shared mailboxes" "Success"
        
        return $MailboxData
        
    } catch {
        Write-ExchangeLog "Error retrieving mailbox data: $($_.Exception.Message)" "Error"
        Write-ExchangeLog "Falling back to Graph API method" "Warning"
        return Get-GraphMailboxFallback
    }
}

function Get-ExchangeDistributionGroupData {
    <#
    .SYNOPSIS
        Discovers distribution lists and mail-enabled security groups via Exchange Online
    .OUTPUTS
        Returns hashtable with distribution group data
    #>
    [CmdletBinding()]
    param()
    
    # Ensure Exchange Online connection
    if (-not (Connect-ExchangeIfNeeded)) {
        Write-ExchangeLog "Exchange Online not available. Using fallback method." "Warning"
        return Get-GraphDistributionGroupFallback
    }
    
    Write-ExchangeLog "Discovering distribution groups via Exchange Online PowerShell..." "Info"
    
    try {
        $GroupData = @{}
        
        # Get distribution lists (not security-enabled)
        Write-ExchangeLog "Retrieving distribution lists..." "Info"
        $DistributionGroups = Get-DistributionGroup -RecipientTypeDetails DistributionGroup -ResultSize Unlimited -ErrorAction Stop
        
        $GroupData.DistributionLists = $DistributionGroups | Select-Object @{
            Name = "Id"
            Expression = { $_.ExternalDirectoryObjectId }
        }, @{
            Name = "DisplayName"
            Expression = { $_.DisplayName }
        }, @{
            Name = "PrimarySmtpAddress"
            Expression = { $_.PrimarySmtpAddress }
        }, @{
            Name = "Alias"
            Expression = { $_.Alias }
        }, @{
            Name = "Description"
            Expression = { $_.Description }
        }, @{
            Name = "RecipientTypeDetails"
            Expression = { $_.RecipientTypeDetails }
        }, @{
            Name = "GroupType"
            Expression = { "Distribution List" }
        }
        
        Write-ExchangeLog "Found $($GroupData.DistributionLists.Count) distribution lists" "Success"
        
        # Get mail-enabled security groups
        Write-ExchangeLog "Retrieving mail-enabled security groups..." "Info"
        $MailEnabledSecGroups = Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited -ErrorAction Stop
        
        $GroupData.MailEnabledSecurityGroups = $MailEnabledSecGroups | Select-Object @{
            Name = "Id"
            Expression = { $_.ExternalDirectoryObjectId }
        }, @{
            Name = "DisplayName"
            Expression = { $_.DisplayName }
        }, @{
            Name = "PrimarySmtpAddress"
            Expression = { $_.PrimarySmtpAddress }
        }, @{
            Name = "Alias"
            Expression = { $_.Alias }
        }, @{
            Name = "Description"
            Expression = { $_.Description }
        }, @{
            Name = "RecipientTypeDetails"
            Expression = { $_.RecipientTypeDetails }
        }, @{
            Name = "GroupType"
            Expression = { "Mail-Enabled Security" }
        }
        
        Write-ExchangeLog "Found $($GroupData.MailEnabledSecurityGroups.Count) mail-enabled security groups" "Success"
        
        return $GroupData
        
    } catch {
        Write-ExchangeLog "Error retrieving distribution group data: $($_.Exception.Message)" "Error"
        Write-ExchangeLog "Falling back to Graph API method" "Warning"
        return Get-GraphDistributionGroupFallback
    }
}

function Get-ExchangeAcceptedDomains {
    <#
    .SYNOPSIS
        Gets accepted domains from Exchange Online
    .OUTPUTS
        Returns array of accepted domains
    #>
    [CmdletBinding()]
    param()
    
    if (-not (Connect-ExchangeIfNeeded)) {
        Write-ExchangeLog "Exchange Online not connected - cannot retrieve accepted domains" "Warning"
        return @()
    }
    
    try {
        Write-ExchangeLog "Retrieving accepted domains..." "Info"
        $Domains = Get-AcceptedDomain -ErrorAction Stop
        
        $AcceptedDomains = $Domains | Select-Object @{
            Name = "Id"
            Expression = { $_.Name }
        }, @{
            Name = "DomainName"
            Expression = { $_.DomainName }
        }, @{
            Name = "IsDefault"
            Expression = { $_.Default }
        }, @{
            Name = "DomainType"
            Expression = { $_.DomainType }
        }
        
        Write-ExchangeLog "Found $($AcceptedDomains.Count) accepted domains" "Success"
        return $AcceptedDomains
        
    } catch {
        Write-ExchangeLog "Error retrieving accepted domains: $($_.Exception.Message)" "Error"
        return @()
    }
}

#endregion

#region User Management Functions

function Add-UserToSharedMailbox {
    <#
    .SYNOPSIS
        Adds user to shared mailbox with proper permissions
    .PARAMETER NewUser
        User object to add
    .PARAMETER SharedMailboxName
        Name or email address of shared mailbox
    .PARAMETER PermissionLevel
        Permission level to grant (default: FullAccess)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$NewUser,
        
        [Parameter(Mandatory)]
        [string]$SharedMailboxName,
        
        [ValidateSet("FullAccess", "ReadOnly")]
        [string]$PermissionLevel = "FullAccess"
    )
    
    # Ensure Exchange Online connection
    if (-not (Connect-ExchangeIfNeeded)) {
        Write-ExchangeLog "Exchange Online not connected - cannot add user to shared mailbox" "Warning"
        Write-ExchangeLog "MANUAL TASK: Add $($NewUser.UserPrincipalName) to shared mailbox '$SharedMailboxName' with FullAccess and SendAs permissions" "Info"
        return $false
    }
    
    try {
        Write-ExchangeLog "Adding $($NewUser.DisplayName) to shared mailbox: $SharedMailboxName" "Info"
        
        # Find the shared mailbox by trying different identifiers
        $SharedMailbox = $null
        try {
            $SharedMailbox = Get-EXOMailbox -Identity $SharedMailboxName -RecipientTypeDetails SharedMailbox -ErrorAction Stop
        } catch {
            Write-ExchangeLog "Shared mailbox '$SharedMailboxName' not found" "Error"
            return $false
        }
        
        if ($SharedMailbox) {
            # Add Full Access permission
            Add-MailboxPermission -Identity $SharedMailbox.PrimarySmtpAddress -User $NewUser.UserPrincipalName -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
            Write-ExchangeLog "Added FullAccess permission for $($NewUser.DisplayName)" "Success"
            
            # Add Send As permission
            Add-RecipientPermission -Identity $SharedMailbox.PrimarySmtpAddress -Trustee $NewUser.UserPrincipalName -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            Write-ExchangeLog "Added SendAs permission for $($NewUser.DisplayName)" "Success"
            
            Write-ExchangeLog "Successfully added all permissions for $($NewUser.DisplayName) to shared mailbox $SharedMailboxName" "Success"
            return $true
        }
        
    } catch {
        Write-ExchangeLog "Failed to add user to shared mailbox '$SharedMailboxName': $($_.Exception.Message)" "Error"
        return $false
    }
}

function Add-UserToDistributionList {
    <#
    .SYNOPSIS
        Adds user to distribution list
    .PARAMETER NewUser
        User object to add
    .PARAMETER DistributionListName
        Name or email address of distribution list
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$NewUser,
        
        [Parameter(Mandatory)]
        [string]$DistributionListName
    )
    
    # Ensure Exchange Online connection
    if (-not (Connect-ExchangeIfNeeded)) {
        Write-ExchangeLog "Exchange Online not connected - cannot add user to distribution list" "Warning"
        Write-ExchangeLog "MANUAL TASK: Add $($NewUser.UserPrincipalName) to distribution list '$DistributionListName'" "Info"
        return $false
    }
    
    try {
        Write-ExchangeLog "Adding $($NewUser.DisplayName) to distribution list: $DistributionListName" "Info"
        
        # Find the distribution list
        $DistList = $null
        try {
            $DistList = Get-DistributionGroup -Identity $DistributionListName -ErrorAction Stop
        } catch {
            Write-ExchangeLog "Distribution list '$DistributionListName' not found" "Error"
            return $false
        }
        
        if ($DistList) {
            Add-DistributionGroupMember -Identity $DistList.PrimarySmtpAddress -Member $NewUser.UserPrincipalName -ErrorAction Stop
            Write-ExchangeLog "Successfully added $($NewUser.DisplayName) to distribution list $DistributionListName" "Success"
            return $true
        }
        
    } catch {
        Write-ExchangeLog "Failed to add user to distribution list '$DistributionListName': $($_.Exception.Message)" "Error"
        return $false
    }
}

function Add-UserToMailEnabledSecurityGroup {
    <#
    .SYNOPSIS
        Adds user to mail-enabled security group
    .PARAMETER NewUser
        User object to add
    .PARAMETER GroupName
        Name or email address of mail-enabled security group
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$NewUser,
        
        [Parameter(Mandatory)]
        [string]$GroupName
    )
    
    # Ensure Exchange Online connection
    if (-not (Connect-ExchangeIfNeeded)) {
        Write-ExchangeLog "Exchange Online not connected - cannot add user to mail-enabled security group" "Warning"
        Write-ExchangeLog "MANUAL TASK: Add $($NewUser.UserPrincipalName) to mail-enabled security group '$GroupName'" "Info"
        return $false
    }
    
    try {
        Write-ExchangeLog "Adding $($NewUser.DisplayName) to mail-enabled security group: $GroupName" "Info"
        
        # Find the mail-enabled security group
        $MailEnabledGroup = $null
        try {
            $MailEnabledGroup = Get-DistributionGroup -Identity $GroupName -RecipientTypeDetails MailUniversalSecurityGroup -ErrorAction Stop
        } catch {
            Write-ExchangeLog "Mail-enabled security group '$GroupName' not found" "Error"
            return $false
        }
        
        if ($MailEnabledGroup) {
            Add-DistributionGroupMember -Identity $MailEnabledGroup.PrimarySmtpAddress -Member $NewUser.UserPrincipalName -ErrorAction Stop
            Write-ExchangeLog "Successfully added $($NewUser.DisplayName) to mail-enabled security group $GroupName" "Success"
            return $true
        }
        
    } catch {
        Write-ExchangeLog "Failed to add user to mail-enabled security group '$GroupName': $($_.Exception.Message)" "Error"
        return $false
    }
}

#endregion

#region Fallback Functions (Graph API)

function Get-GraphMailboxFallback {
    <#
    .SYNOPSIS
        Fallback method using Graph API when Exchange Online unavailable
    #>
    Write-ExchangeLog "Using Graph API fallback for mailbox discovery (limited accuracy)" "Warning"
    
    try {
        # Try to get tenant data if M365.Authentication functions are available
        $GetTenantDataFunction = Get-Command Get-M365TenantData -ErrorAction SilentlyContinue
        
        if ($GetTenantDataFunction) {
            $TenantData = Get-M365TenantData
            
            if ($TenantData -and $TenantData.AvailableUsers) {
                $UserMailboxes = $TenantData.AvailableUsers | Where-Object { $null -ne $_.Mail }
                
                # Try to identify shared mailboxes (very limited accuracy)
                $PotentialShared = $TenantData.AvailableUsers | Where-Object { 
                    $_.AccountEnabled -eq $false -and $null -ne $_.Mail 
                }
                
                return @{
                    UserMailboxes = $UserMailboxes | Select-Object @{
                        Name = "Id"
                        Expression = { $_.Id }
                    }, @{
                        Name = "DisplayName"
                        Expression = { $_.DisplayName }
                    }, @{
                        Name = "EmailAddress"
                        Expression = { $_.Mail }
                    }, @{
                        Name = "MailboxType"
                        Expression = { "User" }
                    }
                    
                    SharedMailboxes = $PotentialShared | Select-Object @{
                        Name = "Id"
                        Expression = { $_.Id }
                    }, @{
                        Name = "DisplayName"
                        Expression = { $_.DisplayName }
                    }, @{
                        Name = "EmailAddress"
                        Expression = { $_.Mail }
                    }, @{
                        Name = "MailboxType"
                        Expression = { "Shared (Potential)" }
                    }
                }
            }
        }
        
    } catch {
        Write-ExchangeLog "Graph API fallback failed: $($_.Exception.Message)" "Error"
    }
    
    return @{
        UserMailboxes = @()
        SharedMailboxes = @()
    }
}

function Get-GraphDistributionGroupFallback {
    <#
    .SYNOPSIS
        Fallback method for distribution groups using Graph API
    #>
    Write-ExchangeLog "Using Graph API fallback for distribution groups (limited functionality)" "Warning"
    
    try {
        # Try to get tenant data if M365.Authentication functions are available
        $GetTenantDataFunction = Get-Command Get-M365TenantData -ErrorAction SilentlyContinue
        
        if ($GetTenantDataFunction) {
            $TenantData = Get-M365TenantData
            
            if ($TenantData -and $TenantData.AvailableGroups) {
                # Separate distribution lists and mail-enabled security groups
                $DistributionLists = $TenantData.AvailableGroups | Where-Object { 
                    $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $false 
                }
                
                $MailEnabledSecGroups = $TenantData.AvailableGroups | Where-Object { 
                    $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $true 
                }
                
                return @{
                    DistributionLists = $DistributionLists | Select-Object Id, DisplayName, @{
                        Name = "PrimarySmtpAddress"
                        Expression = { $_.Mail }
                    }, Description, @{
                        Name = "GroupType"
                        Expression = { "Distribution List (Graph)" }
                    }
                    
                    MailEnabledSecurityGroups = $MailEnabledSecGroups | Select-Object Id, DisplayName, @{
                        Name = "PrimarySmtpAddress"
                        Expression = { $_.Mail }
                    }, Description, @{
                        Name = "GroupType"
                        Expression = { "Mail-Enabled Security (Graph)" }
                    }
                }
            }
        }
        
    } catch {
        Write-ExchangeLog "Graph API fallback failed: $($_.Exception.Message)" "Error"
    }
    
    return @{
        DistributionLists = @()
        MailEnabledSecurityGroups = @()
    }
}

#endregion

#region Public Interface Functions

function Get-AllExchangeData {
    <#
    .SYNOPSIS
        Gets comprehensive Exchange data using the best available method
    .OUTPUTS
        Returns hashtable with all Exchange data
    #>
    [CmdletBinding()]
    param()
    
    Write-ExchangeLog "Starting comprehensive Exchange data discovery..." "Info"
    
    # Get mailbox data
    $MailboxData = Get-ExchangeMailboxData
    
    # Get distribution group data
    $GroupData = Get-ExchangeDistributionGroupData
    
    # Get accepted domains
    $AcceptedDomains = Get-ExchangeAcceptedDomains
    
    # Compile results
    $ExchangeData = @{
        UserMailboxes = $MailboxData.UserMailboxes
        SharedMailboxes = $MailboxData.SharedMailboxes
        DistributionLists = $GroupData.DistributionLists
        MailEnabledSecurityGroups = $GroupData.MailEnabledSecurityGroups
        AcceptedDomains = $AcceptedDomains
        ConnectionStatus = @{
            ExchangeOnlineConnected = (Test-ExchangeOnlineConnection)
            DataSource = if (Test-ExchangeOnlineConnection) { "Exchange Online PowerShell" } else { "Graph API Fallback" }
        }
    }
    
    # Log summary
    Write-ExchangeLog "=== Exchange Data Discovery Complete ===" "Success"
    Write-ExchangeLog "User Mailboxes: $($ExchangeData.UserMailboxes.Count)" "Info"
    Write-ExchangeLog "Shared Mailboxes: $($ExchangeData.SharedMailboxes.Count)" "Info"
    Write-ExchangeLog "Distribution Lists: $($ExchangeData.DistributionLists.Count)" "Info"
    Write-ExchangeLog "Mail-Enabled Security Groups: $($ExchangeData.MailEnabledSecurityGroups.Count)" "Info"
    Write-ExchangeLog "Accepted Domains: $($ExchangeData.AcceptedDomains.Count)" "Info"
    Write-ExchangeLog "Data Source: $($ExchangeData.ConnectionStatus.DataSource)" "Info"
    
    return $ExchangeData
}

function Invoke-ExchangeUserProvisioning {
    <#
    .SYNOPSIS
        Main function to handle Exchange-specific user provisioning
    .PARAMETER NewUser
        User object that was created
    .PARAMETER ExchangeAssignments
        Array of Exchange assignments (shared mailboxes, distribution lists, etc.)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$NewUser,
        
        [Parameter(Mandatory)]
        [array]$ExchangeAssignments
    )
    
    Write-ExchangeLog "Starting Exchange provisioning for user: $($NewUser.DisplayName)" "Info"
    
    $Results = @{
        SharedMailboxes = @()
        DistributionLists = @()
        MailEnabledSecurityGroups = @()
        ManualTasks = @()
    }
    
    foreach ($Assignment in $ExchangeAssignments) {
        switch ($Assignment.Type) {
            "SharedMailbox" {
                $Result = Add-UserToSharedMailbox -NewUser $NewUser -SharedMailboxName $Assignment.Name
                $Results.SharedMailboxes += @{
                    Name = $Assignment.Name
                    Success = $Result
                }
            }
            
            "DistributionList" {
                $Result = Add-UserToDistributionList -NewUser $NewUser -DistributionListName $Assignment.Name
                $Results.DistributionLists += @{
                    Name = $Assignment.Name
                    Success = $Result
                }
            }
            
            "MailEnabledSecurityGroup" {
                $Result = Add-UserToMailEnabledSecurityGroup -NewUser $NewUser -GroupName $Assignment.Name
                $Results.MailEnabledSecurityGroups += @{
                    Name = $Assignment.Name
                    Success = $Result
                }
            }
        }
    }
    
    Write-ExchangeLog "Exchange provisioning completed for user: $($NewUser.DisplayName)" "Success"
    return $Results
}

#endregion

# Export public functions
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