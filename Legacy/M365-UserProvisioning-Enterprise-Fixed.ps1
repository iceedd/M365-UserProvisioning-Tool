#Requires -Version 7.0
<#
.SYNOPSIS
    M365 User Provisioning Tool - Enterprise Edition 2025
    Comprehensive M365 user management with intelligent tenant discovery

.DESCRIPTION
    Advanced user provisioning tool with:
    - Intelligent tenant discovery (users, groups, mailboxes, sites)
    - Single user creation and bulk CSV import
    - License assignment via CustomAttribute1
    - UK-based location management
    - Clean tabbed interface with pagination
    - Robust error handling and validation
    - Exchange Online PowerShell integration for distribution lists and shared mailboxes

.NOTES
    Version: 3.0.2025 - REPLICATION FIXED
    Author: Enterprise Solutions Team
    PowerShell: 7.0+ Required
    Dependencies: Microsoft Graph PowerShell SDK V2.28+, Exchange Online PowerShell (optional)
    Last Updated: July 2025
    Fixes: Emoji corruption resolved, Exchange Online connection RESTORED, Azure AD replication delays FIXED

.EXAMPLE
    .\M365-UserProvisioning-Enterprise-Fixed.ps1

.REPLACED
    # Replace ALL $Global: references with module-scoped variables:
    # Before: $Global:IsConnected
    # After: $Script:IsConnected
#>

# ================================
# GLOBAL VARIABLES & CONFIGURATION
# ================================

$Global:IsConnected = $false
$Global:TenantInfo = $null
$Global:AvailableLicenses = @()
$Global:AvailableGroups = @()
$Global:AvailableUsers = @()
$Global:AvailableMailboxes = @()
$Global:DistributionLists = @()
$Global:MailEnabledSecurityGroups = @()
$Global:SharedMailboxes = @()
$Global:SharePointSites = @()
$Global:AcceptedDomains = @()
$Global:CurrentPage = 1
$Global:PageSize = 50
$Global:TotalItems = 0

# FIXED: Added Exchange Online connection tracking
$Global:ExchangeOnlineConnected = $false

# UK-based locations configuration
$Global:UKLocations = @(
    "United Kingdom - London",
    "United Kingdom - Manchester", 
    "United Kingdom - Birmingham",
    "United Kingdom - Leeds",
    "United Kingdom - Glasgow",
    "United Kingdom - Edinburgh",
    "United Kingdom - Bristol",
    "United Kingdom - Liverpool",
    "Remote/Home Working - UK",
    "Office - Head Office",
    "Office - Branch Office"
)

# License type mappings
$Global:LicenseTypes = @{
    "BusinessBasic" = "BusinessBasic"
    "BusinessPremium" = "BusinessPremium"
    "BusinessStandard" = "BusinessStandard"
    "ExchangeOnline1" = "ExchangeOnline1"
    "ExchangeOnline2" = "ExchangeOnline2"
}

# Activity logging
$Global:LogFile = "M365_Provisioning_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$Global:ActivityLog = @()

# ================================
# ASSEMBLY LOADING & INITIALIZATION
# ================================

function Initialize-Application {
    try {
        Write-Host "Initializing M365 User Provisioning Tool..." -ForegroundColor Cyan
        
        # Load Windows Forms assemblies
        Write-Host "Loading Windows Forms assemblies..." -ForegroundColor Yellow
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        
        # Check for Microsoft Graph PowerShell SDK
        Write-Host "Checking Microsoft Graph PowerShell SDK..." -ForegroundColor Yellow
        $GraphModule = Get-Module -Name Microsoft.Graph -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        
        if (-not $GraphModule) {
            $InstallChoice = [System.Windows.Forms.MessageBox]::Show(
                "Microsoft Graph PowerShell SDK is not installed.`n`nWould you like to install it now?`n`nNote: This requires internet connection and may take several minutes.",
                "Missing Dependency",
                "YesNo",
                "Question"
            )
            
            if ($InstallChoice -eq "Yes") {
                Write-Host "Installing Microsoft Graph PowerShell SDK..." -ForegroundColor Yellow
                try {
                    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
                    Write-Host "✓ Microsoft Graph PowerShell SDK installed successfully" -ForegroundColor Green
                }
                catch {
                    throw "Failed to install Microsoft Graph PowerShell SDK: $($_.Exception.Message)"
                }
            }
            else {
                throw "Microsoft Graph PowerShell SDK is required to run this tool."
            }
        }
        else {
            Write-Host "✓ Microsoft Graph PowerShell SDK found (Version: $($GraphModule.Version))" -ForegroundColor Green
            
            if ($GraphModule.Version -lt [Version]"2.23.0") {
                Write-Warning "Microsoft Graph PowerShell SDK V2.23+ recommended for best compatibility"
            }
        }
        
        # Import required modules
        Write-Host "Importing Microsoft Graph modules..." -ForegroundColor Yellow
        Import-Module Microsoft.Graph.Authentication -Force
        Import-Module Microsoft.Graph.Users -Force
        Import-Module Microsoft.Graph.Groups -Force
        Import-Module Microsoft.Graph.Identity.DirectoryManagement -Force
        
        Write-Host "✓ Application initialized successfully" -ForegroundColor Green
        return $true
    }
    catch {
        $ErrorMsg = "Failed to initialize application: $($_.Exception.Message)"
        Write-Host $ErrorMsg -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show($ErrorMsg, "Initialization Error", "OK", "Error")
        return $false
    }
}

# ================================
# LOGGING FUNCTIONS
# ================================

function Write-ActivityLog {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    $Global:ActivityLog += $LogEntry
    
    try {
        $LogEntry | Out-File -FilePath $Global:LogFile -Append -Encoding UTF8
    }
    catch {
        # Silently continue if file logging fails
    }
    
    switch ($Level) {
        "Info" { Write-Host $LogEntry -ForegroundColor White }
        "Warning" { Write-Host $LogEntry -ForegroundColor Yellow }
        "Error" { Write-Host $LogEntry -ForegroundColor Red }
        "Success" { Write-Host $LogEntry -ForegroundColor Green }
    }
}

# ================================
# EXCHANGE ONLINE HELPER FUNCTIONS - FIXED
# ================================

function Get-CleanGroupName {
    param([string]$GroupDisplayText)
    
    # Extract clean group name from display text (no emojis, no brackets)
    if ($GroupDisplayText -match '^([^[\]]+)\s*\[') {
        return $Matches[1].Trim()
    }
    else {
        # Fallback - split on brackets and take first part
        return ($GroupDisplayText -split '\[')[0].Trim()
    }
}

function Test-ExchangeOnlineModule {
    try {
        Write-ActivityLog "Checking for Exchange Online PowerShell module..." "Info"
        $ExOModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement
        
        if (-not $ExOModule) {
            Write-ActivityLog "Exchange Online PowerShell module not found" "Warning"
            
            $InstallChoice = [System.Windows.Forms.MessageBox]::Show(
                "The Exchange Online PowerShell module is required for shared mailboxes and distribution lists.`n`nInstall now? (Requires admin privileges)",
                "Exchange Online Module Required",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($InstallChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
                Write-ActivityLog "Installing Exchange Online PowerShell module..." "Info"
                Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                Write-ActivityLog "Exchange Online module installed successfully" "Success"
                return $true
            }
            else {
                Write-ActivityLog "User declined to install Exchange Online module" "Info"
                return $false
            }
        }
        else {
            Write-ActivityLog "Exchange Online PowerShell module found (Version: $($ExOModule.Version))" "Success"
            return $true
        }
    }
    catch {
        Write-ActivityLog "Error with Exchange Online module: $($_.Exception.Message)" "Warning"
        return $false
    }
}

function Connect-ExchangeOnlineAtStartup {
    <#
    .SYNOPSIS
        Prompts user whether to connect to Exchange Online for advanced features
    #>
    
    if (Test-ExchangeOnlineModule) {
        # Ask user if they want to connect to Exchange Online now
        $ExchangeChoice = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to connect to Exchange Online now?`n`nThis enables:`n• Distribution list management`n• Shared mailbox permissions`n• Advanced Exchange features`n`nNote: This will open a browser for authentication (same method as Graph)`n`nYou can skip this and operations will be logged for manual processing.",
            "Exchange Online Connection",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($ExchangeChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Write-ActivityLog "User chose to connect to Exchange Online..." "Info"
                Import-Module ExchangeOnlineManagement -Force
                
                Write-ActivityLog "Opening browser for Exchange Online authentication..." "Info"
                
                # Use device code authentication (browser-based, same method as Graph)
                Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
                
                $Global:ExchangeOnlineConnected = $true
                Write-ActivityLog "Successfully connected to Exchange Online" "Success"
                
            }
            catch {
                Write-ActivityLog "Exchange Online connection failed: $($_.Exception.Message)" "Warning"
                Write-ActivityLog "Distribution lists and shared mailboxes will be logged for manual processing" "Info"
                $Global:ExchangeOnlineConnected = $false
            }
        }
        else {
            Write-ActivityLog "User chose to skip Exchange Online connection" "Info"
            Write-ActivityLog "Distribution lists and shared mailboxes will be logged for manual processing" "Info"
            $Global:ExchangeOnlineConnected = $false
        }
    }
    else {
        Write-ActivityLog "Exchange Online PowerShell module not available" "Info"
        Write-ActivityLog "Distribution lists and shared mailboxes will be logged for manual processing" "Info"
        $Global:ExchangeOnlineConnected = $false
    }
}

function Connect-ExchangeOnlineIfNeeded {
    # Check if Exchange Online is already connected
    if ($Global:ExchangeOnlineConnected) {
        return $true
    }
    else {
        # Try to connect now if not already connected
        Write-ActivityLog "Exchange Online not connected, attempting to connect..." "Info"
        
        if (Test-ExchangeOnlineModule) {
            try {
                Write-ActivityLog "Connecting to Exchange Online for distribution list operations..." "Info"
                Import-Module ExchangeOnlineManagement -Force
                
                # Use device code authentication (browser-based, same method as Graph)
                Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
                
                $Global:ExchangeOnlineConnected = $true
                Write-ActivityLog "Successfully connected to Exchange Online" "Success"
                return $true
                
            }
            catch {
                Write-ActivityLog "Exchange Online connection failed: $($_.Exception.Message)" "Warning"
                Write-ActivityLog "Distribution lists and shared mailboxes will be logged for manual processing" "Info"
                $Global:ExchangeOnlineConnected = $false
                return $false
            }
        }
        else {
            Write-ActivityLog "Exchange Online PowerShell module not available" "Warning"
            Write-ActivityLog "Operations will be logged for manual processing" "Info"
            return $false
        }
    }
}

function Add-UserToDistributionList {
    param(
        [object]$NewUser, 
        [string]$CleanName, 
        [string]$GroupDisplayText = "",
        [bool]$IsSecondPass = $false
    )
    
    Write-ActivityLog "Attempting to add to distribution list: $CleanName" "Info"
    
    if (Connect-ExchangeOnlineIfNeeded) {
        # Adjust retry logic based on whether this is second pass
        $MaxRetries = if ($IsSecondPass) { 2 } else { 3 }
        $RetryDelay = if ($IsSecondPass) { 15 } else { 30 } # Shorter delay on second pass
        
        for ($Attempt = 1; $Attempt -le $MaxRetries; $Attempt++) {
            try {
                if ($Attempt -gt 1) {
                    $PassText = if ($IsSecondPass) { "(Second Pass) " } else { "" }
                    Write-ActivityLog "${PassText}Retry attempt $Attempt of $MaxRetries for distribution list: $CleanName" "Info"
                    Write-ActivityLog "Waiting $RetryDelay seconds for Azure AD replication..." "Info"
                    Start-Sleep -Seconds $RetryDelay
                }
                
                $PassText = if ($IsSecondPass) { "(Second Pass) " } else { "" }
                Write-ActivityLog "${PassText}Adding to distribution list via Exchange Online: $CleanName (Attempt $Attempt)" "Info"
                Add-DistributionGroupMember -Identity $CleanName -Member $NewUser.UserPrincipalName -ErrorAction Stop
                Write-ActivityLog "Successfully added to distribution list: $CleanName" "Success"
                return # Success - exit the retry loop
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                
                # Check if it's a "user not found" error (replication delay)
                if ($ErrorMessage -like "*Couldn't find object*" -or $ErrorMessage -like "*wasn't found*") {
                    if ($Attempt -lt $MaxRetries) {
                        Write-ActivityLog "User not found in Exchange Online (replication delay). Will retry in $RetryDelay seconds..." "Warning"
                        continue # Try again
                    }
                    else {
                        if (-not $IsSecondPass) {
                            # Add to failed operations for second pass
                            Write-ActivityLog "Adding to second pass retry list: Distribution List '$CleanName'" "Info"
                            $Script:FailedExchangeOperations += @{
                                Type = "DistributionList"
                                CleanName = $CleanName
                                GroupDisplayText = $GroupDisplayText
                            }
                        }
                        else {
                            Write-ActivityLog "Second pass also failed for distribution list '$CleanName'. Logging for manual processing." "Warning"
                            Write-ActivityLog "Manual task: Add $($NewUser.UserPrincipalName) to distribution list '$CleanName' (allow 5-10 minutes for replication)" "Info"
                        }
                    }
                }
                else {
                    # Different error - don't retry
                    Write-ActivityLog "Failed to add to distribution list '$CleanName': $ErrorMessage" "Warning"
                    Write-ActivityLog "Manual task: Add $($NewUser.UserPrincipalName) to distribution list '$CleanName'" "Info"
                    return
                }
            }
        }
    }
    else {
        Write-ActivityLog "Manual task: Add $($NewUser.UserPrincipalName) to distribution list '$CleanName'" "Info"
    }
}

function Add-UserToSharedMailbox {
    param(
        [object]$NewUser, 
        [string]$CleanName, 
        [string]$GroupDisplayText = "",
        [bool]$IsSecondPass = $false
    )
    
    $SharedMailbox = $Global:SharedMailboxes | Where-Object { 
        $_.DisplayName -eq $CleanName -or $_.DisplayName -like "*$CleanName*" 
    }
    
    if ($SharedMailbox -and (Connect-ExchangeOnlineIfNeeded)) {
        # Adjust retry logic based on whether this is second pass
        $MaxRetries = if ($IsSecondPass) { 2 } else { 3 }
        $RetryDelay = if ($IsSecondPass) { 15 } else { 30 } # Shorter delay on second pass
        
        for ($Attempt = 1; $Attempt -le $MaxRetries; $Attempt++) {
            try {
                if ($Attempt -gt 1) {
                    $PassText = if ($IsSecondPass) { "(Second Pass) " } else { "" }
                    Write-ActivityLog "${PassText}Retry attempt $Attempt of $MaxRetries for shared mailbox: $CleanName" "Info"
                    Write-ActivityLog "Waiting $RetryDelay seconds for Azure AD replication..." "Info"
                    Start-Sleep -Seconds $RetryDelay
                }
                
                $PassText = if ($IsSecondPass) { "(Second Pass) " } else { "" }
                Write-ActivityLog "${PassText}Granting shared mailbox permissions: $CleanName (Attempt $Attempt)" "Info"
                
                # Grant Full Access
                Add-MailboxPermission -Identity $SharedMailbox.EmailAddress -User $NewUser.UserPrincipalName -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
                
                # Grant Send As
                Add-RecipientPermission -Identity $SharedMailbox.EmailAddress -Trustee $NewUser.UserPrincipalName -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                
                Write-ActivityLog "Successfully granted shared mailbox permissions: $CleanName" "Success"
                return # Success - exit the retry loop
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                
                # Check if it's a "user not found" error (replication delay)
                if ($ErrorMessage -like "*wasn't found*" -or $ErrorMessage -like "*Couldn't find object*") {
                    if ($Attempt -lt $MaxRetries) {
                        Write-ActivityLog "User not found in Exchange Online (replication delay). Will retry in $RetryDelay seconds..." "Warning"
                        continue # Try again
                    }
                    else {
                        if (-not $IsSecondPass) {
                            # Add to failed operations for second pass
                            Write-ActivityLog "Adding to second pass retry list: Shared Mailbox '$CleanName'" "Info"
                            $Script:FailedExchangeOperations += @{
                                Type = "SharedMailbox"
                                CleanName = $CleanName
                                GroupDisplayText = $GroupDisplayText
                            }
                        }
                        else {
                            Write-ActivityLog "Second pass also failed for shared mailbox '$CleanName'. Logging for manual processing." "Warning"
                            Write-ActivityLog "Manual task: Grant $($NewUser.UserPrincipalName) access to shared mailbox $($SharedMailbox.EmailAddress) (allow 5-10 minutes for replication)" "Info"
                        }
                    }
                }
                else {
                    # Different error - don't retry
                    Write-ActivityLog "Failed shared mailbox permissions: $ErrorMessage" "Warning"
                    Write-ActivityLog "Manual task: Grant $($NewUser.UserPrincipalName) access to shared mailbox $($SharedMailbox.EmailAddress)" "Info"
                    return
                }
            }
        }
    }
    else {
        Write-ActivityLog "Manual task: Grant $($NewUser.UserPrincipalName) access to shared mailbox with name '$CleanName'" "Info"
    }
}

function Process-GroupMemberships {
    param([object]$NewUser, [array]$Groups)
    
    # Track failed Exchange operations for second pass
    $Script:FailedExchangeOperations = @()
    
    foreach ($GroupDisplayText in $Groups) {
        Write-ActivityLog "Processing group: $GroupDisplayText" "Info"
        
        # Skip separator lines
        if ($GroupDisplayText -like "*---*" -or $GroupDisplayText -like "*───*") {
            continue
        }
        
        # Get clean group name without formatting
        $CleanGroupName = Get-CleanGroupName -GroupDisplayText $GroupDisplayText
        Write-ActivityLog "Clean group name: '$CleanGroupName'" "Info"
        
        # Check if it's a shared mailbox
        if ($GroupDisplayText -match 'Shared Mailbox') {
            Add-UserToSharedMailbox -NewUser $NewUser -CleanName $CleanGroupName -GroupDisplayText $GroupDisplayText
            continue
        }
        
        # Check if it's a distribution list
        if ($GroupDisplayText -match 'Distribution List') {
            Add-UserToDistributionList -NewUser $NewUser -CleanName $CleanGroupName -GroupDisplayText $GroupDisplayText
            continue
        }
        
        # Handle regular groups and mail-enabled security groups
        $Group = $Global:AvailableGroups | Where-Object { 
            $_.DisplayName -eq $CleanGroupName 
        }
        
        if ($Group) {
            try {
                Write-ActivityLog "Adding to $($Group.GroupType): $CleanGroupName" "Info"
                
                $GroupMember = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($NewUser.Id)"
                }
                
                New-MgGroupMember -GroupId $Group.Id -BodyParameter $GroupMember -ErrorAction Stop
                Write-ActivityLog "Successfully added to $($Group.GroupType): $CleanGroupName" "Success"
            }
            catch {
                Write-ActivityLog "Failed to add to group '$CleanGroupName': $($_.Exception.Message)" "Error"
                
                # If it's a mail-enabled group, try distribution list method
                if ($Group.MailEnabled -and $Group.GroupType -eq 'Distribution List') {
                    Add-UserToDistributionList -NewUser $NewUser -CleanName $CleanGroupName -GroupDisplayText $GroupDisplayText
                }
            }
        }
        else {
            Write-ActivityLog "Group '$CleanGroupName' not found in tenant" "Warning"
        }
    }
    
    # Second pass: Retry failed Exchange operations
    if ($Script:FailedExchangeOperations.Count -gt 0) {
        Write-ActivityLog "Starting second pass for $($Script:FailedExchangeOperations.Count) failed Exchange operations..." "Info"
        Write-ActivityLog "Additional wait time should allow for complete replication..." "Info"
        
        foreach ($FailedOperation in $Script:FailedExchangeOperations) {
            Write-ActivityLog "Second pass: Retrying $($FailedOperation.Type) - $($FailedOperation.CleanName)" "Info"
            
            if ($FailedOperation.Type -eq "DistributionList") {
                Add-UserToDistributionList -NewUser $NewUser -CleanName $FailedOperation.CleanName -GroupDisplayText $FailedOperation.GroupDisplayText -IsSecondPass $true
            }
            elseif ($FailedOperation.Type -eq "SharedMailbox") {
                Add-UserToSharedMailbox -NewUser $NewUser -CleanName $FailedOperation.CleanName -GroupDisplayText $FailedOperation.GroupDisplayText -IsSecondPass $true
            }
        }
        
        if ($Script:FailedExchangeOperations.Count -gt 0) {
            Write-ActivityLog "Second pass completed. Some operations may still require manual processing." "Info"
        }
    }
}

# ================================
# MICROSOFT GRAPH CONNECTION
# ================================

function Connect-ToMicrosoftGraph {
    try {
        Write-ActivityLog "Initiating connection to Microsoft Graph..." "Info"
        Update-StatusLabel "Connecting to Microsoft Graph..."
        
        $RequiredScopes = @(
            "User.ReadWrite.All",
            "Directory.ReadWrite.All", 
            "Group.ReadWrite.All",
            "Organization.Read.All",
            "Domain.Read.All",
            "Sites.Read.All",
            "Mail.Read",
            "MailboxSettings.ReadWrite"
        )
        
        Write-ActivityLog "Requesting scopes: $($RequiredScopes -join ', ')" "Info"
        
        Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
        
        $Context = Get-MgContext -ErrorAction Stop
        
        if ($Context -and $Context.TenantId) {
            $Global:IsConnected = $true
            $Global:TenantInfo = $Context
            
            Write-ActivityLog "Successfully connected to Microsoft Graph" "Success"
            Write-ActivityLog "Tenant ID: $($Context.TenantId)" "Info"
            Write-ActivityLog "Account: $($Context.Account)" "Info"
            Write-ActivityLog "Environment: $($Context.Environment)" "Info"
            
            Update-StatusLabel "Discovering tenant resources..."
            Start-TenantDiscovery
            
            # FIXED: Call Exchange Online connection during startup
            Write-ActivityLog "Checking Exchange Online connectivity for advanced features..." "Info"
            Connect-ExchangeOnlineAtStartup
            
            Update-UIAfterConnection
            
            $ExchangeStatus = if ($Global:ExchangeOnlineConnected) { "Connected" } else { "Manual Tasks Only" }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Successfully connected to Microsoft Graph!`n`nTenant: $($Context.TenantId)`nAccount: $($Context.Account)`n`nTenant discovery completed:`n• Licenses: $($Global:AvailableLicenses.Count)`n• Groups: $($Global:AvailableGroups.Count) (including $($Global:DistributionLists.Count) distribution lists)`n• Mail-Enabled Security Groups: $($Global:MailEnabledSecurityGroups.Count)`n• Users: $($Global:AvailableUsers.Count)`n• Shared Mailboxes: $($Global:SharedMailboxes.Count)`n• Domains: $($Global:AcceptedDomains.Count)`n`nExchange Online: $ExchangeStatus",
                "Connection Successful",
                "OK",
                "Information"
            )
            
            return $true
        }
        else {
            throw "Connection established but context is invalid"
        }
    }
    catch {
        $ErrorMsg = "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        Write-ActivityLog $ErrorMsg "Error"
        Update-StatusLabel "Connection failed - $($_.Exception.Message)"
        
        [System.Windows.Forms.MessageBox]::Show(
            $ErrorMsg + "`n`nPlease ensure:`n• You have appropriate permissions`n• Your account has Microsoft Graph access`n• Network connectivity is available`n• You consent to the requested permissions",
            "Connection Failed",
            "OK",
            "Error"
        )
        
        return $false
    }
}

function Start-TenantDiscovery {
    try {
        Write-ActivityLog "Starting comprehensive tenant discovery..." "Info"
        
        # Discover available licenses
        Write-ActivityLog "Discovering available licenses..." "Info"
        $Licenses = Get-MgSubscribedSku -ErrorAction Stop
        $Global:AvailableLicenses = $Licenses | Select-Object SkuId, SkuPartNumber, DisplayName, @{
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
        Write-ActivityLog "Found $($Global:AvailableLicenses.Count) license types" "Success"
        
        # Discover all groups (including mail-enabled security groups)
        Write-ActivityLog "Discovering all groups..." "Info"
        $AllGroups = Get-MgGroup -All -ErrorAction Stop
        $Global:AvailableGroups = $AllGroups | Select-Object Id, DisplayName, Description, @{
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
        Write-ActivityLog "Found $($Global:AvailableGroups.Count) groups total" "Success"
        
        # Separate distribution lists for easier access
        Write-ActivityLog "Identifying distribution lists and mail-enabled groups..." "Info"
        $Global:DistributionLists = $AllGroups | Where-Object { 
            $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $false 
        } | Select-Object Id, DisplayName, Mail, Description
        Write-ActivityLog "Found $($Global:DistributionLists.Count) distribution lists" "Success"
        
        $Global:MailEnabledSecurityGroups = $AllGroups | Where-Object { 
            $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $true 
        } | Select-Object Id, DisplayName, Mail, Description
        Write-ActivityLog "Found $($Global:MailEnabledSecurityGroups.Count) mail-enabled security groups" "Success"
        
        # Discover all users
        Write-ActivityLog "Discovering users..." "Info"
        $Users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Mail,JobTitle,Department -ErrorAction Stop
        $Global:AvailableUsers = $Users | Select-Object Id, DisplayName, UserPrincipalName, Mail, JobTitle, Department
        Write-ActivityLog "Found $($Global:AvailableUsers.Count) users" "Success"
        
        # Discover accepted domains
        Write-ActivityLog "Discovering accepted domains..." "Info"
        $Domains = Get-MgDomain -ErrorAction Stop
        $Global:AcceptedDomains = $Domains | Where-Object { $_.IsVerified -eq $true } | Select-Object Id, @{
            Name = "DomainName"
            Expression = { $_.Id }
        }, IsDefault, IsVerified
        Write-ActivityLog "Found $($Global:AcceptedDomains.Count) verified domains" "Success"
        
        # Discover mailboxes and shared mailboxes
        Write-ActivityLog "Discovering mailboxes..." "Info"
        try {
            # Get all user mailboxes
            $UserMailboxes = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Mail,MailNickname | Where-Object { $null -ne $_.Mail }
            
            $Global:AvailableMailboxes = $UserMailboxes | Select-Object @{
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
            
            Write-ActivityLog "Found $($Global:AvailableMailboxes.Count) user mailboxes" "Success"
            
            # Try to discover shared mailboxes using Graph API
            Write-ActivityLog "Attempting to discover shared mailboxes..." "Info"
            try {
                # Note: This requires Exchange Online admin permissions
                # Shared mailboxes are users with specific recipient type
                $SharedMailboxQuery = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Mail,AccountEnabled | Where-Object { 
                    $_.AccountEnabled -eq $false -and $null -ne $_.Mail 
                }
                
                $Global:SharedMailboxes = $SharedMailboxQuery | Select-Object @{
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
                    Expression = { "Shared" }
                }
                
                Write-ActivityLog "Found $($Global:SharedMailboxes.Count) potential shared mailboxes" "Success"
            }
            catch {
                Write-ActivityLog "Could not discover shared mailboxes via Graph API: $($_.Exception.Message)" "Warning"
                Write-ActivityLog "Note: Shared mailbox discovery may require Exchange Online PowerShell or different permissions" "Info"
                $Global:SharedMailboxes = @()
            }
        }
        catch {
            Write-ActivityLog "Mailbox discovery completed with limited data due to permissions" "Warning"
            $Global:AvailableMailboxes = @()
            $Global:SharedMailboxes = @()
        }
        
        # Discover SharePoint sites
        Write-ActivityLog "Discovering SharePoint sites..." "Info"
        try {
            $Sites = Get-MgSite -All -ErrorAction Stop
            $Global:SharePointSites = $Sites | Select-Object Id, DisplayName, WebUrl, Description
            Write-ActivityLog "Found $($Global:SharePointSites.Count) SharePoint sites" "Success"
        }
        catch {
            Write-ActivityLog "SharePoint site discovery requires additional permissions" "Warning"
            $Global:SharePointSites = @()
        }
        
        Write-ActivityLog "Tenant discovery completed successfully" "Success"
        Write-ActivityLog "Summary: $($Global:AvailableUsers.Count) users, $($Global:AvailableGroups.Count) groups, $($Global:DistributionLists.Count) distribution lists, $($Global:MailEnabledSecurityGroups.Count) mail-enabled security groups, $($Global:SharedMailboxes.Count) shared mailboxes" "Info"
        
    }
    catch {
        Write-ActivityLog "Error during tenant discovery: $($_.Exception.Message)" "Error"
        throw
    }
}

function Disconnect-FromMicrosoftGraph {
    try {
        # FIXED: Disconnect Exchange Online if connected
        if ($Global:ExchangeOnlineConnected) {
            try {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                $Global:ExchangeOnlineConnected = $false
                Write-ActivityLog "Disconnected from Exchange Online" "Info"
            }
            catch {
                Write-ActivityLog "Note: Exchange Online disconnection may have failed" "Warning"
            }
        }
        
        if ($Global:IsConnected) {
            Disconnect-MgGraph -ErrorAction Stop
            $Global:IsConnected = $false
            $Global:TenantInfo = $null
            
            # Clear cached data
            $Global:AvailableLicenses = @()
            $Global:AvailableGroups = @()
            $Global:AvailableUsers = @()
            $Global:AvailableMailboxes = @()
            $Global:DistributionLists = @()
            $Global:MailEnabledSecurityGroups = @()
            $Global:SharedMailboxes = @()
            $Global:SharePointSites = @()
            $Global:AcceptedDomains = @()
            
            Write-ActivityLog "Disconnected from Microsoft Graph" "Info"
            Update-StatusLabel "Disconnected from Microsoft Graph"
            Update-UIAfterDisconnection
            
            [System.Windows.Forms.MessageBox]::Show(
                "Successfully disconnected from Microsoft Graph.",
                "Disconnected",
                "OK",
                "Information"
            )
        }
    }
    catch {
        Write-ActivityLog "Error during disconnection: $($_.Exception.Message)" "Error"
    }
}

# ================================
# USER MANAGEMENT FUNCTIONS
# ================================

function New-M365User {
    param(
        [string]$DisplayName,
        [string]$UserPrincipalName,
        [string]$FirstName,
        [string]$LastName,
        [string]$Department,
        [string]$JobTitle,
        [string]$Office,
        [string]$Manager,
        [string]$LicenseType,
        [string[]]$Groups,
        [string]$Password,
        [bool]$ForcePasswordChange = $true
    )
    
    try {
        Write-ActivityLog "Creating user: $DisplayName ($UserPrincipalName)" "Info"
        
        # Validate UPN domain
        $Domain = $UserPrincipalName.Split('@')[1]
        if ($Domain -notin $Global:AcceptedDomains.DomainName) {
            throw "Domain '$Domain' is not a verified domain in this tenant"
        }
        
        # Check if user already exists
        try {
            $ExistingUser = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
            if ($ExistingUser) {
                throw "User with UPN '$UserPrincipalName' already exists"
            }
        }
        catch {
            # User doesn't exist - continue with creation
        }
        
        # Prepare password profile
        $PasswordProfile = @{
            Password = $Password
            ForceChangePasswordNextSignIn = $ForcePasswordChange
        }
        
        # Prepare user parameters
        $UserParams = @{
            DisplayName = $DisplayName
            UserPrincipalName = $UserPrincipalName
            AccountEnabled = $true
            PasswordProfile = $PasswordProfile
            MailNickname = $UserPrincipalName.Split('@')[0]
            UsageLocation = "GB"
        }
        
        # Add optional properties
        if ($FirstName) { $UserParams.GivenName = $FirstName }
        if ($LastName) { $UserParams.Surname = $LastName }
        if ($Department) { $UserParams.Department = $Department }
        if ($JobTitle) { $UserParams.JobTitle = $JobTitle }
        if ($Office) { $UserParams.OfficeLocation = $Office }
        
        # Create the user
        Write-ActivityLog "Creating user account..." "Info"
        $NewUser = New-MgUser @UserParams -ErrorAction Stop
        Write-ActivityLog "User account created successfully with ID: $($NewUser.Id)" "Success"
        
        # Set CustomAttribute1 for licensing
        if ($LicenseType) {
            Write-ActivityLog "Setting CustomAttribute1 to '$LicenseType' for licensing automation" "Info"
            try {
                # Correct way to set CustomAttribute1 in Microsoft Graph
                $ExtensionAttributes = @{
                    "onPremisesExtensionAttributes" = @{
                        "extensionAttribute1" = $LicenseType
                    }
                }
                Update-MgUser -UserId $NewUser.Id -BodyParameter $ExtensionAttributes -ErrorAction Stop
                
                Write-ActivityLog "CustomAttribute1 set successfully to '$LicenseType'" "Success"
            }
            catch {
                Write-ActivityLog "Warning: Could not set CustomAttribute1. Trying alternative method..." "Warning"
                try {
                    # Alternative method using AdditionalProperties
                    Update-MgUser -UserId $NewUser.Id -AdditionalProperties @{
                        "extensionAttribute1" = $LicenseType
                    } -ErrorAction Stop
                    Write-ActivityLog "CustomAttribute1 set via alternative method" "Success"
                }
                catch {
                    Write-ActivityLog "Warning: Could not set CustomAttribute1. License assignment may need to be done manually: $($_.Exception.Message)" "Warning"
                }
            }
        }
        
        # Set manager if provided
        if ($Manager) {
            try {
                Write-ActivityLog "Setting manager: $Manager" "Info"
                $ManagerUser = Get-MgUser -Filter "userPrincipalName eq '$Manager' or displayName eq '$Manager'" -ErrorAction Stop
                if ($ManagerUser) {
                    $ManagerRef = @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($ManagerUser.Id)"
                    }
                    Set-MgUserManagerByRef -UserId $NewUser.Id -BodyParameter $ManagerRef -ErrorAction Stop
                    Write-ActivityLog "Manager set successfully" "Success"
                }
                else {
                    Write-ActivityLog "Warning: Manager '$Manager' not found" "Warning"
                }
            }
            catch {
                Write-ActivityLog "Warning: Could not set manager: $($_.Exception.Message)" "Warning"
            }
        }
        
        # FIXED: Process group memberships using the new improved function
        if ($Groups -and $Groups.Count -gt 0) {
            Write-ActivityLog "Processing $($Groups.Count) group memberships..." "Info"
            
            # Initialize failed operations tracking
            $Script:FailedExchangeOperations = @()
            
            # Check if we have Exchange Online operations (distribution lists or shared mailboxes)
            $HasExchangeOperations = $Groups | Where-Object { 
                $_ -match 'Distribution List' -or $_ -match 'Shared Mailbox' 
            }
            
            if ($HasExchangeOperations) {
                Write-ActivityLog "Exchange Online operations detected. Allowing time for Azure AD replication..." "Info"
                Write-ActivityLog "Waiting 15 seconds for initial replication before Exchange operations..." "Info"
                Start-Sleep -Seconds 15
            }
            
            Process-GroupMemberships -NewUser $NewUser -Groups $Groups
        }
        
        Write-ActivityLog "User creation completed successfully for: $DisplayName" "Success"
        return $NewUser
    }
    catch {
        $ErrorMsg = "Failed to create user '$DisplayName': $($_.Exception.Message)"
        Write-ActivityLog $ErrorMsg "Error"
        throw $ErrorMsg
    }
}

function Import-UsersFromCSV {
    param(
        [string]$CSVPath
    )
    
    try {
        Write-ActivityLog "Starting CSV import from: $CSVPath" "Info"
        
        $Users = Import-Csv -Path $CSVPath -ErrorAction Stop
        Write-ActivityLog "CSV imported successfully. Found $($Users.Count) users to process" "Info"
        
        $RequiredColumns = @('DisplayName', 'UserPrincipalName', 'Password')
        $CSVColumns = $Users[0].PSObject.Properties.Name
        
        foreach ($Column in $RequiredColumns) {
            if ($Column -notin $CSVColumns) {
                throw "Required column '$Column' not found in CSV file"
            }
        }
        
        Write-ActivityLog "CSV validation passed" "Success"
        
        $SuccessCount = 0
        $ErrorCount = 0
        $Errors = @()
        
        for ($i = 0; $i -lt $Users.Count; $i++) {
            $User = $Users[$i]
            $Progress = [Math]::Round((($i + 1) / $Users.Count) * 100, 1)
            
            try {
                Update-StatusLabel "Processing user $($i + 1) of $($Users.Count) ($Progress%): $($User.DisplayName)"
                
                $Groups = @()
                if ($User.Groups) {
                    $Groups = $User.Groups -split ',' | ForEach-Object { $_.Trim() }
                }
                
                $NewUserResult = New-M365User -DisplayName $User.DisplayName -UserPrincipalName $User.UserPrincipalName -FirstName $User.FirstName -LastName $User.LastName -Department $User.Department -JobTitle $User.JobTitle -Office $User.Office -Manager $User.Manager -LicenseType $User.LicenseType -Groups $Groups -Password $User.Password -ForcePasswordChange:([string]::IsNullOrEmpty($User.ForcePasswordChange) -or [bool]::Parse($User.ForcePasswordChange))
                
                $SuccessCount++
                Write-ActivityLog "Successfully processed user: $($User.DisplayName)" "Success"
            }
            catch {
                $ErrorCount++
                $ErrorMsg = "Failed to process user '$($User.DisplayName)': $($_.Exception.Message)"
                $Errors += $ErrorMsg
                Write-ActivityLog $ErrorMsg "Error"
            }
            
            if ($Script:ProgressBar) {
                $Script:ProgressBar.Value = $Progress
            }
        }
        
        $Summary = "CSV Import Summary:`n`n" +
                   "Total Users: $($Users.Count)`n" +
                   "Successful: $SuccessCount`n" +
                   "Failed: $ErrorCount"
        
        if ($Errors.Count -gt 0) {
            $Summary += "`n`nErrors:`n" + ($Errors -join "`n")
        }
        
        Write-ActivityLog "CSV import completed. Success: $SuccessCount, Errors: $ErrorCount" "Info"
        
        [System.Windows.Forms.MessageBox]::Show($Summary, "Import Complete", "OK", "Information")
        
        Update-StatusLabel "CSV import completed"
        
    }
    catch {
        $ErrorMsg = "CSV import failed: $($_.Exception.Message)"
        Write-ActivityLog $ErrorMsg "Error"
        [System.Windows.Forms.MessageBox]::Show($ErrorMsg, "Import Error", "OK", "Error")
        Update-StatusLabel "CSV import failed"
    }
}

# ================================
# UI HELPER FUNCTIONS
# ================================

function Update-StatusLabel {
    param([string]$Message)
    
    if ($Script:StatusLabel) {
        $Script:StatusLabel.Text = $Message
        if ($Script:StatusStrip) {
            $Script:StatusStrip.Refresh()
        }
    }
}

function Update-UIAfterConnection {
    if ($Script:ConnectButton) { 
        $Script:ConnectButton.Enabled = $false 
        $Script:ConnectButton.Text = "Connected"
        $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightGray
    }
    if ($Script:DisconnectButton) { $Script:DisconnectButton.Enabled = $true }
    if ($Script:RefreshButton) { $Script:RefreshButton.Enabled = $true }
    if ($Script:CreateUserButton) { $Script:CreateUserButton.Enabled = $true }
    if ($Script:ImportCSVButton) { $Script:ImportCSVButton.Enabled = $true }
    
    if ($Script:ConnectionInfoLabel) {
        $Script:ConnectionInfoLabel.Text = "Status: Connected to Microsoft Graph"
        $Script:ConnectionInfoLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    }
    if ($Script:ConnectionDetailsLabel) {
        $TenantName = if ($Global:TenantInfo.TenantId) { $Global:TenantInfo.TenantId.Substring(0, 8) + "..." } else { "Unknown" }
        $ExchangeStatus = if ($Global:ExchangeOnlineConnected) { "Connected" } else { "Manual Tasks Only" }
        $Script:ConnectionDetailsLabel.Text = "Tenant: $TenantName | Account: $($Global:TenantInfo.Account) | Exchange: $ExchangeStatus"
    }
    
    Update-DomainDropdown
    Update-ManagerDropdown
    Update-GroupsList
    Update-LicenseDropdown
    Refresh-TenantDataViews
}

function Update-UIAfterDisconnection {
    if ($Script:ConnectButton) { 
        $Script:ConnectButton.Enabled = $true 
        $Script:ConnectButton.Text = "Connect to M365"
        $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightGreen
    }
    if ($Script:DisconnectButton) { $Script:DisconnectButton.Enabled = $false }
    if ($Script:RefreshButton) { $Script:RefreshButton.Enabled = $false }
    if ($Script:CreateUserButton) { $Script:CreateUserButton.Enabled = $false }
    if ($Script:ImportCSVButton) { $Script:ImportCSVButton.Enabled = $false }
    
    if ($Script:ConnectionInfoLabel) {
        $Script:ConnectionInfoLabel.Text = "Status: Disconnected"
        $Script:ConnectionInfoLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    }
    if ($Script:ConnectionDetailsLabel) {
        $Script:ConnectionDetailsLabel.Text = "Click 'Connect to M365' to begin"
    }
    
    Clear-AllDropdowns
}

function Update-DomainDropdown {
    if ($Script:DomainDropdown) {
        $Script:DomainDropdown.Items.Clear()
        foreach ($Domain in $Global:AcceptedDomains) {
            $null = $Script:DomainDropdown.Items.Add($Domain.DomainName)
            if ($Domain.IsDefault) {
                $Script:DomainDropdown.SelectedItem = $Domain.DomainName
            }
        }
    }
}

function Update-ManagerDropdown {
    if ($Script:ManagerDropdown) {
        $Script:ManagerDropdown.Items.Clear()
        $null = $Script:ManagerDropdown.Items.Add("(No Manager)")
        foreach ($User in ($Global:AvailableUsers | Sort-Object DisplayName)) {
            $DisplayText = "$($User.DisplayName) ($($User.UserPrincipalName))"
            $null = $Script:ManagerDropdown.Items.Add($DisplayText)
        }
        $Script:ManagerDropdown.SelectedIndex = 0
    }
}

# FIXED: Update-GroupsList function - removed emojis to prevent corruption
function Update-GroupsList {
    if ($Script:GroupsCheckedListBox) {
        $Script:GroupsCheckedListBox.Items.Clear()
        
        # Add all groups with clean type indicators (NO EMOJIS)
        foreach ($Group in ($Global:AvailableGroups | Sort-Object GroupType, DisplayName)) {
            $DisplayText = "$($Group.DisplayName) [$($Group.GroupType)]"
            if ($Group.Mail) {
                $DisplayText += " - $($Group.Mail)"
            }
            $null = $Script:GroupsCheckedListBox.Items.Add($DisplayText)
        }
        
        # Add shared mailboxes if any found
        if ($Global:SharedMailboxes -and $Global:SharedMailboxes.Count -gt 0) {
            # Add separator
            $null = $Script:GroupsCheckedListBox.Items.Add("--- Shared Mailboxes ---")
            
            foreach ($SharedMailbox in ($Global:SharedMailboxes | Sort-Object DisplayName)) {
                $DisplayText = "$($SharedMailbox.DisplayName) [Shared Mailbox] - $($SharedMailbox.EmailAddress)"
                $null = $Script:GroupsCheckedListBox.Items.Add($DisplayText)
            }
        }
        
        Write-ActivityLog "Updated groups list with $($Global:AvailableGroups.Count) groups and $($Global:SharedMailboxes.Count) shared mailboxes" "Info"
    }
}

function Update-LicenseDropdown {
    if ($Script:LicenseDropdown) {
        $Script:LicenseDropdown.Items.Clear()
        $null = $Script:LicenseDropdown.Items.Add("(No License Assignment)")
        foreach ($LicenseType in $Global:LicenseTypes.Keys) {
            $null = $Script:LicenseDropdown.Items.Add($LicenseType)
        }
        $Script:LicenseDropdown.SelectedIndex = 0
    }
}

function Clear-AllDropdowns {
    if ($Script:DomainDropdown) { $Script:DomainDropdown.Items.Clear() }
    if ($Script:ManagerDropdown) { $Script:ManagerDropdown.Items.Clear() }
    if ($Script:GroupsCheckedListBox) { $Script:GroupsCheckedListBox.Items.Clear() }
    if ($Script:LicenseDropdown) { $Script:LicenseDropdown.Items.Clear() }
}

function Refresh-TenantDataViews {
    if ($Script:UsersDataGridView) {
        Update-UsersDataGridView
    }
    
    if ($Script:GroupsDataGridView) {
        Update-GroupsDataGridView
    }
    
    if ($Script:LicensesDataGridView) {
        Update-LicensesDataGridView
    }
}

function Update-UsersDataGridView {
    if ($Script:UsersDataGridView -and $Global:AvailableUsers.Count -gt 0) {
        $Script:UsersDataGridView.DataSource = $null
        
        $StartIndex = ($Global:CurrentPage - 1) * $Global:PageSize
        $EndIndex = [Math]::Min($StartIndex + $Global:PageSize - 1, $Global:AvailableUsers.Count - 1)
        $PagedUsers = $Global:AvailableUsers[$StartIndex..$EndIndex]
        
        $Script:UsersDataGridView.DataSource = $PagedUsers
        Update-PaginationInfo "Users" $Global:AvailableUsers.Count
    }
}

function Update-GroupsDataGridView {
    if ($Script:GroupsDataGridView -and $Global:AvailableGroups.Count -gt 0) {
        $Script:GroupsDataGridView.DataSource = $null
        
        $StartIndex = ($Global:CurrentPage - 1) * $Global:PageSize
        $EndIndex = [Math]::Min($StartIndex + $Global:PageSize - 1, $Global:AvailableGroups.Count - 1)
        $PagedGroups = $Global:AvailableGroups[$StartIndex..$EndIndex]
        
        $Script:GroupsDataGridView.DataSource = $PagedGroups
        Update-PaginationInfo "Groups" $Global:AvailableGroups.Count
    }
}

function Update-LicensesDataGridView {
    if ($Script:LicensesDataGridView -and $Global:AvailableLicenses.Count -gt 0) {
        $Script:LicensesDataGridView.DataSource = $null
        $Script:LicensesDataGridView.DataSource = $Global:AvailableLicenses
    }
}

function Update-PaginationInfo {
    param(
        [string]$DataType,
        [int]$TotalCount
    )
    
    $TotalPages = [Math]::Ceiling($TotalCount / $Global:PageSize)
    $StartItem = ($Global:CurrentPage - 1) * $Global:PageSize + 1
    $EndItem = [Math]::Min($Global:CurrentPage * $Global:PageSize, $TotalCount)
    
    $PaginationText = "$DataType $StartItem-$EndItem of $TotalCount (Page $Global:CurrentPage of $TotalPages)"
    
    if ($Script:PaginationLabel) {
        $Script:PaginationLabel.Text = $PaginationText
    }
}

# ================================
# MAIN GUI CREATION
# ================================

function New-MainForm {
    $Script:MainForm = New-Object System.Windows.Forms.Form
    $Script:MainForm.Text = "M365 User Provisioning Tool - Enterprise Edition 2025 (REPLICATION FIXED)"
    $Script:MainForm.Size = New-Object System.Drawing.Size(1400, 900)
    $Script:MainForm.StartPosition = "CenterScreen"
    $Script:MainForm.MinimumSize = New-Object System.Drawing.Size(1200, 800)
    $Script:MainForm.MaximizeBox = $true
    $Script:MainForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$env:SystemRoot\System32\shell32.dll")
    $Script:MainForm.WindowState = "Maximized"
    
    # Status strip at bottom
    $Script:StatusStrip = New-Object System.Windows.Forms.StatusStrip
    $Script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $Script:StatusLabel.Text = "Ready - Click 'Connect to M365' to begin"
    $Script:StatusLabel.Spring = $true
    $null = $Script:StatusStrip.Items.Add($Script:StatusLabel)
    $null = $Script:MainForm.Controls.Add($Script:StatusStrip)
    
    # Connection panel at top
    $ConnectionPanel = New-Object System.Windows.Forms.Panel
    $ConnectionPanel.Height = 90
    $ConnectionPanel.Dock = "Top"
    $ConnectionPanel.BackColor = [System.Drawing.Color]::LightSteelBlue
    
    # Connection buttons
    $Script:ConnectButton = New-Object System.Windows.Forms.Button
    $Script:ConnectButton.Text = "Connect to M365"
    $Script:ConnectButton.Size = New-Object System.Drawing.Size(120, 30)
    $Script:ConnectButton.Location = New-Object System.Drawing.Point(10, 10)
    $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightGreen
    $Script:ConnectButton.Add_Click({
        try {
            $Script:ConnectButton.Enabled = $false
            $Script:ConnectButton.Text = "Connecting..."
            $Script:ConnectButton.Refresh()
            Connect-ToMicrosoftGraph
        }
        finally {
            if (-not $Global:IsConnected) {
                $Script:ConnectButton.Enabled = $true
                $Script:ConnectButton.Text = "Connect to M365"
            }
        }
    })
    
    $Script:DisconnectButton = New-Object System.Windows.Forms.Button
    $Script:DisconnectButton.Text = "Disconnect"
    $Script:DisconnectButton.Size = New-Object System.Drawing.Size(100, 30)
    $Script:DisconnectButton.Location = New-Object System.Drawing.Point(140, 10)
    $Script:DisconnectButton.Enabled = $false
    $Script:DisconnectButton.BackColor = [System.Drawing.Color]::LightCoral
    $Script:DisconnectButton.Add_Click({
        Disconnect-FromMicrosoftGraph
    })
    
    $Script:RefreshButton = New-Object System.Windows.Forms.Button
    $Script:RefreshButton.Text = "Refresh Data"
    $Script:RefreshButton.Size = New-Object System.Drawing.Size(100, 30)
    $Script:RefreshButton.Location = New-Object System.Drawing.Point(250, 10)
    $Script:RefreshButton.Enabled = $false
    $Script:RefreshButton.Add_Click({
        if ($Global:IsConnected) {
            Update-StatusLabel "Refreshing tenant data..."
            Start-TenantDiscovery
            Refresh-TenantDataViews
            Update-StatusLabel "Tenant data refreshed"
        }
    })
    
    # Connection info labels
    $Script:ConnectionInfoLabel = New-Object System.Windows.Forms.Label
    $Script:ConnectionInfoLabel.Text = "Status: Disconnected"
    $Script:ConnectionInfoLabel.Location = New-Object System.Drawing.Point(10, 45)
    $Script:ConnectionInfoLabel.Size = New-Object System.Drawing.Size(600, 20)
    $Script:ConnectionInfoLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    
    $Script:ConnectionDetailsLabel = New-Object System.Windows.Forms.Label
    $Script:ConnectionDetailsLabel.Text = "Click 'Connect to M365' to begin"
    $Script:ConnectionDetailsLabel.Location = New-Object System.Drawing.Point(10, 65)
    $Script:ConnectionDetailsLabel.Size = New-Object System.Drawing.Size(600, 20)
    $Script:ConnectionDetailsLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    
    $null = $ConnectionPanel.Controls.Add($Script:ConnectButton)
    $null = $ConnectionPanel.Controls.Add($Script:DisconnectButton)
    $null = $ConnectionPanel.Controls.Add($Script:RefreshButton)
    $null = $ConnectionPanel.Controls.Add($Script:ConnectionInfoLabel)
    $null = $ConnectionPanel.Controls.Add($Script:ConnectionDetailsLabel)
    
    # Add connection panel to form FIRST
    $null = $Script:MainForm.Controls.Add($ConnectionPanel)
    
    # Tab separator panel to ensure proper spacing
    $TabSeparatorPanel = New-Object System.Windows.Forms.Panel
    $TabSeparatorPanel.Height = 10
    $TabSeparatorPanel.Dock = "Top"
    $TabSeparatorPanel.BackColor = [System.Drawing.Color]::LightGray
    $null = $Script:MainForm.Controls.Add($TabSeparatorPanel)
    
    # Tab control with proper positioning - NOT docked to fill
    $Script:TabControl = New-Object System.Windows.Forms.TabControl
    $Script:TabControl.Location = New-Object System.Drawing.Point(0, 100)  # Start below connection panel
    $Script:TabControl.Size = New-Object System.Drawing.Size($Script:MainForm.ClientSize.Width, ($Script:MainForm.ClientSize.Height - 150))
    $Script:TabControl.Anchor = "Top,Bottom,Left,Right"
    $Script:TabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $Script:TabControl.Appearance = "Normal"
    $Script:TabControl.Alignment = "Top"
    $Script:TabControl.HotTrack = $true
    $Script:TabControl.Multiline = $false
    $Script:TabControl.ShowToolTips = $true
    $Script:TabControl.ItemSize = New-Object System.Drawing.Size(120, 25)  # Ensure tab headers are visible
    $Script:TabControl.SizeMode = "Normal"
    
    Write-Host "Creating tabs..." -ForegroundColor Yellow
    
    # Create tabs
    New-UserCreationTab
    New-BulkImportTab  
    New-TenantDataTab
    New-ActivityLogTab
    
    Write-Host "Created $($Script:TabControl.TabCount) tabs" -ForegroundColor Green
    
    # Add TabControl to form AFTER tabs are created
    $null = $Script:MainForm.Controls.Add($Script:TabControl)
    
    # Ensure first tab is selected and visible
    if ($Script:TabControl.TabCount -gt 0) {
        $Script:TabControl.SelectedIndex = 0
        Write-Host "TabControl positioned at Y=100 with $($Script:TabControl.TabCount) tabs" -ForegroundColor Green
        Write-Host "Tab headers should be visible with size: $($Script:TabControl.ItemSize)" -ForegroundColor Green
    }
    else {
        Write-Host "ERROR: No tabs were created!" -ForegroundColor Red
    }
    
    # Force refresh to ensure visibility
    $Script:TabControl.Refresh()
    
    return $Script:MainForm
}

function New-UserCreationTab {
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Create User"
    $Tab.BackColor = [System.Drawing.Color]::White
    
    $MainPanel = New-Object System.Windows.Forms.Panel
    $MainPanel.Dock = "Fill"
    $MainPanel.AutoScroll = $true
    $MainPanel.Padding = New-Object System.Windows.Forms.Padding(15)
    
    # Add a title header to ensure visibility
    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Text = "Create New M365 User Account"
    $TitleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $TitleLabel.Size = New-Object System.Drawing.Size(400, 25)
    $TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $TitleLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    
    # User details group (left side) - moved down significantly
    $UserDetailsGroup = New-Object System.Windows.Forms.GroupBox
    $UserDetailsGroup.Text = "User Details"
    $UserDetailsGroup.Location = New-Object System.Drawing.Point(20, 50)
    $UserDetailsGroup.Size = New-Object System.Drawing.Size(480, 380)
    
    # Display Name - moved down within the group
    $DisplayNameLabel = New-Object System.Windows.Forms.Label
    $DisplayNameLabel.Text = "Display Name *:"
    $DisplayNameLabel.Location = New-Object System.Drawing.Point(10, 35)
    $DisplayNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:DisplayNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:DisplayNameTextBox.Location = New-Object System.Drawing.Point(120, 33)
    $Script:DisplayNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    # Display Name helper text
    $DisplayNameHelpLabel = New-Object System.Windows.Forms.Label
    $DisplayNameHelpLabel.Text = "(Auto-populated from First + Last Name)"
    $DisplayNameHelpLabel.Location = New-Object System.Drawing.Point(330, 35)
    $DisplayNameHelpLabel.Size = New-Object System.Drawing.Size(140, 15)
    $DisplayNameHelpLabel.Font = New-Object System.Drawing.Font("Segoe UI", 7.5, [System.Drawing.FontStyle]::Italic)
    $DisplayNameHelpLabel.ForeColor = [System.Drawing.Color]::Gray
    
    # First Name
    $FirstNameLabel = New-Object System.Windows.Forms.Label
    $FirstNameLabel.Text = "First Name:"
    $FirstNameLabel.Location = New-Object System.Drawing.Point(10, 65)
    $FirstNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:FirstNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:FirstNameTextBox.Location = New-Object System.Drawing.Point(120, 63)
    $Script:FirstNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $Script:FirstNameTextBox.Add_TextChanged({
        # Auto-populate Display Name when First Name or Last Name changes
        if (-not [string]::IsNullOrWhiteSpace($Script:FirstNameTextBox.Text) -or -not [string]::IsNullOrWhiteSpace($Script:LastNameTextBox.Text)) {
            $FirstName = $Script:FirstNameTextBox.Text.Trim()
            $LastName = $Script:LastNameTextBox.Text.Trim()
            $Script:DisplayNameTextBox.Text = "$FirstName $LastName".Trim()
        }
    })
    
    # Last Name
    $LastNameLabel = New-Object System.Windows.Forms.Label
    $LastNameLabel.Text = "Last Name:"
    $LastNameLabel.Location = New-Object System.Drawing.Point(10, 95)
    $LastNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:LastNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:LastNameTextBox.Location = New-Object System.Drawing.Point(120, 93)
    $Script:LastNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $Script:LastNameTextBox.Add_TextChanged({
        # Auto-populate Display Name when First Name or Last Name changes
        if (-not [string]::IsNullOrWhiteSpace($Script:FirstNameTextBox.Text) -or -not [string]::IsNullOrWhiteSpace($Script:LastNameTextBox.Text)) {
            $FirstName = $Script:FirstNameTextBox.Text.Trim()
            $LastName = $Script:LastNameTextBox.Text.Trim()
            $Script:DisplayNameTextBox.Text = "$FirstName $LastName".Trim()
        }
    })
    
    # Username
    $UsernameLabel = New-Object System.Windows.Forms.Label
    $UsernameLabel.Text = "Username *:"
    $UsernameLabel.Location = New-Object System.Drawing.Point(10, 125)
    $UsernameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:UsernameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:UsernameTextBox.Location = New-Object System.Drawing.Point(120, 123)
    $Script:UsernameTextBox.Size = New-Object System.Drawing.Size(130, 20)
    
    # Domain dropdown
    $DomainLabel = New-Object System.Windows.Forms.Label
    $DomainLabel.Text = "@"
    $DomainLabel.Location = New-Object System.Drawing.Point(255, 125)
    $DomainLabel.Size = New-Object System.Drawing.Size(15, 20)
    
    $Script:DomainDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:DomainDropdown.Location = New-Object System.Drawing.Point(270, 123)
    $Script:DomainDropdown.Size = New-Object System.Drawing.Size(180, 20)
    $Script:DomainDropdown.DropDownStyle = "DropDownList"
    
    # Password
    $PasswordLabel = New-Object System.Windows.Forms.Label
    $PasswordLabel.Text = "Password *:"
    $PasswordLabel.Location = New-Object System.Drawing.Point(10, 155)
    $PasswordLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:PasswordTextBox = New-Object System.Windows.Forms.TextBox
    $Script:PasswordTextBox.Location = New-Object System.Drawing.Point(120, 153)
    $Script:PasswordTextBox.Size = New-Object System.Drawing.Size(150, 20)
    $Script:PasswordTextBox.UseSystemPasswordChar = $true
    
    # Generate password button
    $GeneratePasswordButton = New-Object System.Windows.Forms.Button
    $GeneratePasswordButton.Text = "Generate"
    $GeneratePasswordButton.Location = New-Object System.Drawing.Point(280, 153)
    $GeneratePasswordButton.Size = New-Object System.Drawing.Size(70, 22)
    $GeneratePasswordButton.Add_Click({
        $GeneratedPassword = New-SecurePassword
        $Script:PasswordTextBox.Text = $GeneratedPassword
        [System.Windows.Forms.MessageBox]::Show("Generated password: $GeneratedPassword`n`nPlease save this password securely!", "Password Generated", "OK", "Information")
    })
    
    # Department
    $DepartmentLabel = New-Object System.Windows.Forms.Label
    $DepartmentLabel.Text = "Department:"
    $DepartmentLabel.Location = New-Object System.Drawing.Point(10, 185)
    $DepartmentLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:DepartmentTextBox = New-Object System.Windows.Forms.TextBox
    $Script:DepartmentTextBox.Location = New-Object System.Drawing.Point(120, 183)
    $Script:DepartmentTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    # Job Title
    $JobTitleLabel = New-Object System.Windows.Forms.Label
    $JobTitleLabel.Text = "Job Title:"
    $JobTitleLabel.Location = New-Object System.Drawing.Point(10, 215)
    $JobTitleLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:JobTitleTextBox = New-Object System.Windows.Forms.TextBox
    $Script:JobTitleTextBox.Location = New-Object System.Drawing.Point(120, 213)
    $Script:JobTitleTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    # Office Location
    $OfficeLabel = New-Object System.Windows.Forms.Label
    $OfficeLabel.Text = "Office Location:"
    $OfficeLabel.Location = New-Object System.Drawing.Point(10, 245)
    $OfficeLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:OfficeDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:OfficeDropdown.Location = New-Object System.Drawing.Point(120, 243)
    $Script:OfficeDropdown.Size = New-Object System.Drawing.Size(250, 20)
    $Script:OfficeDropdown.DropDownStyle = "DropDownList"
    $Global:UKLocations | ForEach-Object { $null = $Script:OfficeDropdown.Items.Add($_) }
    
    $null = $UserDetailsGroup.Controls.Add($DisplayNameLabel)
    $null = $UserDetailsGroup.Controls.Add($Script:DisplayNameTextBox)
    $null = $UserDetailsGroup.Controls.Add($DisplayNameHelpLabel)
    $null = $UserDetailsGroup.Controls.Add($FirstNameLabel)
    $null = $UserDetailsGroup.Controls.Add($Script:FirstNameTextBox)
    $null = $UserDetailsGroup.Controls.Add($LastNameLabel)
    $null = $UserDetailsGroup.Controls.Add($Script:LastNameTextBox)
    $null = $UserDetailsGroup.Controls.Add($UsernameLabel)
    $null = $UserDetailsGroup.Controls.Add($Script:UsernameTextBox)
    $null = $UserDetailsGroup.Controls.Add($DomainLabel)
    $null = $UserDetailsGroup.Controls.Add($Script:DomainDropdown)
    $null = $UserDetailsGroup.Controls.Add($PasswordLabel)
    $null = $UserDetailsGroup.Controls.Add($Script:PasswordTextBox)
    $null = $UserDetailsGroup.Controls.Add($GeneratePasswordButton)
    $null = $UserDetailsGroup.Controls.Add($DepartmentLabel)
    $null = $UserDetailsGroup.Controls.Add($Script:DepartmentTextBox)
    $null = $UserDetailsGroup.Controls.Add($JobTitleLabel)
    $null = $UserDetailsGroup.Controls.Add($Script:JobTitleTextBox)
    $null = $UserDetailsGroup.Controls.Add($OfficeLabel)
    $null = $UserDetailsGroup.Controls.Add($Script:OfficeDropdown)
    
    # Management group (right side) - moved down to match
    $ManagementGroup = New-Object System.Windows.Forms.GroupBox
    $ManagementGroup.Text = "Management & Licensing"
    $ManagementGroup.Location = New-Object System.Drawing.Point(520, 50)
    $ManagementGroup.Size = New-Object System.Drawing.Size(450, 380)
    
    # Manager
    $ManagerLabel = New-Object System.Windows.Forms.Label
    $ManagerLabel.Text = "Manager:"
    $ManagerLabel.Location = New-Object System.Drawing.Point(10, 30)
    $ManagerLabel.Size = New-Object System.Drawing.Size(80, 20)
    
    $Script:ManagerDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:ManagerDropdown.Location = New-Object System.Drawing.Point(100, 28)
    $Script:ManagerDropdown.Size = New-Object System.Drawing.Size(330, 20)
    $Script:ManagerDropdown.DropDownStyle = "DropDownList"
    
    # License Type
    $LicenseLabel = New-Object System.Windows.Forms.Label
    $LicenseLabel.Text = "License Type:"
    $LicenseLabel.Location = New-Object System.Drawing.Point(10, 70)
    $LicenseLabel.Size = New-Object System.Drawing.Size(80, 20)
    
    $Script:LicenseDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:LicenseDropdown.Location = New-Object System.Drawing.Point(100, 68)
    $Script:LicenseDropdown.Size = New-Object System.Drawing.Size(330, 20)
    $Script:LicenseDropdown.DropDownStyle = "DropDownList"
    
    # License info
    $LicenseInfoLabel = New-Object System.Windows.Forms.Label
    $LicenseInfoLabel.Text = "Note: License assignment is handled via CustomAttribute1"
    $LicenseInfoLabel.Location = New-Object System.Drawing.Point(10, 100)
    $LicenseInfoLabel.Size = New-Object System.Drawing.Size(430, 30)
    $LicenseInfoLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $LicenseInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)
    
    # License types info with Exchange Online note
    $LicenseTypesLabel = New-Object System.Windows.Forms.Label
    $LicenseTypesLabel.Text = "Available License Types:`n• BusinessBasic`n• BusinessPremium`n• BusinessStandard`n• ExchangeOnline1`n• ExchangeOnline2`n`nEnhancements:`n✓ Emoji corruption fixed`n✓ Exchange Online connection RESTORED`n✓ Azure AD replication delays FIXED`n✓ Second pass retry system for maximum success`n✓ Distribution lists and shared mailboxes working`n✓ Browser-based authentication for both services"
    $LicenseTypesLabel.Location = New-Object System.Drawing.Point(10, 140)
    $LicenseTypesLabel.Size = New-Object System.Drawing.Size(430, 180)
    $LicenseTypesLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    $LicenseTypesLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    
    $null = $ManagementGroup.Controls.Add($ManagerLabel)
    $null = $ManagementGroup.Controls.Add($Script:ManagerDropdown)
    $null = $ManagementGroup.Controls.Add($LicenseLabel)
    $null = $ManagementGroup.Controls.Add($Script:LicenseDropdown)
    $null = $ManagementGroup.Controls.Add($LicenseInfoLabel)
    $null = $ManagementGroup.Controls.Add($LicenseTypesLabel)
    
    # Groups group (full width below) - moved down significantly
    $GroupsGroup = New-Object System.Windows.Forms.GroupBox
    $GroupsGroup.Text = "Group Membership & Shared Mailboxes"
    $GroupsGroup.Location = New-Object System.Drawing.Point(20, 450)
    $GroupsGroup.Size = New-Object System.Drawing.Size(950, 220)
    
    $Script:GroupsCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $Script:GroupsCheckedListBox.Location = New-Object System.Drawing.Point(10, 20)
    $Script:GroupsCheckedListBox.Size = New-Object System.Drawing.Size(930, 190)
    $Script:GroupsCheckedListBox.CheckOnClick = $true
    
    $null = $GroupsGroup.Controls.Add($Script:GroupsCheckedListBox)
    
    # Action buttons - moved down accordingly
    $ButtonPanel = New-Object System.Windows.Forms.Panel
    $ButtonPanel.Location = New-Object System.Drawing.Point(20, 680)
    $ButtonPanel.Size = New-Object System.Drawing.Size(950, 50)
    
    $Script:CreateUserButton = New-Object System.Windows.Forms.Button
    $Script:CreateUserButton.Text = "Create User"
    $Script:CreateUserButton.Size = New-Object System.Drawing.Size(100, 30)
    $Script:CreateUserButton.Location = New-Object System.Drawing.Point(0, 10)
    $Script:CreateUserButton.BackColor = [System.Drawing.Color]::LightGreen
    $Script:CreateUserButton.Enabled = $false
    $Script:CreateUserButton.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($Script:DisplayNameTextBox.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Display Name is required", "Validation Error", "OK", "Warning")
                return
            }
            
            if ([string]::IsNullOrWhiteSpace($Script:UsernameTextBox.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Username is required", "Validation Error", "OK", "Warning")
                return
            }
            
            if ([string]::IsNullOrWhiteSpace($Script:PasswordTextBox.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Password is required", "Validation Error", "OK", "Warning")
                return
            }
            
            if (-not $Script:DomainDropdown.SelectedItem) {
                [System.Windows.Forms.MessageBox]::Show("Please select a domain", "Validation Error", "OK", "Warning")
                return
            }
            
            $UserPrincipalName = "$($Script:UsernameTextBox.Text)@$($Script:DomainDropdown.SelectedItem)"
            
            # FIXED: Get selected groups using the new process
            $SelectedGroups = @()
            for ($i = 0; $i -lt $Script:GroupsCheckedListBox.CheckedItems.Count; $i++) {
                $SelectedGroups += $Script:GroupsCheckedListBox.CheckedItems[$i].ToString()
            }
            
            $Manager = $null
            if ($Script:ManagerDropdown.SelectedIndex -gt 0) {
                $ManagerText = $Script:ManagerDropdown.SelectedItem
                # Extract UPN from "Display Name (user@domain.com)" format
                if ($ManagerText -match '\(([^)]+)\)$') {
                    $Manager = $Matches[1]
                }
            }
            
            $LicenseType = $null
            if ($Script:LicenseDropdown.SelectedIndex -gt 0) {
                $LicenseType = $Script:LicenseDropdown.SelectedItem
            }
            
            $Script:CreateUserButton.Enabled = $false
            $Script:CreateUserButton.Text = "Creating..."
            
            $NewUserResult = New-M365User -DisplayName $Script:DisplayNameTextBox.Text -UserPrincipalName $UserPrincipalName -FirstName $Script:FirstNameTextBox.Text -LastName $Script:LastNameTextBox.Text -Department $Script:DepartmentTextBox.Text -JobTitle $Script:JobTitleTextBox.Text -Office $Script:OfficeDropdown.SelectedItem -Manager $Manager -LicenseType $LicenseType -Groups $SelectedGroups -Password $Script:PasswordTextBox.Text
            
            [System.Windows.Forms.MessageBox]::Show("User created successfully!`n`nDisplay Name: $($NewUserResult.DisplayName)`nUPN: $($NewUserResult.UserPrincipalName)`nUser ID: $($NewUserResult.Id)", "User Created", "OK", "Information")
            
            Clear-UserCreationForm
            
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to create user: $($_.Exception.Message)", "Error", "OK", "Error")
        }
        finally {
            $Script:CreateUserButton.Enabled = $true
            $Script:CreateUserButton.Text = "Create User"
        }
    })
    
    $ClearFormButton = New-Object System.Windows.Forms.Button
    $ClearFormButton.Text = "Clear Form"
    $ClearFormButton.Size = New-Object System.Drawing.Size(100, 30)
    $ClearFormButton.Location = New-Object System.Drawing.Point(110, 10)
    $ClearFormButton.Add_Click({
        Clear-UserCreationForm
    })
    
    $null = $ButtonPanel.Controls.Add($Script:CreateUserButton)
    $null = $ButtonPanel.Controls.Add($ClearFormButton)
    
    $null = $MainPanel.Controls.Add($TitleLabel)
    $null = $MainPanel.Controls.Add($UserDetailsGroup)
    $null = $MainPanel.Controls.Add($ManagementGroup)
    $null = $MainPanel.Controls.Add($GroupsGroup)
    $null = $MainPanel.Controls.Add($ButtonPanel)
    
    $null = $Tab.Controls.Add($MainPanel)
    $null = $Script:TabControl.TabPages.Add($Tab)
}

function New-BulkImportTab {
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Bulk Import"
    $Tab.BackColor = [System.Drawing.Color]::White
    
    # Instructions
    $InstructionsGroup = New-Object System.Windows.Forms.GroupBox
    $InstructionsGroup.Text = "CSV Import Instructions"
    $InstructionsGroup.Location = New-Object System.Drawing.Point(10, 10)
    $InstructionsGroup.Size = New-Object System.Drawing.Size(900, 120)
    
    $InstructionsText = New-Object System.Windows.Forms.RichTextBox
    $InstructionsText.Dock = "Fill"
    $InstructionsText.ReadOnly = $true
    $InstructionsText.Text = "CSV Format Requirements:`n• Required columns: DisplayName, UserPrincipalName, Password`n• Optional columns: FirstName, LastName, Department, JobTitle, Office, Manager, LicenseType, Groups, ForcePasswordChange`n• Groups column should contain comma-separated group names`n• Manager should be the UPN of the manager`n• Office should match one of the UK locations from the dropdown`n• LicenseType should be one of: BusinessBasic, BusinessPremium, BusinessStandard, ExchangeOnline1, ExchangeOnline2`n`nExample CSV content:`nDisplayName,UserPrincipalName,FirstName,LastName,Department,JobTitle,Office,Manager,LicenseType,Groups,Password,ForcePasswordChange`nJohn Smith,john.smith@company.com,John,Smith,IT,Developer,United Kingdom - London,manager@company.com,BusinessPremium,IT Team;Developers,TempPass123!,true"
    $InstructionsText.BackColor = [System.Drawing.Color]::LightYellow
    
    $null = $InstructionsGroup.Controls.Add($InstructionsText)
    
    # File selection
    $FileSelectionGroup = New-Object System.Windows.Forms.GroupBox
    $FileSelectionGroup.Text = "File Selection"
    $FileSelectionGroup.Location = New-Object System.Drawing.Point(10, 140)
    $FileSelectionGroup.Size = New-Object System.Drawing.Size(900, 80)
    
    $FilePathLabel = New-Object System.Windows.Forms.Label
    $FilePathLabel.Text = "CSV File:"
    $FilePathLabel.Location = New-Object System.Drawing.Point(10, 30)
    $FilePathLabel.Size = New-Object System.Drawing.Size(60, 20)
    
    $Script:FilePathTextBox = New-Object System.Windows.Forms.TextBox
    $Script:FilePathTextBox.Location = New-Object System.Drawing.Point(80, 28)
    $Script:FilePathTextBox.Size = New-Object System.Drawing.Size(600, 20)
    $Script:FilePathTextBox.ReadOnly = $true
    
    $BrowseButton = New-Object System.Windows.Forms.Button
    $BrowseButton.Text = "Browse..."
    $BrowseButton.Location = New-Object System.Drawing.Point(690, 28)
    $BrowseButton.Size = New-Object System.Drawing.Size(80, 23)
    $BrowseButton.Add_Click({
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $OpenFileDialog.FilterIndex = 1
        
        if ($OpenFileDialog.ShowDialog() -eq "OK") {
            $Script:FilePathTextBox.Text = $OpenFileDialog.FileName
        }
    })
    
    $null = $FileSelectionGroup.Controls.Add($FilePathLabel)
    $null = $FileSelectionGroup.Controls.Add($Script:FilePathTextBox)
    $null = $FileSelectionGroup.Controls.Add($BrowseButton)
    
    # Progress
    $ProgressGroup = New-Object System.Windows.Forms.GroupBox
    $ProgressGroup.Text = "Import Progress"
    $ProgressGroup.Location = New-Object System.Drawing.Point(10, 230)
    $ProgressGroup.Size = New-Object System.Drawing.Size(900, 100)
    
    $Script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $Script:ProgressBar.Location = New-Object System.Drawing.Point(10, 30)
    $Script:ProgressBar.Size = New-Object System.Drawing.Size(880, 25)
    
    $Script:ProgressLabel = New-Object System.Windows.Forms.Label
    $Script:ProgressLabel.Text = "Ready to import"
    $Script:ProgressLabel.Location = New-Object System.Drawing.Point(10, 60)
    $Script:ProgressLabel.Size = New-Object System.Drawing.Size(880, 20)
    
    $null = $ProgressGroup.Controls.Add($Script:ProgressBar)
    $null = $ProgressGroup.Controls.Add($Script:ProgressLabel)
    
    # Import buttons
    $ImportButtonPanel = New-Object System.Windows.Forms.Panel
    $ImportButtonPanel.Location = New-Object System.Drawing.Point(10, 340)
    $ImportButtonPanel.Size = New-Object System.Drawing.Size(900, 50)
    
    $Script:ImportCSVButton = New-Object System.Windows.Forms.Button
    $Script:ImportCSVButton.Text = "Import Users from CSV"
    $Script:ImportCSVButton.Size = New-Object System.Drawing.Size(150, 30)
    $Script:ImportCSVButton.Location = New-Object System.Drawing.Point(0, 10)
    $Script:ImportCSVButton.BackColor = [System.Drawing.Color]::LightBlue
    $Script:ImportCSVButton.Enabled = $false
    $Script:ImportCSVButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($Script:FilePathTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a CSV file first", "No File Selected", "OK", "Warning")
            return
        }
        
        if (-not (Test-Path $Script:FilePathTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("The selected file does not exist", "File Not Found", "OK", "Error")
            return
        }
        
        $Confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to import users from the selected CSV file?`n`nThis will create new user accounts in your M365 tenant.", "Confirm Import", "YesNo", "Question")
        
        if ($Confirmation -eq "Yes") {
            try {
                $Script:ImportCSVButton.Enabled = $false
                $Script:ImportCSVButton.Text = "Importing..."
                $Script:ProgressBar.Value = 0
                $Script:ProgressLabel.Text = "Starting import..."
                
                Import-UsersFromCSV -CSVPath $Script:FilePathTextBox.Text
            }
            finally {
                $Script:ImportCSVButton.Enabled = $true
                $Script:ImportCSVButton.Text = "Import Users from CSV"
                $Script:ProgressBar.Value = 0
                $Script:ProgressLabel.Text = "Import completed"
            }
        }
    })
    
    $ValidateCSVButton = New-Object System.Windows.Forms.Button
    $ValidateCSVButton.Text = "Validate CSV"
    $ValidateCSVButton.Size = New-Object System.Drawing.Size(100, 30)
    $ValidateCSVButton.Location = New-Object System.Drawing.Point(160, 10)
    $ValidateCSVButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($Script:FilePathTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a CSV file first", "No File Selected", "OK", "Warning")
            return
        }
        
        try {
            $Users = Import-Csv -Path $Script:FilePathTextBox.Text -ErrorAction Stop
            $RequiredColumns = @('DisplayName', 'UserPrincipalName', 'Password')
            $CSVColumns = $Users[0].PSObject.Properties.Name
            
            $MissingColumns = $RequiredColumns | Where-Object { $_ -notin $CSVColumns }
            
            if ($MissingColumns.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Missing required columns: $($MissingColumns -join ', ')", "Validation Failed", "OK", "Error")
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("CSV validation passed!`n`nRows to import: $($Users.Count)`nColumns found: $($CSVColumns -join ', ')", "Validation Successful", "OK", "Information")
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to validate CSV: $($_.Exception.Message)", "Validation Error", "OK", "Error")
        }
    })
    
    $DownloadTemplateButton = New-Object System.Windows.Forms.Button
    $DownloadTemplateButton.Text = "Download Template"
    $DownloadTemplateButton.Size = New-Object System.Drawing.Size(120, 30)
    $DownloadTemplateButton.Location = New-Object System.Drawing.Point(270, 10)
    $DownloadTemplateButton.Add_Click({
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $SaveFileDialog.FilterIndex = 1
        $SaveFileDialog.FileName = "M365_Users_Template.csv"
        
        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            try {
                $Template = "DisplayName,UserPrincipalName,FirstName,LastName,Department,JobTitle,Office,Manager,LicenseType,Groups,Password,ForcePasswordChange`nJohn Smith,john.smith@yourcompany.com,John,Smith,IT,Developer,United Kingdom - London,manager@yourcompany.com,BusinessPremium,IT Team;Developers,TempPass123!,true`nJane Doe,jane.doe@yourcompany.com,Jane,Doe,Marketing,Manager,Remote/Home Working - UK,,BusinessStandard,Marketing Team,SecurePass456!,false"
                $Template | Out-File -FilePath $SaveFileDialog.FileName -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Template saved successfully to:`n$($SaveFileDialog.FileName)", "Template Saved", "OK", "Information")
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to save template: $($_.Exception.Message)", "Save Error", "OK", "Error")
            }
        }
    })
    
    $null = $ImportButtonPanel.Controls.Add($Script:ImportCSVButton)
    $null = $ImportButtonPanel.Controls.Add($ValidateCSVButton)
    $null = $ImportButtonPanel.Controls.Add($DownloadTemplateButton)
    
    $null = $Tab.Controls.Add($InstructionsGroup)
    $null = $Tab.Controls.Add($FileSelectionGroup)
    $null = $Tab.Controls.Add($ProgressGroup)
    $null = $Tab.Controls.Add($ImportButtonPanel)
    
    $null = $Script:TabControl.TabPages.Add($Tab)
}

function New-TenantDataTab {
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Tenant Data"
    $Tab.BackColor = [System.Drawing.Color]::White
    
    $SubTabControl = New-Object System.Windows.Forms.TabControl
    $SubTabControl.Dock = "Fill"
    
    # Users tab
    $UsersTab = New-Object System.Windows.Forms.TabPage
    $UsersTab.Text = "Users"
    
    $Script:UsersDataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:UsersDataGridView.Dock = "Fill"
    $Script:UsersDataGridView.ReadOnly = $true
    $Script:UsersDataGridView.AutoSizeColumnsMode = "AllCells"
    $Script:UsersDataGridView.AllowUserToAddRows = $false
    
    $null = $UsersTab.Controls.Add($Script:UsersDataGridView)
    
    # Groups tab
    $GroupsTab = New-Object System.Windows.Forms.TabPage
    $GroupsTab.Text = "Groups"
    
    $Script:GroupsDataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:GroupsDataGridView.Dock = "Fill"
    $Script:GroupsDataGridView.ReadOnly = $true
    $Script:GroupsDataGridView.AutoSizeColumnsMode = "AllCells"
    $Script:GroupsDataGridView.AllowUserToAddRows = $false
    
    $null = $GroupsTab.Controls.Add($Script:GroupsDataGridView)
    
    # Licenses tab
    $LicensesTab = New-Object System.Windows.Forms.TabPage
    $LicensesTab.Text = "Licenses"
    
    $Script:LicensesDataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:LicensesDataGridView.Dock = "Fill"
    $Script:LicensesDataGridView.ReadOnly = $true
    $Script:LicensesDataGridView.AutoSizeColumnsMode = "AllCells"
    $Script:LicensesDataGridView.AllowUserToAddRows = $false
    
    $null = $LicensesTab.Controls.Add($Script:LicensesDataGridView)
    
    # Domains tab
    $DomainsTab = New-Object System.Windows.Forms.TabPage
    $DomainsTab.Text = "Domains"
    
    $Script:DomainsDataGridView = New-Object System.Windows.Forms.DataGridView
    $Script:DomainsDataGridView.Dock = "Fill"
    $Script:DomainsDataGridView.ReadOnly = $true
    $Script:DomainsDataGridView.AutoSizeColumnsMode = "AllCells"
    $Script:DomainsDataGridView.AllowUserToAddRows = $false
    
    $null = $DomainsTab.Controls.Add($Script:DomainsDataGridView)
    
    # Pagination panel
    $PaginationPanel = New-Object System.Windows.Forms.Panel
    $PaginationPanel.Height = 40
    $PaginationPanel.Dock = "Bottom"
    
    $Script:PaginationLabel = New-Object System.Windows.Forms.Label
    $Script:PaginationLabel.Text = "No data loaded"
    $Script:PaginationLabel.Location = New-Object System.Drawing.Point(10, 10)
    $Script:PaginationLabel.Size = New-Object System.Drawing.Size(400, 20)
    
    $PrevPageButton = New-Object System.Windows.Forms.Button
    $PrevPageButton.Text = "< Previous"
    $PrevPageButton.Location = New-Object System.Drawing.Point(420, 8)
    $PrevPageButton.Size = New-Object System.Drawing.Size(80, 25)
    $PrevPageButton.Add_Click({
        if ($Global:CurrentPage -gt 1) {
            $Global:CurrentPage--
            Refresh-TenantDataViews
        }
    })
    
    $NextPageButton = New-Object System.Windows.Forms.Button
    $NextPageButton.Text = "Next >"
    $NextPageButton.Location = New-Object System.Drawing.Point(510, 8)
    $NextPageButton.Size = New-Object System.Drawing.Size(80, 25)
    $NextPageButton.Add_Click({
        $TotalPages = [Math]::Ceiling($Global:AvailableUsers.Count / $Global:PageSize)
        if ($Global:CurrentPage -lt $TotalPages) {
            $Global:CurrentPage++
            Refresh-TenantDataViews
        }
    })
    
    $null = $PaginationPanel.Controls.Add($Script:PaginationLabel)
    $null = $PaginationPanel.Controls.Add($PrevPageButton)
    $null = $PaginationPanel.Controls.Add($NextPageButton)
    
    $null = $SubTabControl.TabPages.Add($UsersTab)
    $null = $SubTabControl.TabPages.Add($GroupsTab)
    $null = $SubTabControl.TabPages.Add($LicensesTab)
    $null = $SubTabControl.TabPages.Add($DomainsTab)
    
    $null = $Tab.Controls.Add($PaginationPanel)
    $null = $Tab.Controls.Add($SubTabControl)
    
    $null = $Script:TabControl.TabPages.Add($Tab)
}

function New-ActivityLogTab {
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Activity Log"
    $Tab.BackColor = [System.Drawing.Color]::White
    
    $Script:LogTextBox = New-Object System.Windows.Forms.RichTextBox
    $Script:LogTextBox.Dock = "Fill"
    $Script:LogTextBox.ReadOnly = $true
    $Script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $Script:LogTextBox.BackColor = [System.Drawing.Color]::Black
    $Script:LogTextBox.ForeColor = [System.Drawing.Color]::White
    
    $LogControlsPanel = New-Object System.Windows.Forms.Panel
    $LogControlsPanel.Height = 40
    $LogControlsPanel.Dock = "Bottom"
    
    $RefreshLogButton = New-Object System.Windows.Forms.Button
    $RefreshLogButton.Text = "Refresh Log"
    $RefreshLogButton.Location = New-Object System.Drawing.Point(10, 8)
    $RefreshLogButton.Size = New-Object System.Drawing.Size(100, 25)
    $RefreshLogButton.Add_Click({
        Update-ActivityLogDisplay
    })
    
    $ClearLogButton = New-Object System.Windows.Forms.Button
    $ClearLogButton.Text = "Clear Log"
    $ClearLogButton.Location = New-Object System.Drawing.Point(120, 8)
    $ClearLogButton.Size = New-Object System.Drawing.Size(80, 25)
    $ClearLogButton.Add_Click({
        $Global:ActivityLog = @()
        $Script:LogTextBox.Clear()
    })
    
    $SaveLogButton = New-Object System.Windows.Forms.Button
    $SaveLogButton.Text = "Save Log"
    $SaveLogButton.Location = New-Object System.Drawing.Point(210, 8)
    $SaveLogButton.Size = New-Object System.Drawing.Size(80, 25)
    $SaveLogButton.Add_Click({
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt"
        $SaveFileDialog.FilterIndex = 1
        $SaveFileDialog.FileName = "M365_Activity_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        
        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            try {
                $Global:ActivityLog | Out-File -FilePath $SaveFileDialog.FileName -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Log saved successfully", "Save Complete", "OK", "Information")
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to save log: $($_.Exception.Message)", "Save Error", "OK", "Error")
            }
        }
    })
    
    $null = $LogControlsPanel.Controls.Add($RefreshLogButton)
    $null = $LogControlsPanel.Controls.Add($ClearLogButton)
    $null = $LogControlsPanel.Controls.Add($SaveLogButton)
    
    $null = $Tab.Controls.Add($LogControlsPanel)
    $null = $Tab.Controls.Add($Script:LogTextBox)
    
    $null = $Script:TabControl.TabPages.Add($Tab)
}

# ================================
# UTILITY FUNCTIONS
# ================================

function New-SecurePassword {
    $Length = 16
    $Characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*"
    $Password = ""
    
    for ($i = 0; $i -lt $Length; $i++) {
        $Password += $Characters[(Get-Random -Maximum $Characters.Length)]
    }
    
    return $Password
}

function Clear-UserCreationForm {
    if ($Script:DisplayNameTextBox) { $Script:DisplayNameTextBox.Clear() }
    if ($Script:FirstNameTextBox) { $Script:FirstNameTextBox.Clear() }
    if ($Script:LastNameTextBox) { $Script:LastNameTextBox.Clear() }
    if ($Script:UsernameTextBox) { $Script:UsernameTextBox.Clear() }
    if ($Script:PasswordTextBox) { $Script:PasswordTextBox.Clear() }
    if ($Script:DepartmentTextBox) { $Script:DepartmentTextBox.Clear() }
    if ($Script:JobTitleTextBox) { $Script:JobTitleTextBox.Clear() }
    
    if ($Script:OfficeDropdown -and $Script:OfficeDropdown.Items.Count -gt 0) {
        $Script:OfficeDropdown.SelectedIndex = -1
    }
    
    if ($Script:ManagerDropdown -and $Script:ManagerDropdown.Items.Count -gt 0) {
        $Script:ManagerDropdown.SelectedIndex = 0
    }
    
    if ($Script:LicenseDropdown -and $Script:LicenseDropdown.Items.Count -gt 0) {
        $Script:LicenseDropdown.SelectedIndex = 0
    }
    
    if ($Script:GroupsCheckedListBox) {
        for ($i = 0; $i -lt $Script:GroupsCheckedListBox.Items.Count; $i++) {
            $Script:GroupsCheckedListBox.SetItemChecked($i, $false)
        }
    }
}

function Update-ActivityLogDisplay {
    if ($Script:LogTextBox) {
        $Script:LogTextBox.Clear()
        foreach ($LogEntry in $Global:ActivityLog) {
            $Script:LogTextBox.AppendText($LogEntry + "`n")
        }
        
        $Script:LogTextBox.SelectionStart = $Script:LogTextBox.Text.Length
        $Script:LogTextBox.ScrollToCaret()
    }
}

# ================================
# MAIN APPLICATION ENTRY POINT
# ================================

function Start-M365ProvisioningTool {
    try {
        Write-Host "Starting M365 User Provisioning Tool - Enterprise Edition 2025 (REPLICATION FIXED)" -ForegroundColor Cyan
        Write-Host "================================================================" -ForegroundColor Cyan
        Write-Host "✓ Emoji corruption fixes applied" -ForegroundColor Green
        Write-Host "✓ Exchange Online PowerShell integration RESTORED" -ForegroundColor Green
        Write-Host "✓ Azure AD replication delays FIXED with intelligent retry logic" -ForegroundColor Green
        Write-Host "✓ Second pass system for maximum success rate" -ForegroundColor Green
        Write-Host "✓ Distribution list and shared mailbox support working" -ForegroundColor Green
        Write-Host "✓ Exchange Online connects when needed for distribution lists" -ForegroundColor Green
        Write-Host "================================================================" -ForegroundColor Cyan
        
        if (-not (Initialize-Application)) {
            Write-Host "Application initialization failed. Exiting." -ForegroundColor Red
            return
        }
        
        Write-ActivityLog "M365 User Provisioning Tool starting..." "Info"
        
        $MainForm = New-MainForm
        
        Write-ActivityLog "Application interface loaded successfully" "Success"
        
        $MainForm.Add_Shown({
            $MainForm.Activate()
            Write-ActivityLog "Application ready for use" "Info"
            Update-ActivityLogDisplay
        })
        
        $MainForm.Add_FormClosing({
            param($FormSender, $e)
            
            if ($Global:IsConnected) {
                $Result = [System.Windows.Forms.MessageBox]::Show(
                    "You are still connected to Microsoft Graph. Do you want to disconnect before closing?",
                    "Confirm Exit",
                    "YesNoCancel",
                    "Question"
                )
                
                if ($Result -eq "Yes") {
                    Disconnect-FromMicrosoftGraph
                }
                elseif ($Result -eq "Cancel") {
                    $e.Cancel = $true
                    return
                }
            }
            
            Write-ActivityLog "Application closing..." "Info"
        })
        
        [System.Windows.Forms.Application]::Run($MainForm)
        
    }
    catch {
        $ErrorMsg = "Critical error in M365 Provisioning Tool: $($_.Exception.Message)"
        Write-Host $ErrorMsg -ForegroundColor Red
        Write-ActivityLog $ErrorMsg "Error"
        
        [System.Windows.Forms.MessageBox]::Show(
            $ErrorMsg + "`n`nThe application will now exit.",
            "Critical Error",
            "OK",
            "Error"
        )
    }
    finally {
        Write-Host "M365 User Provisioning Tool session ended." -ForegroundColor Yellow
        Write-Host "Thank you for using the M365 User Provisioning Tool!" -ForegroundColor Cyan
    }
}

# ================================
# APPLICATION STARTUP
# ================================

Start-M365ProvisioningTool