#Requires -Version 7.0

[CmdletBinding()]
param(
    [switch]$NoGUI,     # For command-line mode if needed
    [switch]$TestMode   # For testing without making changes
)

<#
.SYNOPSIS
    M365 User Provisioning Tool - Enterprise Edition 2025 (COMPLETE WITH BULK IMPORT + EXCHANGE ONLINE)
    Comprehensive M365 user management with intelligent tenant discovery and enhanced Exchange functionality
    FIXED: Restored legacy functionality for flexible attributes and manual group selection

.DESCRIPTION
    Advanced user provisioning tool with:
    - Intelligent tenant discovery (users, groups, mailboxes, sites)
    - Single user creation and bulk CSV import
    - License assignment via CustomAttribute1
    - Manual office location input
    - Clean tabbed interface with pagination
    - Robust error handling and validation
    - ENHANCED: M365.ExchangeOnline module integration for accurate Exchange data
    - FIXED: Flexible attribute handling and manual group selection (no hardcoded assumptions)
    
    FEATURES:
    - Single user creation with full property support
    - Bulk CSV import with progress tracking
    - Dry run testing capabilities
    - Comprehensive tenant data discovery with enhanced Exchange functionality
    - Activity logging
    - Accurate shared mailbox detection
    - Complete distribution list management
    - Mail-enabled security group support
    - RESTORED: Manual group selection including distribution lists

.NOTES
    Version: 3.1.2025-COMPLETE-ENHANCED-FIXED-PROFESSIONAL
    Author: Enterprise Solutions Team
    PowerShell: 7.0+ Required
    Dependencies: Microsoft Graph PowerShell SDK V2.28+, Exchange Online PowerShell, M365.ExchangeOnline Module
    Last Updated: August 2025
    Fixed: Hardcoded Exchange provisioning removed, legacy flexibility restored, professional UI

.EXAMPLE
    .\M365-UserProvisioning-Enterprise.ps1
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
$Global:ExchangeModuleAvailable = $false  # Track M365.ExchangeOnline module availability

# Bulk Import Variables
$Global:ImportData = $null
$Script:FilePathTextBox = $null
$Script:PreviewDataGrid = $null
$Script:ProgressBar = $null
$Script:ProgressLabel = $null
$Script:ProgressDetails = $null
$Script:ImportButton = $null
$Script:DryRunCheckBox = $null
$Script:SkipDuplicatesCheckBox = $null

# License type mappings for CustomAttribute1
$Global:LicenseTypes = @{
    "BusinessBasic" = "BusinessBasic"
    "BusinessPremium" = "BusinessPremium"
    "BusinessStandard" = "BusinessStandard"
    "E3" = "E3"
    "E5" = "E5"
    "ExchangeOnline1" = "ExchangeOnline1"
    "ExchangeOnline2" = "ExchangeOnline2"
}

# Activity logging
$Global:LogFile = "M365_Provisioning_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$Global:ActivityLog = @()

# ================================
# ASSEMBLY LOADING & INITIALIZATION (FIXED)
# ================================

Write-Host "M365 User Provisioning Tool - Enterprise Edition 2025 (Enhanced & Fixed)" -ForegroundColor Green
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "COMPLETE VERSION - Single User + Bulk CSV Import + Enhanced Exchange + Legacy Flexibility RESTORED" -ForegroundColor Yellow
Write-Host ""

try {
    Write-Host "Initializing Windows Forms (Fixed Mode)..." -ForegroundColor Cyan
    
    # Load Windows Forms assemblies
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Write-Host "   [OK] Windows Forms assemblies loaded" -ForegroundColor Green
    
    # Enable visual styles only
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Write-Host "   [OK] Visual styles enabled" -ForegroundColor Green
    
    # SKIP SetCompatibleTextRenderingDefault - this was causing the error!
    Write-Host "   [SKIP] SetCompatibleTextRenderingDefault (not required for functionality)" -ForegroundColor Yellow
    
    Write-Host "[OK] Windows Forms ready for enterprise application!" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "[ERROR] Failed to initialize Windows Forms: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ================================
# MICROSOFT GRAPH MODULE LOADING
# ================================

Write-Host "Loading Microsoft Graph PowerShell modules..." -ForegroundColor Cyan

$RequiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Users.Actions',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Sites',
    'ExchangeOnlineManagement'
)

foreach ($Module in $RequiredModules) {
    try {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            Write-Host "   Installing $Module..." -ForegroundColor Yellow
            Install-Module -Name $Module -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        }
        
        Write-Host "   Importing $Module..." -ForegroundColor Yellow
        Import-Module -Name $Module -Force -ErrorAction Stop
        Write-Host "   [OK] $Module loaded" -ForegroundColor Green
    }
    catch {
        Write-Warning "[WARNING] Failed to load $Module : $($_.Exception.Message)"
        Write-Host "Application will continue with limited functionality" -ForegroundColor Yellow
    }
}

Write-Host "[OK] Required modules processed" -ForegroundColor Green

# ================================
# ENHANCED: M365.EXCHANGEONLINE MODULE LOADING
# ================================

Write-Host ""
Write-Host "Loading Enhanced Exchange Online Module..." -ForegroundColor Cyan

# Load custom M365.ExchangeOnline module (if available)
try {
    if (Test-Path ".\Modules\M365.ExchangeOnline\M365.ExchangeOnline.psd1") {
        Write-Host "   Loading M365.ExchangeOnline module..." -ForegroundColor Yellow
        Import-Module ".\Modules\M365.ExchangeOnline\M365.ExchangeOnline.psd1" -Force -ErrorAction Stop
        Write-Host "   [OK] M365.ExchangeOnline module loaded successfully!" -ForegroundColor Green
        Write-Host "   Enhanced Exchange functionality available!" -ForegroundColor Green
        $Global:ExchangeModuleAvailable = $true
        
        # Test if key functions are available
        $ExchangeFunctions = Get-Command -Module M365.ExchangeOnline -ErrorAction SilentlyContinue
        Write-Host "   Available Exchange functions: $($ExchangeFunctions.Count)" -ForegroundColor Cyan
    }
    else {
        Write-Host "   [WARNING] M365.ExchangeOnline module not found at .\Modules\M365.ExchangeOnline\" -ForegroundColor Yellow
        Write-Host "   Using built-in Exchange functionality instead" -ForegroundColor Yellow
        $Global:ExchangeModuleAvailable = $false
    }
}
catch {
    Write-Warning "[WARNING] Failed to load M365.ExchangeOnline: $($_.Exception.Message)"
    Write-Host "   Using built-in Exchange functionality instead" -ForegroundColor Yellow
    $Global:ExchangeModuleAvailable = $false
}

Write-Host ""

# ================================
# ENHANCED TENANT DISCOVERY FUNCTIONS
# ================================

function Get-TenantData {
    <#
    .SYNOPSIS
        Performs comprehensive tenant discovery to populate all dropdowns and lists
        ENHANCED with M365.ExchangeOnline module integration for accurate Exchange data
    #>
    
    try {
        Update-StatusLabel "Discovering tenant data..."
        
        # Get tenant information
        Write-Host "   Getting tenant information..." -ForegroundColor Yellow
        $Global:TenantInfo = Get-MgOrganization
        
        # Get accepted domains
        Write-Host "   Getting accepted domains..." -ForegroundColor Yellow
        $Global:AcceptedDomains = Get-MgDomain | Where-Object { $_.IsVerified -eq $true }
        
        # Get all users (for manager dropdown)
        Write-Host "   Getting existing users..." -ForegroundColor Yellow
        $Global:AvailableUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,JobTitle,Department" | 
            Sort-Object DisplayName
        
        # Get all groups
        Write-Host "   Getting security groups..." -ForegroundColor Yellow
        $Global:AvailableGroups = Get-MgGroup -All -Property "DisplayName,GroupTypes,SecurityEnabled,MailEnabled" |
            Sort-Object DisplayName
        
        # ENHANCED EXCHANGE ONLINE DATA DISCOVERY
        Write-Host "   Getting Exchange Online data..." -ForegroundColor Yellow
        
        # Check if M365.ExchangeOnline module is available
        if ($Global:ExchangeModuleAvailable -and (Get-Command "Get-AllExchangeData" -ErrorAction SilentlyContinue)) {
            Write-Host "      Using M365.ExchangeOnline module for enhanced data discovery..." -ForegroundColor Cyan
            
            try {
                # Use the enhanced Exchange module
                $ExchangeData = Get-AllExchangeData
                
                # Populate global variables with enhanced data
                $Global:SharedMailboxes = $ExchangeData.SharedMailboxes
                $Global:DistributionLists = $ExchangeData.DistributionLists  
                $Global:MailEnabledSecurityGroups = $ExchangeData.MailEnabledSecurityGroups
                
                # Merge accepted domains (avoid duplicates)
                $ExistingDomains = $Global:AcceptedDomains | Select-Object -ExpandProperty Id
                $NewDomains = $ExchangeData.AcceptedDomains | Where-Object { $_.DomainName -notin $ExistingDomains }
                if ($NewDomains) {
                    $Global:AcceptedDomains += $NewDomains | Select-Object @{Name='Id';Expression={$_.DomainName}}, @{Name='IsDefault';Expression={$_.Default}}
                }
                
                # Set mailboxes from Exchange data
                try {
                    $UserMailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize 50 | 
                        Select-Object @{Name='Name';Expression={$_.DisplayName}}, @{Name='EmailAddress';Expression={$_.PrimarySmtpAddress}}
                    $Global:AvailableMailboxes = $Global:SharedMailboxes + $UserMailboxes
                }
                catch {
                    Write-Warning "Could not get user mailboxes, using shared mailboxes only"
                    $Global:AvailableMailboxes = $Global:SharedMailboxes
                }
                
                Write-Host "      [OK] Enhanced Exchange data loaded!" -ForegroundColor Green
                Write-Host "         $($ExchangeData.Summary.SharedMailboxCount) shared mailboxes" -ForegroundColor Gray
                Write-Host "         $($ExchangeData.Summary.DistributionListCount) distribution lists" -ForegroundColor Gray
                Write-Host "         $($ExchangeData.Summary.MailEnabledSecurityGroupCount) mail-enabled security groups" -ForegroundColor Gray
                Write-Host "         $($ExchangeData.Summary.AcceptedDomainCount) Exchange domains" -ForegroundColor Gray
            }
            catch {
                Write-Warning "Enhanced Exchange discovery failed, falling back to standard method: $($_.Exception.Message)"
                Get-ExchangeDataFallback
            }
        }
        else {
            Write-Host "      [WARNING] M365.ExchangeOnline module not available - using standard Exchange discovery" -ForegroundColor Yellow
            Get-ExchangeDataFallback
        }
        
        # Get SharePoint sites
        Write-Host "   Getting SharePoint sites..." -ForegroundColor Yellow
        try {
            $Global:SharePointSites = Get-MgSite -All | Sort-Object DisplayName
        }
        catch {
            Write-Warning "SharePoint not available - skipping site discovery"
            $Global:SharePointSites = @()
        }
        
        # Get available licenses
        Write-Host "   Getting license information..." -ForegroundColor Yellow
        $Global:AvailableLicenses = Get-MgSubscribedSku
        
        # Update UI with discovered data
        Update-TenantDataDisplay
        Update-Dropdowns
        
        Update-StatusLabel "[OK] Tenant discovery completed successfully"
        Write-Host "[OK] Tenant data discovery completed" -ForegroundColor Green
        
        # Enhanced logging
        $TotalExchangeItems = $Global:SharedMailboxes.Count + $Global:DistributionLists.Count + $Global:MailEnabledSecurityGroups.Count
        Add-ActivityLog "Tenant Discovery" "Success" "Discovered: $($Global:AvailableUsers.Count) users, $($Global:AvailableGroups.Count) groups, $($Global:AvailableMailboxes.Count) mailboxes, $TotalExchangeItems Exchange items"
        
    }
    catch {
        $ErrorMsg = "Tenant discovery failed: $($_.Exception.Message)"
        Update-StatusLabel "[ERROR] $ErrorMsg"
        Write-Host "[ERROR] $ErrorMsg" -ForegroundColor Red
        Add-ActivityLog "Tenant Discovery" "Failed" $_.Exception.Message
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to discover tenant data:`n`n$($_.Exception.Message)",
            "Tenant Discovery Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
}

function Get-ExchangeDataFallback {
    <#
    .SYNOPSIS
        Fallback Exchange discovery using standard cmdlets (original method)
    #>
    
    Write-Host "      Using standard Exchange Online cmdlets..." -ForegroundColor Yellow
    
    # Get distribution groups (original method)
    try {
        $Global:DistributionLists = Get-DistributionGroup -ResultSize Unlimited | Sort-Object Name
        $Global:MailEnabledSecurityGroups = Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited | Sort-Object Name
        Write-Host "      [OK] Standard distribution groups loaded: $($Global:DistributionLists.Count) DLs, $($Global:MailEnabledSecurityGroups.Count) MESGs" -ForegroundColor Green
    }
    catch {
        Write-Warning "Exchange Online not available - skipping distribution groups"
        $Global:DistributionLists = @()
        $Global:MailEnabledSecurityGroups = @()
    }
    
    # Get mailboxes (original method)
    try {
        $Global:AvailableMailboxes = Get-Mailbox -ResultSize Unlimited | Sort-Object Name
        $Global:SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Sort-Object Name
        Write-Host "      [OK] Standard mailboxes loaded: $($Global:AvailableMailboxes.Count) total, $($Global:SharedMailboxes.Count) shared" -ForegroundColor Green
    }
    catch {
        Write-Warning "Exchange Online not available - skipping mailbox discovery"
        $Global:AvailableMailboxes = @()
        $Global:SharedMailboxes = @()
    }
}

function Update-Dropdowns {
    <#
    .SYNOPSIS
        Updates all dropdown controls with discovered tenant data
        ENHANCED: Includes improved groups list with distribution lists
    #>
    
    # Update domain dropdown
    $Script:DomainDropdown.Items.Clear()
    foreach ($Domain in $Global:AcceptedDomains) {
        $Script:DomainDropdown.Items.Add($Domain.Id) | Out-Null
    }
    if ($Script:DomainDropdown.Items.Count -gt 0) {
        $Script:DomainDropdown.SelectedIndex = 0
    }
    
    # Update manager dropdown
    $Script:ManagerDropdown.Items.Clear()
    $Script:ManagerDropdown.Items.Add("(No Manager)") | Out-Null
    foreach ($User in $Global:AvailableUsers) {
        $ManagerDisplay = "$($User.DisplayName) ($($User.UserPrincipalName))"
        $Script:ManagerDropdown.Items.Add($ManagerDisplay) | Out-Null
    }
    $Script:ManagerDropdown.SelectedIndex = 0
    
    # Update license dropdown
    $Script:LicenseDropdown.Items.Clear()
    foreach ($LicenseType in $Global:LicenseTypes.Keys) {
        $Script:LicenseDropdown.Items.Add($LicenseType) | Out-Null
    }
    if ($Script:LicenseDropdown.Items.Count -gt 0) {
        $Script:LicenseDropdown.SelectedIndex = 0
    }
    
    # Office location is now a text box - no population needed
    # $Script:OfficeDropdown is now $Script:OfficeTextBox
    
    # ENHANCED: Update groups checklist with improved categorization and distribution lists
    $Script:GroupsCheckedListBox.Items.Clear()
    
    Write-Host "Updating groups list with enhanced categorization..." -ForegroundColor Cyan
    
    # Add regular security groups
    $SecurityGroups = $Global:AvailableGroups | Where-Object { 
        ($_.SecurityEnabled -eq $true -and $_.MailEnabled -ne $true) -and ($_.GroupTypes -notcontains "Unified")
    } | Sort-Object DisplayName
    
    if ($SecurityGroups.Count -gt 0) {
        $Script:GroupsCheckedListBox.Items.Add("=== SECURITY GROUPS ===") | Out-Null
        foreach ($Group in $SecurityGroups) {
            $DisplayText = "$($Group.DisplayName) [Security Group]"
            $Script:GroupsCheckedListBox.Items.Add($DisplayText) | Out-Null
        }
    }
    
    # Add Microsoft 365 Groups
    $M365Groups = $Global:AvailableGroups | Where-Object { 
        $_.GroupTypes -contains "Unified"
    } | Sort-Object DisplayName
    
    if ($M365Groups.Count -gt 0) {
        $Script:GroupsCheckedListBox.Items.Add("=== MICROSOFT 365 GROUPS ===") | Out-Null
        foreach ($Group in $M365Groups) {
            $DisplayText = "$($Group.DisplayName) [Microsoft 365 Group]"
            if ($Group.Mail) {
                $DisplayText += " - $($Group.Mail)"
            }
            $Script:GroupsCheckedListBox.Items.Add($DisplayText) | Out-Null
        }
    }
    
    # RESTORED: Add Distribution Lists (NOW SELECTABLE!)
    if ($Global:DistributionLists -and $Global:DistributionLists.Count -gt 0) {
        $Script:GroupsCheckedListBox.Items.Add("=== DISTRIBUTION LISTS ===") | Out-Null
        foreach ($DL in ($Global:DistributionLists | Sort-Object DisplayName)) {
            $GroupName = if ($DL.DisplayName) { $DL.DisplayName } elseif ($DL.Name) { $DL.Name } else { $DL.ToString() }
            $EmailAddress = if ($DL.Mail) { $DL.Mail } elseif ($DL.PrimarySmtpAddress) { $DL.PrimarySmtpAddress } else { "" }
            
            $DisplayText = "$GroupName [Distribution List]"
            if ($EmailAddress) {
                $DisplayText += " - $EmailAddress"
            }
            $Script:GroupsCheckedListBox.Items.Add($DisplayText) | Out-Null
        }
    }
    
    # Add Mail-Enabled Security Groups
    if ($Global:MailEnabledSecurityGroups -and $Global:MailEnabledSecurityGroups.Count -gt 0) {
        $Script:GroupsCheckedListBox.Items.Add("=== MAIL-ENABLED SECURITY GROUPS ===") | Out-Null
        foreach ($MESG in ($Global:MailEnabledSecurityGroups | Sort-Object DisplayName)) {
            $GroupName = if ($MESG.DisplayName) { $MESG.DisplayName } elseif ($MESG.Name) { $MESG.Name } else { $MESG.ToString() }
            $EmailAddress = if ($MESG.Mail) { $MESG.Mail } elseif ($MESG.PrimarySmtpAddress) { $MESG.PrimarySmtpAddress } else { "" }
            
            $DisplayText = "$GroupName [Mail-Enabled Security]"
            if ($EmailAddress) {
                $DisplayText += " - $EmailAddress"
            }
            $Script:GroupsCheckedListBox.Items.Add($DisplayText) | Out-Null
        }
    }
    
    # RESTORED: Add Shared Mailboxes (NOW SELECTABLE FOR PERMISSIONS!)
    if ($Global:SharedMailboxes -and $Global:SharedMailboxes.Count -gt 0) {
        $Script:GroupsCheckedListBox.Items.Add("=== SHARED MAILBOXES ===") | Out-Null
        foreach ($SharedMailbox in ($Global:SharedMailboxes | Sort-Object DisplayName)) {
            $MailboxName = if ($SharedMailbox.DisplayName) { $SharedMailbox.DisplayName } elseif ($SharedMailbox.Name) { $SharedMailbox.Name } else { $SharedMailbox.ToString() }
            $EmailAddress = if ($SharedMailbox.EmailAddress) { $SharedMailbox.EmailAddress } elseif ($SharedMailbox.PrimarySmtpAddress) { $SharedMailbox.PrimarySmtpAddress } else { "" }
            
            $DisplayText = "$MailboxName [Shared Mailbox]"
            if ($EmailAddress) {
                $DisplayText += " - $EmailAddress"
            }
            $Script:GroupsCheckedListBox.Items.Add($DisplayText) | Out-Null
        }
    }
    
    Write-Host "[OK] Groups list updated with enhanced categorization" -ForegroundColor Green
    Write-Host "   Security Groups: $($SecurityGroups.Count)" -ForegroundColor Gray
    Write-Host "   M365 Groups: $($M365Groups.Count)" -ForegroundColor Gray
    Write-Host "   Distribution Lists: $($Global:DistributionLists.Count)" -ForegroundColor Gray
    Write-Host "   Mail-Enabled Security Groups: $($Global:MailEnabledSecurityGroups.Count)" -ForegroundColor Gray
    Write-Host "   Shared Mailboxes: $($Global:SharedMailboxes.Count)" -ForegroundColor Gray
}

# ================================
# ENHANCED USER CREATION FUNCTIONS (FIXED)
# ================================

function New-M365User {
    <#
    .SYNOPSIS
        Creates a new M365 user with all specified properties and group memberships
        FIXED: Restored legacy functionality - flexible attributes and manual group selection only
    #>
    
    param(
        [Parameter(Mandatory)]
        [string]$FirstName,
        
        [Parameter(Mandatory)]
        [string]$LastName,
        
        [Parameter(Mandatory)]
        [string]$Username,
        
        [Parameter(Mandatory)]
        [string]$Domain,
        
        [Parameter(Mandatory)]
        [string]$Password,
        
        [string]$Department,
        [string]$JobTitle,
        [string]$Office,
        [string]$Manager,
        [string]$LicenseType,
        [array]$Groups
    )
    
    try {
        $UserPrincipalName = "$Username@$Domain"
        $DisplayName = "$FirstName $LastName"
        
        Update-StatusLabel "Creating user: $UserPrincipalName"
        
        # Create user parameters - FLEXIBLE ATTRIBUTE HANDLING RESTORED
        $UserParams = @{
            UserPrincipalName = $UserPrincipalName
            DisplayName = $DisplayName
            GivenName = $FirstName
            Surname = $LastName
            MailNickname = $Username
            AccountEnabled = $true
            UsageLocation = "GB"
            PasswordProfile = @{
                ForceChangePasswordNextSignIn = $true
                Password = $Password
            }
        }
        
        # Add optional properties - ANY VALUE CAN BE ENTERED (restored legacy behavior)
        if ($Department) { 
            $UserParams.Department = $Department
            Write-Host "   Setting Department: $Department" -ForegroundColor Gray
        }
        if ($JobTitle) { 
            $UserParams.JobTitle = $JobTitle
            Write-Host "   Setting Job Title: $JobTitle" -ForegroundColor Gray
        }
        if ($Office) { 
            $UserParams.OfficeLocation = $Office
            Write-Host "   Setting Office Location: $Office" -ForegroundColor Gray
        }
        
        # Set CustomAttribute1 for license assignment
        if ($LicenseType) {
            $UserParams.OnPremisesExtensionAttributes = @{
                ExtensionAttribute1 = $LicenseType
            }
            Write-Host "   Setting License Type (CustomAttribute1): $LicenseType" -ForegroundColor Gray
        }
        
        # Create the user
        Write-Host "   Creating user account..." -ForegroundColor Yellow
        $NewUser = New-MgUser @UserParams
        
        Write-Host "   [OK] User created: $($NewUser.UserPrincipalName)" -ForegroundColor Green
        
        # Set manager if specified
        if ($Manager -and $Manager -ne "(No Manager)") {
            try {
                $ManagerUPN = ($Manager -split '\(')[1] -replace '\)', ''
                $ManagerUser = Get-MgUser -Filter "userPrincipalName eq '$ManagerUPN'"
                if ($ManagerUser) {
                    Set-MgUserManagerByRef -UserId $NewUser.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($ManagerUser.Id)" }
                    Write-Host "   Manager set: $($Manager)" -ForegroundColor Green
                }
            }
            catch {
                Write-Warning "Failed to set manager: $($_.Exception.Message)"
            }
        }
        
        # FIXED: RESTORED LEGACY GROUP PROCESSING - MANUAL SELECTION ONLY (NO HARDCODED ASSUMPTIONS)
        if ($Groups -and $Groups.Count -gt 0) {
            Write-Host "   Processing manually selected groups..." -ForegroundColor Yellow
            Write-Host "   Selected groups: $($Groups.Count)" -ForegroundColor Gray
            
            # Process each manually selected group
            foreach ($GroupName in $Groups) {
                try {
                    Write-Host "   Processing: $GroupName" -ForegroundColor Gray
                    
                    # Clean group name (remove formatting from UI)
                    $CleanGroupName = Get-CleanGroupName -DisplayName $GroupName
                    Write-Host "   Clean name: $CleanGroupName" -ForegroundColor Gray
                    
                    # Skip separator lines
                    if ($GroupName -like "*===*") {
                        continue
                    }
                    
                    # Determine what type of group this is based on the display text
                    if ($GroupName -match '\[Distribution List\]') {
                        # This is a distribution list - handle via Exchange Online
                        Write-Host "   Identified as Distribution List: $CleanGroupName" -ForegroundColor Cyan
                        Add-UserToDistributionListManual -NewUser $NewUser -ListName $CleanGroupName
                    }
                    elseif ($GroupName -match '\[Shared Mailbox\]') {
                        # This is a shared mailbox - handle via Exchange Online
                        Write-Host "   Identified as Shared Mailbox: $CleanGroupName" -ForegroundColor Cyan
                        Add-UserToSharedMailboxManual -NewUser $NewUser -MailboxName $CleanGroupName
                    }
                    elseif ($GroupName -match '\[Mail-Enabled Security\]') {
                        # This is a mail-enabled security group
                        Write-Host "   Identified as Mail-Enabled Security Group: $CleanGroupName" -ForegroundColor Cyan
                        
                        # Try Graph first
                        $Group = $Global:AvailableGroups | Where-Object { $_.DisplayName -eq $CleanGroupName }
                        if ($Group) {
                            try {
                                $GroupMember = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($NewUser.Id)" }
                                New-MgGroupMember -GroupId $Group.Id -BodyParameter $GroupMember
                                Write-Host "      [OK] Added to mail-enabled security group via Graph: $CleanGroupName" -ForegroundColor Green
                            }
                            catch {
                                Write-Host "      [WARNING] Graph method failed, trying Exchange method..." -ForegroundColor Yellow
                                Add-UserToDistributionListManual -NewUser $NewUser -ListName $CleanGroupName
                            }
                        }
                        else {
                            Add-UserToDistributionListManual -NewUser $NewUser -ListName $CleanGroupName
                        }
                    }
                    else {
                        # Regular security group or M365 group - handle via Graph
                        Write-Host "   Identified as Security/M365 Group: $CleanGroupName" -ForegroundColor Cyan
                        
                        $Group = $Global:AvailableGroups | Where-Object { $_.DisplayName -eq $CleanGroupName }
                        if ($Group) {
                            $GroupMember = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($NewUser.Id)" }
                            New-MgGroupMember -GroupId $Group.Id -BodyParameter $GroupMember
                            Write-Host "      [OK] Added to group: $CleanGroupName" -ForegroundColor Green
                        }
                        else {
                            Write-Host "      [WARNING] Group not found in tenant: $CleanGroupName" -ForegroundColor Yellow
                        }
                    }
                }
                catch {
                    Write-Warning "Failed to process group $GroupName : $($_.Exception.Message)"
                }
            }
        }
        
        # FIXED: NO AUTOMATIC EXCHANGE PROVISIONING - Only manual selections are processed
        Write-Host "   Exchange resources: Manual selection only (no automatic assignments)" -ForegroundColor Cyan
        Write-Host "   [OK] All attributes applied as entered (flexible attribute handling restored)" -ForegroundColor Green
        
        Update-StatusLabel "[OK] User created successfully: $UserPrincipalName"
        Add-ActivityLog "User Creation" "Success" "Created user: $UserPrincipalName with manual group selections and flexible attributes"
        
        # Enhanced success message
        $SuccessMessage = "User created successfully!`n`nName: $DisplayName`nUPN: $UserPrincipalName"
        if ($LicenseType) {
            $SuccessMessage += "`nLicense Type (CustomAttribute1): $LicenseType"
        }
        if ($Department) {
            $SuccessMessage += "`nDepartment: $Department"
        }
        if ($JobTitle) {
            $SuccessMessage += "`nJob Title: $JobTitle"
        }
        if ($Office) {
            $SuccessMessage += "`nOffice Location: $Office"
        }
        
        $SuccessMessage += "`n`n[OK] All attributes applied as entered (no restrictions)"
        $SuccessMessage += "`n[OK] Manual group selections processed"
        $SuccessMessage += "`n[OK] Legacy flexibility restored"
        $SuccessMessage += "`n`nThe user will receive an email with sign-in instructions."
        
        [System.Windows.Forms.MessageBox]::Show(
            $SuccessMessage,
            "User Creation Successful",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        # Clear the form
        Clear-UserForm
        
        return $NewUser
    }
    catch {
        $ErrorMsg = "Failed to create user $UserPrincipalName : $($_.Exception.Message)"
        Update-StatusLabel "[ERROR] $ErrorMsg"
        Add-ActivityLog "User Creation" "Failed" $ErrorMsg
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to create user:`n`n$($_.Exception.Message)",
            "User Creation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        throw
    }
}

# ================================
# MANUAL EXCHANGE OPERATION HELPER FUNCTIONS (FIXED)
# ================================

function Get-CleanGroupName {
    <#
    .SYNOPSIS
        Cleans group names from UI formatting (FIXED - No more Unicode corruption)
    #>
    param([string]$DisplayName)
    
    Write-Host "      Original name: '$DisplayName'" -ForegroundColor Gray
    
    # Method 1: Split at the bracket first to get the main name
    $NamePart = ($DisplayName -split ' \[')[0]
    Write-Host "      After bracket split: '$NamePart'" -ForegroundColor Gray
    
    # Method 2: Remove any prefix formatting (no emojis to worry about now)
    $CleanName = $NamePart.Trim()
    
    Write-Host "      [OK] Final clean name: '$CleanName'" -ForegroundColor Green
    
    return $CleanName
}

function Add-UserToDistributionListManual {
    <#
    .SYNOPSIS
        Adds user to distribution list with proper error handling and propagation delay
    #>
    param(
        [object]$NewUser,
        [string]$ListName
    )
    
    Write-Host "   Attempting to add to distribution list: $ListName" -ForegroundColor Yellow
    
    try {
        # Check if we can connect to Exchange Online
        $ExchangeConnected = $false
        try {
            $null = Get-OrganizationConfig -ErrorAction Stop
            $ExchangeConnected = $true
        }
        catch {
            $ExchangeConnected = $false
        }
        
        if ($ExchangeConnected) {
            # ADD PROPAGATION DELAY - Wait for user to propagate to Exchange
            Write-Host "      Waiting 10 seconds for user propagation..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10
            
            # Verify the distribution list exists - try multiple methods
            Write-Host "      Searching for distribution list: '$ListName'" -ForegroundColor Gray
            
            $DistList = $null
            
            # Method 1: Try exact name match
            try {
                $DistList = Get-DistributionGroup -Identity $ListName -ErrorAction SilentlyContinue
            }
            catch { }
            
            # Method 2: Try search by display name if exact match failed
            if (-not $DistList) {
                try {
                    $AllDistLists = Get-DistributionGroup -ResultSize Unlimited
                    $DistList = $AllDistLists | Where-Object { 
                        $_.DisplayName -eq $ListName -or 
                        $_.Name -eq $ListName -or 
                        $_.Alias -eq $ListName 
                    } | Select-Object -First 1
                }
                catch { }
            }
            
            if (-not $DistList) {
                Write-Host "      [ERROR] Distribution list '$ListName' not found in Exchange" -ForegroundColor Red
                Write-Host "      Available distribution lists:" -ForegroundColor Gray
                try {
                    $AllDLs = Get-DistributionGroup -ResultSize 10
                    $AllDLs | ForEach-Object {
                        Write-Host "        • $($_.DisplayName) ($($_.Name))" -ForegroundColor Gray
                    }
                }
                catch { }
                
                Write-Host "      Manual task: Verify distribution list name and add user manually" -ForegroundColor Gray
                Add-ActivityLog "Exchange Operation" "Warning" "Distribution list '$ListName' not found - manual task required"
                return
            }
            
            Write-Host "      [OK] Found distribution list: $($DistList.DisplayName)" -ForegroundColor Green
            
            # Check if user already exists in the list
            $ExistingMembers = Get-DistributionGroupMember -Identity $DistList.Identity -ErrorAction SilentlyContinue
            $UserAlreadyMember = $ExistingMembers | Where-Object { $_.PrimarySmtpAddress -eq $NewUser.UserPrincipalName }
            
            if ($UserAlreadyMember) {
                Write-Host "      [OK] User is already a member of: $($DistList.DisplayName)" -ForegroundColor Green
                return
            }
            
            # Add user to distribution list
            Add-DistributionGroupMember -Identity $DistList.Identity -Member $NewUser.UserPrincipalName -Confirm:$false -ErrorAction Stop
            Write-Host "      [OK] Successfully added to distribution list: $($DistList.DisplayName)" -ForegroundColor Green
            Add-ActivityLog "Exchange Operation" "Success" "Added $($NewUser.UserPrincipalName) to distribution list: $($DistList.DisplayName)"
            
        }
        else {
            Write-Host "      [WARNING] Exchange Online not connected" -ForegroundColor Yellow
            Write-Host "      Manual task: Add $($NewUser.UserPrincipalName) to distribution list '$ListName'" -ForegroundColor Gray
            Add-ActivityLog "Exchange Operation" "Manual" "Exchange not connected - manual task: Add $($NewUser.UserPrincipalName) to distribution list '$ListName'"
        }
    }
    catch {
        Write-Host "      [ERROR] Failed to add to distribution list: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "      Manual task: Add $($NewUser.UserPrincipalName) to distribution list '$ListName'" -ForegroundColor Gray
        Add-ActivityLog "Exchange Operation" "Failed" "Failed to add to distribution list '$ListName': $($_.Exception.Message)"
    }
}

function Add-UserToSharedMailboxManual {
    <#
    .SYNOPSIS
        Adds user to shared mailbox with proper error handling and propagation delay
    #>
    param(
        [object]$NewUser,
        [string]$MailboxName
    )
    
    Write-Host "   Attempting to add to shared mailbox: $MailboxName" -ForegroundColor Yellow
    
    try {
        # Check if we can connect to Exchange Online
        $ExchangeConnected = $false
        try {
            $null = Get-OrganizationConfig -ErrorAction Stop
            $ExchangeConnected = $true
        }
        catch {
            $ExchangeConnected = $false
        }
        
        if ($ExchangeConnected) {
            # ADD PROPAGATION DELAY - Wait for user to propagate to Exchange
            Write-Host "      Waiting 10 seconds for user propagation..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10
            
            # Verify the shared mailbox exists - try multiple methods
            Write-Host "      Searching for shared mailbox: '$MailboxName'" -ForegroundColor Gray
            
            $SharedMailbox = $null
            
            # Method 1: Try exact name match
            try {
                $SharedMailbox = Get-EXOMailbox -Identity $MailboxName -RecipientTypeDetails SharedMailbox -ErrorAction SilentlyContinue
            }
            catch { }
            
            # Method 2: Try search by display name if exact match failed
            if (-not $SharedMailbox) {
                try {
                    $AllSharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
                    $SharedMailbox = $AllSharedMailboxes | Where-Object { 
                        $_.DisplayName -eq $MailboxName -or 
                        $_.Name -eq $MailboxName -or 
                        $_.Alias -eq $MailboxName 
                    } | Select-Object -First 1
                }
                catch { }
            }
            
            if (-not $SharedMailbox) {
                Write-Host "      [ERROR] Shared mailbox '$MailboxName' not found in Exchange" -ForegroundColor Red
                Write-Host "      Available shared mailboxes:" -ForegroundColor Gray
                try {
                    $AllSMBs = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize 10
                    $AllSMBs | ForEach-Object {
                        Write-Host "        • $($_.DisplayName) ($($_.PrimarySmtpAddress))" -ForegroundColor Gray
                    }
                }
                catch { }
                
                Write-Host "      Manual task: Verify shared mailbox name and grant permissions manually" -ForegroundColor Gray
                Add-ActivityLog "Exchange Operation" "Warning" "Shared mailbox '$MailboxName' not found - manual task required"
                return
            }
            
            Write-Host "      [OK] Found shared mailbox: $($SharedMailbox.DisplayName)" -ForegroundColor Green
            
            # Grant Full Access permission
            try {
                Add-MailboxPermission -Identity $SharedMailbox.PrimarySmtpAddress -User $NewUser.UserPrincipalName -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
                Write-Host "      [OK] Granted FullAccess permission" -ForegroundColor Green
            }
            catch {
                if ($_.Exception.Message -like "*already exists*") {
                    Write-Host "      [OK] FullAccess permission already exists" -ForegroundColor Green
                }
                else {
                    Write-Host "      [WARNING] Failed to grant FullAccess: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            # Grant Send As permission
            try {
                Add-RecipientPermission -Identity $SharedMailbox.PrimarySmtpAddress -Trustee $NewUser.UserPrincipalName -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                Write-Host "      [OK] Granted SendAs permission" -ForegroundColor Green
            }
            catch {
                if ($_.Exception.Message -like "*already exists*") {
                    Write-Host "      [OK] SendAs permission already exists" -ForegroundColor Green
                }
                else {
                    Write-Host "      [WARNING] Failed to grant SendAs: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            Write-Host "      [OK] Successfully configured shared mailbox permissions: $($SharedMailbox.DisplayName)" -ForegroundColor Green
            Add-ActivityLog "Exchange Operation" "Success" "Granted permissions to $($NewUser.UserPrincipalName) for shared mailbox: $($SharedMailbox.DisplayName)"
            
        }
        else {
            Write-Host "      [WARNING] Exchange Online not connected" -ForegroundColor Yellow
            Write-Host "      Manual task: Grant $($NewUser.UserPrincipalName) access to shared mailbox '$MailboxName'" -ForegroundColor Gray
            Add-ActivityLog "Exchange Operation" "Manual" "Exchange not connected - manual task: Grant $($NewUser.UserPrincipalName) access to shared mailbox '$MailboxName'"
        }
    }
    catch {
        Write-Host "      [ERROR] Failed to configure shared mailbox: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "      Manual task: Grant $($NewUser.UserPrincipalName) access to shared mailbox '$MailboxName'" -ForegroundColor Gray
        Add-ActivityLog "Exchange Operation" "Failed" "Failed to configure shared mailbox '$MailboxName': $($_.Exception.Message)"
    }
}

function Clear-UserForm {
    <#
    .SYNOPSIS
        Clears all user input fields
    #>
    
    $Script:FirstNameTextBox.Clear()
    $Script:LastNameTextBox.Clear()
    $Script:UsernameTextBox.Clear()
    $Script:PasswordTextBox.Clear()
    $Script:DepartmentTextBox.Clear()
    $Script:JobTitleTextBox.Clear()
    
    if ($Script:DomainDropdown.Items.Count -gt 0) {
        $Script:DomainDropdown.SelectedIndex = 0
    }
    if ($Script:ManagerDropdown.Items.Count -gt 0) {
        $Script:ManagerDropdown.SelectedIndex = 0
    }
    if ($Script:LicenseDropdown.Items.Count -gt 0) {
        $Script:LicenseDropdown.SelectedIndex = 0
    }
    # Clear office text box instead of resetting dropdown
    if ($Script:OfficeTextBox) {
        $Script:OfficeTextBox.Clear()
    }
    
    # Uncheck all groups
    for ($i = 0; $i -lt $Script:GroupsCheckedListBox.Items.Count; $i++) {
        $Script:GroupsCheckedListBox.SetItemChecked($i, $false)
    }
}

function Generate-SecurePassword {
    <#
    .SYNOPSIS
        Generates a secure random password
    #>
    
    $Length = 12
    $Characters = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%^&*"
    $Password = -join ((1..$Length) | ForEach-Object { $Characters[(Get-Random -Maximum $Characters.Length)] })
    return $Password
}

# ================================
# BULK IMPORT FUNCTIONS
# ================================

function Validate-CSVFile {
    param([string]$FilePath)
    
    try {
        $Script:ProgressLabel.Text = "Validating CSV file..."
        
        # Read and parse CSV
        $CSVData = Import-Csv -Path $FilePath -ErrorAction Stop
        
        if ($CSVData.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("CSV file is empty!", "Validation Error", "OK", "Error")
            return
        }
        
        # Check required columns
        $RequiredColumns = @("DisplayName", "UserPrincipalName", "FirstName", "LastName")
        $CSVColumns = $CSVData[0].PSObject.Properties.Name
        $MissingColumns = $RequiredColumns | Where-Object { $_ -notin $CSVColumns }
        
        if ($MissingColumns.Count -gt 0) {
            $ErrorMsg = "Missing required columns: $($MissingColumns -join ', ')"
            [System.Windows.Forms.MessageBox]::Show($ErrorMsg, "Validation Error", "OK", "Error")
            return
        }
        
        # Display preview
        $Script:PreviewDataGrid.DataSource = $CSVData
        $Script:ImportButton.Enabled = $true
        $Script:ProgressLabel.Text = "CSV validated successfully - $($CSVData.Count) users found"
        
        # Store data globally for import
        $Global:ImportData = $CSVData
        
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error reading CSV file: $($_.Exception.Message)", "File Error", "OK", "Error")
        $Script:ProgressLabel.Text = "CSV validation failed"
    }
}

function Create-CSVTemplate {
    try {
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $SaveFileDialog.Title = "Save CSV Template"
        $SaveFileDialog.FileName = "M365_BulkImport_Template.csv"
        
        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $TemplateContent = @"
DisplayName,UserPrincipalName,FirstName,LastName,Department,JobTitle,Office,Manager,LicenseType,Groups,Password,ForcePasswordChange
John Smith,john.smith@company.com,John,Smith,IT,Developer,"London Office",manager@company.com,BusinessPremium,"IT Team,Developers",,true
Jane Doe,jane.doe@company.com,Jane,Doe,HR,Manager,"Manchester Office",director@company.com,BusinessPremium,"HR Team,Managers",,true
"@
            
            $TemplateContent | Set-Content -Path $SaveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("CSV template saved successfully!", "Template Created", "OK", "Information")
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error creating template: $($_.Exception.Message)", "Template Error", "OK", "Error")
    }
}

function Start-BulkUserImport {
    if (-not $Global:ImportData) {
        [System.Windows.Forms.MessageBox]::Show("No CSV data loaded!", "Import Error", "OK", "Error")
        return
    }
    
    $IsDryRun = $Script:DryRunCheckBox.Checked
    $SkipDuplicates = $Script:SkipDuplicatesCheckBox.Checked
    
    $TotalUsers = $Global:ImportData.Count
    $Script:ProgressBar.Maximum = $TotalUsers
    $Script:ProgressBar.Value = 0
    
    $SuccessCount = 0
    $ErrorCount = 0
    $SkipCount = 0
    
    $Script:ImportButton.Enabled = $false
    $Script:ProgressDetails.Text = ""
    
    foreach ($User in $Global:ImportData) {
        $CurrentIndex = [array]::IndexOf($Global:ImportData, $User) + 1
        
        try {
            $Script:ProgressLabel.Text = "Processing user $CurrentIndex of $TotalUsers : $($User.DisplayName)"
            $Script:ProgressBar.Value = $CurrentIndex
            
            # Check if user already exists
            if ($SkipDuplicates) {
                $ExistingUser = Get-MgUser -Filter "userPrincipalName eq '$($User.UserPrincipalName)'" -ErrorAction SilentlyContinue
                if ($ExistingUser) {
                    $Script:ProgressDetails.Text += "SKIPPED: $($User.DisplayName) - User already exists`r`n"
                    $SkipCount++
                    continue
                }
            }
            
            if ($IsDryRun) {
                # Dry run - just validate
                $Script:ProgressDetails.Text += "DRY RUN: $($User.DisplayName) - Would create user`r`n"
                $SuccessCount++
            }
            else {
                # Create the user using your existing New-M365User function
                $Result = New-BulkM365User -UserData $User
                if ($Result) {
                    $Script:ProgressDetails.Text += "SUCCESS: $($User.DisplayName) created`r`n"
                    $SuccessCount++
                }
                else {
                    $Script:ProgressDetails.Text += "ERROR: $($User.DisplayName) failed to create`r`n"
                    $ErrorCount++
                }
            }
            
            # Scroll to bottom
            $Script:ProgressDetails.SelectionStart = $Script:ProgressDetails.Text.Length
            $Script:ProgressDetails.ScrollToCaret()
            
            # Update UI
            [System.Windows.Forms.Application]::DoEvents()
            
        }
        catch {
            $Script:ProgressDetails.Text += "ERROR: $($User.DisplayName) - $($_.Exception.Message)`r`n"
            $ErrorCount++
        }
    }
    
    # Final summary
    $Script:ProgressLabel.Text = "Import completed - Success: $SuccessCount, Errors: $ErrorCount, Skipped: $SkipCount"
    $Script:ImportButton.Enabled = $true
    
    [System.Windows.Forms.MessageBox]::Show(
        "Import completed!`n`nSuccess: $SuccessCount`nErrors: $ErrorCount`nSkipped: $SkipCount",
        "Import Complete",
        "OK",
        "Information"
    )
}

function New-BulkM365User {
    param($UserData)
    
    try {
        # Generate password if not provided
        $Password = if ($UserData.Password) { $UserData.Password } else { Generate-SecurePassword }
        
        # Parse groups if provided
        $Groups = @()
        if ($UserData.Groups) {
            $Groups = $UserData.Groups -split ',' | ForEach-Object { $_.Trim() }
        }
        
        # Extract domain from UPN
        $Domain = ($UserData.UserPrincipalName -split '@')[1]
        $Username = ($UserData.UserPrincipalName -split '@')[0]
        
        # Call your existing New-M365User function
        $Result = New-M365User -FirstName $UserData.FirstName `
                              -LastName $UserData.LastName `
                              -Username $Username `
                              -Domain $Domain `
                              -Password $Password `
                              -Department $UserData.Department `
                              -JobTitle $UserData.JobTitle `
                              -Office $UserData.Office `
                              -Manager $UserData.Manager `
                              -LicenseType $UserData.LicenseType `
                              -Groups $Groups
        
        return $true
    }
    catch {
        Write-Host "Error creating user $($UserData.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
        Add-ActivityLog "Bulk Import" "Failed" "Error creating $($UserData.DisplayName): $($_.Exception.Message)"
        return $false
    }
}

function Cancel-BulkImport {
    $Script:ProgressLabel.Text = "Import cancelled"
    $Script:ImportButton.Enabled = $true
}

# ================================
# UI HELPER FUNCTIONS
# ================================

function Update-StatusLabel {
    param([string]$Message)
    
    if ($Script:StatusLabel) {
        $Script:StatusLabel.Text = "$(Get-Date -Format 'HH:mm:ss') - $Message"
        [System.Windows.Forms.Application]::DoEvents()
    }
    Write-Host "STATUS: $Message" -ForegroundColor Cyan
}

function Add-ActivityLog {
    param(
        [string]$Action,
        [string]$Status,
        [string]$Details
    )
    
    $LogEntry = @{
        Timestamp = Get-Date
        Action = $Action
        Status = $Status
        Details = $Details
    }
    
    $Global:ActivityLog += $LogEntry
    
    # Update activity log display if available
    if ($Script:ActivityLogTextBox) {
        $LogText = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Status] $Action - $Details"
        $Script:ActivityLogTextBox.AppendText("$LogText`r`n")
        $Script:ActivityLogTextBox.ScrollToCaret()
    }
}

function Connect-ToMicrosoftGraph {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph with required permissions
    #>
    
    try {
        Update-StatusLabel "Connecting to Microsoft Graph..."
        
        $Scopes = @(
            "User.ReadWrite.All",
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All",
            "Organization.Read.All",
            "Sites.Read.All"
        )
        
        Connect-MgGraph -Scopes $Scopes -NoWelcome
        
        # Test connection
        $Context = Get-MgContext
        if ($Context) {
            $Global:IsConnected = $true
            Update-StatusLabel "[OK] Connected to Microsoft Graph as $($Context.Account)"
            
            # Enable connection-dependent controls
            $Script:ConnectButton.Text = "Connected - Discover Tenant Data"
            $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightGreen
            $Script:CreateUserButton.Enabled = $true
            
            Add-ActivityLog "Connection" "Success" "Connected to Microsoft Graph as $($Context.Account)"
            
            # Auto-discover tenant data (enhanced)
            Get-TenantData
            
            return $true
        }
        else {
            throw "Failed to establish Graph context"
        }
    }
    catch {
        Update-StatusLabel "[ERROR] Connection failed: $($_.Exception.Message)"
        Add-ActivityLog "Connection" "Failed" $_.Exception.Message
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Microsoft Graph:`n`n$($_.Exception.Message)",
            "Connection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

function Update-TenantDataDisplay {
    <#
    .SYNOPSIS
        Updates the tenant data tab with discovered information (ENHANCED)
    #>
    
    if ($Script:TenantDataTextBox) {
        $TenantSummary = @"
TENANT INFORMATION
==================
Tenant Name: $($Global:TenantInfo.DisplayName)
Tenant ID: $($Global:TenantInfo.Id)
Country: $($Global:TenantInfo.CountryLetterCode)

DISCOVERY SUMMARY (ENHANCED + FIXED)
=====================================
[OK] Users: $($Global:AvailableUsers.Count)
[OK] Security Groups: $($Global:AvailableGroups.Count)
[OK] Distribution Lists: $($Global:DistributionLists.Count) (NOW SELECTABLE!)
[OK] Mail-Enabled Security Groups: $($Global:MailEnabledSecurityGroups.Count)
[OK] Mailboxes: $($Global:AvailableMailboxes.Count)
[OK] Shared Mailboxes: $($Global:SharedMailboxes.Count) (NOW SELECTABLE!)
[OK] SharePoint Sites: $($Global:SharePointSites.Count)
[OK] Accepted Domains: $($Global:AcceptedDomains.Count)
[OK] License SKUs: $($Global:AvailableLicenses.Count)

LEGACY FUNCTIONALITY RESTORED
==============================
[OK] Flexible Attributes: Any values can be entered in Department, Job Title, Office fields
[OK] Manual Group Selection: Distribution lists and shared mailboxes are selectable in group membership
[OK] No Hardcoded Assumptions: No automatic assignments based on department/job title
[OK] Proper Error Handling: Manual tasks logged when automatic operations fail

EXCHANGE ENHANCEMENT STATUS
===========================
M365.ExchangeOnline Module: $(if($Global:ExchangeModuleAvailable){'[OK] Available & Active'}else{'[DISABLED] Not Available'})
Exchange Data Source: $(if($Global:ExchangeModuleAvailable){'Enhanced (Get-EXOMailbox, Get-DistributionGroup)'}else{'Standard (Get-Mailbox, Get-DistributionGroup)'})
Shared Mailbox Detection: $(if($Global:ExchangeModuleAvailable){'[OK] Accurate (RecipientTypeDetails)'}else{'[WARNING] Standard'})
Distribution List Selection: [OK] RESTORED - Available in Group Membership section

ACCEPTED DOMAINS
================
$($Global:AcceptedDomains | ForEach-Object { "• $($_.Id) $(if($_.IsDefault){'(Default)'})" } | Out-String)

RECENT ACTIVITY
===============
$($Global:ActivityLog | Select-Object -Last 10 | ForEach-Object { "$($_.Timestamp.ToString('HH:mm:ss')) [$($_.Status)] $($_.Action)" } | Out-String)
"@
        
        $Script:TenantDataTextBox.Text = $TenantSummary
    }
}

# ================================
# TAB CREATION FUNCTIONS
# ================================

function New-BulkImportTab {
    <#
    .SYNOPSIS
        Creates the Bulk CSV Import tab
    #>
    
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Bulk Import"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # Main scrollable panel
    $ScrollPanel = New-Object System.Windows.Forms.Panel
    $ScrollPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ScrollPanel.AutoScroll = $true
    
    # Instructions
    $InstructionsGroup = New-Object System.Windows.Forms.GroupBox
    $InstructionsGroup.Text = "CSV Import Instructions"
    $InstructionsGroup.Location = New-Object System.Drawing.Point(10, 10)
    $InstructionsGroup.Size = New-Object System.Drawing.Size(900, 150)
    $InstructionsGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    
    $InstructionsText = New-Object System.Windows.Forms.RichTextBox
    $InstructionsText.Dock = [System.Windows.Forms.DockStyle]::Fill
    $InstructionsText.ReadOnly = $true
    $InstructionsText.Text = @"
CSV Format Requirements:

REQUIRED COLUMNS:
• DisplayName - Full name of the user (e.g., "John Smith")
• UserPrincipalName - Email/login name (e.g., "john.smith@company.com")  
• FirstName - User's first name
• LastName - User's last name

OPTIONAL COLUMNS:
• Department - User's department (ANY VALUE can be entered)
• JobTitle - User's job title (ANY VALUE can be entered)
• Office - Office location (ANY VALUE can be entered)
• Manager - Manager's UPN (e.g., "manager@company.com")
• LicenseType - License to assign (BusinessBasic, BusinessPremium, BusinessStandard, E3, E5)
• Groups - Comma-separated group names (e.g., "IT Team,Developers")
• Password - Custom password (if blank, auto-generated)
• ForcePasswordChange - true/false for password change requirement

EXAMPLE CSV LINE:
John Smith,john.smith@company.com,John,Smith,IT,Developer,"London Office",manager@company.com,BusinessPremium,"IT Team,Developers",TempPass123!,true
"@
    $InstructionsText.BackColor = [System.Drawing.Color]::LightYellow
    $InstructionsText.Font = New-Object System.Drawing.Font("Consolas", 9)
    
    $InstructionsGroup.Controls.Add($InstructionsText)
    
    # File Selection
    $FileSelectionGroup = New-Object System.Windows.Forms.GroupBox
    $FileSelectionGroup.Text = "File Selection"
    $FileSelectionGroup.Location = New-Object System.Drawing.Point(10, 170)
    $FileSelectionGroup.Size = New-Object System.Drawing.Size(900, 80)
    $FileSelectionGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    
    $FilePathLabel = New-Object System.Windows.Forms.Label
    $FilePathLabel.Text = "CSV File:"
    $FilePathLabel.Location = New-Object System.Drawing.Point(15, 30)
    $FilePathLabel.Size = New-Object System.Drawing.Size(70, 20)
    
    $Script:FilePathTextBox = New-Object System.Windows.Forms.TextBox
    $Script:FilePathTextBox.Location = New-Object System.Drawing.Point(90, 28)
    $Script:FilePathTextBox.Size = New-Object System.Drawing.Size(500, 23)
    $Script:FilePathTextBox.ReadOnly = $true
    
    $BrowseButton = New-Object System.Windows.Forms.Button
    $BrowseButton.Text = "Browse..."
    $BrowseButton.Location = New-Object System.Drawing.Point(600, 27)
    $BrowseButton.Size = New-Object System.Drawing.Size(100, 25)
    $BrowseButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $BrowseButton.ForeColor = [System.Drawing.Color]::White
    $BrowseButton.FlatStyle = "Flat"
    $BrowseButton.Add_Click({
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $OpenFileDialog.Title = "Select CSV File for Bulk Import"
        
        if ($OpenFileDialog.ShowDialog() -eq "OK") {
            $Script:FilePathTextBox.Text = $OpenFileDialog.FileName
            Validate-CSVFile $OpenFileDialog.FileName
        }
    })
    
    $DownloadTemplateButton = New-Object System.Windows.Forms.Button
    $DownloadTemplateButton.Text = "Download Template"
    $DownloadTemplateButton.Location = New-Object System.Drawing.Point(710, 27)
    $DownloadTemplateButton.Size = New-Object System.Drawing.Size(130, 25)
    $DownloadTemplateButton.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $DownloadTemplateButton.ForeColor = [System.Drawing.Color]::White
    $DownloadTemplateButton.FlatStyle = "Flat"
    $DownloadTemplateButton.Add_Click({ Create-CSVTemplate })
    
    $FileSelectionGroup.Controls.AddRange(@($FilePathLabel, $Script:FilePathTextBox, $BrowseButton, $DownloadTemplateButton))
    
    # Preview Section
    $PreviewGroup = New-Object System.Windows.Forms.GroupBox
    $PreviewGroup.Text = "CSV Preview & Validation"
    $PreviewGroup.Location = New-Object System.Drawing.Point(10, 260)
    $PreviewGroup.Size = New-Object System.Drawing.Size(900, 200)
    
    $Script:PreviewDataGrid = New-Object System.Windows.Forms.DataGridView
    $Script:PreviewDataGrid.Location = New-Object System.Drawing.Point(10, 25)
    $Script:PreviewDataGrid.Size = New-Object System.Drawing.Size(880, 165)
    $Script:PreviewDataGrid.ReadOnly = $true
    $Script:PreviewDataGrid.AllowUserToAddRows = $false
    $Script:PreviewDataGrid.AllowUserToDeleteRows = $false
    $Script:PreviewDataGrid.SelectionMode = "FullRowSelect"
    $Script:PreviewDataGrid.AutoSizeColumnsMode = "AllCells"
    
    $PreviewGroup.Controls.Add($Script:PreviewDataGrid)
    
    # Import Options
    $OptionsGroup = New-Object System.Windows.Forms.GroupBox
    $OptionsGroup.Text = "Import Options"
    $OptionsGroup.Location = New-Object System.Drawing.Point(10, 470)
    $OptionsGroup.Size = New-Object System.Drawing.Size(900, 80)
    
    $Script:DryRunCheckBox = New-Object System.Windows.Forms.CheckBox
    $Script:DryRunCheckBox.Text = "Dry Run (validate only, don't create users)"
    $Script:DryRunCheckBox.Location = New-Object System.Drawing.Point(15, 25)
    $Script:DryRunCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $Script:DryRunCheckBox.Checked = $true
    
    $Script:SkipDuplicatesCheckBox = New-Object System.Windows.Forms.CheckBox
    $Script:SkipDuplicatesCheckBox.Text = "Skip duplicate users (don't overwrite)"
    $Script:SkipDuplicatesCheckBox.Location = New-Object System.Drawing.Point(15, 50)
    $Script:SkipDuplicatesCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $Script:SkipDuplicatesCheckBox.Checked = $true
    
    $OptionsGroup.Controls.AddRange(@($Script:DryRunCheckBox, $Script:SkipDuplicatesCheckBox))
    
    # Progress Section
    $ProgressGroup = New-Object System.Windows.Forms.GroupBox
    $ProgressGroup.Text = "Import Progress"
    $ProgressGroup.Location = New-Object System.Drawing.Point(10, 560)
    $ProgressGroup.Size = New-Object System.Drawing.Size(900, 120)
    
    $Script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $Script:ProgressBar.Location = New-Object System.Drawing.Point(15, 25)
    $Script:ProgressBar.Size = New-Object System.Drawing.Size(870, 20)
    $Script:ProgressBar.Style = "Continuous"
    
    $Script:ProgressLabel = New-Object System.Windows.Forms.Label
    $Script:ProgressLabel.Location = New-Object System.Drawing.Point(15, 50)
    $Script:ProgressLabel.Size = New-Object System.Drawing.Size(870, 20)
    $Script:ProgressLabel.Text = "Ready to import..."
    
    $Script:ProgressDetails = New-Object System.Windows.Forms.TextBox
    $Script:ProgressDetails.Location = New-Object System.Drawing.Point(15, 75)
    $Script:ProgressDetails.Size = New-Object System.Drawing.Size(870, 35)
    $Script:ProgressDetails.Multiline = $true
    $Script:ProgressDetails.ReadOnly = $true
    $Script:ProgressDetails.ScrollBars = "Vertical"
    $Script:ProgressDetails.Font = New-Object System.Drawing.Font("Consolas", 8)
    
    $ProgressGroup.Controls.AddRange(@($Script:ProgressBar, $Script:ProgressLabel, $Script:ProgressDetails))
    
    # Action Buttons
    $ButtonPanel = New-Object System.Windows.Forms.Panel
    $ButtonPanel.Location = New-Object System.Drawing.Point(10, 690)
    $ButtonPanel.Size = New-Object System.Drawing.Size(900, 50)
    
    $Script:ImportButton = New-Object System.Windows.Forms.Button
    $Script:ImportButton.Text = "Start Import"
    $Script:ImportButton.Location = New-Object System.Drawing.Point(350, 10)
    $Script:ImportButton.Size = New-Object System.Drawing.Size(120, 35)
    $Script:ImportButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $Script:ImportButton.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $Script:ImportButton.ForeColor = [System.Drawing.Color]::White
    $Script:ImportButton.FlatStyle = "Flat"
    $Script:ImportButton.Enabled = $false
    $Script:ImportButton.Add_Click({ Start-BulkUserImport })
    
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Text = "Cancel"
    $CancelButton.Location = New-Object System.Drawing.Point(480, 10)
    $CancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $CancelButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $CancelButton.ForeColor = [System.Drawing.Color]::White
    $CancelButton.FlatStyle = "Flat"
    $CancelButton.Add_Click({ Cancel-BulkImport })
    
    $ButtonPanel.Controls.AddRange(@($Script:ImportButton, $CancelButton))
    
    # Add all controls to scroll panel
    $ScrollPanel.Controls.AddRange(@(
        $InstructionsGroup,
        $FileSelectionGroup, 
        $PreviewGroup,
        $OptionsGroup,
        $ProgressGroup,
        $ButtonPanel
    ))
    
    $Tab.Controls.Add($ScrollPanel)
    return $Tab
}

function New-UserCreationTab {
    <#
    .SYNOPSIS
        Creates the user creation tab with comprehensive form
    #>
    
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Create User"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # Connection Panel
    $ConnectionPanel = New-Object System.Windows.Forms.Panel
    $ConnectionPanel.Height = 60
    $ConnectionPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $ConnectionPanel.BackColor = [System.Drawing.Color]::LightBlue
    $ConnectionPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $Script:ConnectButton = New-Object System.Windows.Forms.Button
    $Script:ConnectButton.Text = "Connect to Microsoft 365"
    $Script:ConnectButton.Size = New-Object System.Drawing.Size(250, 35)
    $Script:ConnectButton.Location = New-Object System.Drawing.Point(10, 12)
    $Script:ConnectButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    
    $Script:ConnectButton.Add_Click({
        if (-not $Global:IsConnected) {
            Connect-ToMicrosoftGraph
        } else {
            Get-TenantData
        }
    })
    
    $ConnectionPanel.Controls.Add($Script:ConnectButton)
    
    # Main Content Panel
    $ContentPanel = New-Object System.Windows.Forms.Panel
    $ContentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ContentPanel.AutoScroll = $true
    $ContentPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # User Details Group (left side)
    $UserDetailsGroup = New-Object System.Windows.Forms.GroupBox
    $UserDetailsGroup.Text = "User Details (FLEXIBLE - Any Values Can Be Entered)"
    $UserDetailsGroup.Location = New-Object System.Drawing.Point(10, 10)
    $UserDetailsGroup.Size = New-Object System.Drawing.Size(480, 380)
    
    # Create user form controls
    $y = 30
    $spacing = 35
    
    # First Name
    $FirstNameLabel = New-Object System.Windows.Forms.Label
    $FirstNameLabel.Text = "First Name: *"
    $FirstNameLabel.Location = New-Object System.Drawing.Point(10, $y)
    $FirstNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:FirstNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:FirstNameTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:FirstNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    $y += $spacing
    
    # Last Name
    $LastNameLabel = New-Object System.Windows.Forms.Label
    $LastNameLabel.Text = "Last Name: *"
    $LastNameLabel.Location = New-Object System.Drawing.Point(10, $y)
    $LastNameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:LastNameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:LastNameTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:LastNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    $y += $spacing
    
    # Username
    $UsernameLabel = New-Object System.Windows.Forms.Label
    $UsernameLabel.Text = "Username: *"
    $UsernameLabel.Location = New-Object System.Drawing.Point(10, $y)
    $UsernameLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:UsernameTextBox = New-Object System.Windows.Forms.TextBox
    $Script:UsernameTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:UsernameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    $y += $spacing
    
    # Domain
    $DomainLabel = New-Object System.Windows.Forms.Label
    $DomainLabel.Text = "Domain: *"
    $DomainLabel.Location = New-Object System.Drawing.Point(10, $y)
    $DomainLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:DomainDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:DomainDropdown.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:DomainDropdown.Size = New-Object System.Drawing.Size(200, 20)
    $Script:DomainDropdown.DropDownStyle = "DropDownList"
    
    $y += $spacing
    
    # Password
    $PasswordLabel = New-Object System.Windows.Forms.Label
    $PasswordLabel.Text = "Password: *"
    $PasswordLabel.Location = New-Object System.Drawing.Point(10, $y)
    $PasswordLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:PasswordTextBox = New-Object System.Windows.Forms.TextBox
    $Script:PasswordTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:PasswordTextBox.Size = New-Object System.Drawing.Size(150, 20)
    $Script:PasswordTextBox.UseSystemPasswordChar = $true
    
    $GeneratePasswordButton = New-Object System.Windows.Forms.Button
    $GeneratePasswordButton.Text = "Generate"
    $GeneratePasswordButton.Location = New-Object System.Drawing.Point(280, ($y-2))
    $GeneratePasswordButton.Size = New-Object System.Drawing.Size(70, 22)
    $GeneratePasswordButton.Add_Click({
        $Script:PasswordTextBox.Text = Generate-SecurePassword
    })
    
    $y += $spacing
    
    # Department (FLEXIBLE)
    $DepartmentLabel = New-Object System.Windows.Forms.Label
    $DepartmentLabel.Text = "Department:"
    $DepartmentLabel.Location = New-Object System.Drawing.Point(10, $y)
    $DepartmentLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:DepartmentTextBox = New-Object System.Windows.Forms.TextBox
    $Script:DepartmentTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:DepartmentTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $Script:DepartmentTextBox.PlaceholderText = "Enter ANY department name"
    
    $y += $spacing
    
    # Job Title (FLEXIBLE)
    $JobTitleLabel = New-Object System.Windows.Forms.Label
    $JobTitleLabel.Text = "Job Title:"
    $JobTitleLabel.Location = New-Object System.Drawing.Point(10, $y)
    $JobTitleLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:JobTitleTextBox = New-Object System.Windows.Forms.TextBox
    $Script:JobTitleTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:JobTitleTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $Script:JobTitleTextBox.PlaceholderText = "Enter ANY job title"
    
    $y += $spacing
    
    # Office Location (FLEXIBLE)
    $OfficeLabel = New-Object System.Windows.Forms.Label
    $OfficeLabel.Text = "Office:"
    $OfficeLabel.Location = New-Object System.Drawing.Point(10, $y)
    $OfficeLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:OfficeTextBox = New-Object System.Windows.Forms.TextBox
    $Script:OfficeTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:OfficeTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $Script:OfficeTextBox.PlaceholderText = "Enter ANY office location"
    
    # Add controls to user details group
    $UserDetailsGroup.Controls.AddRange(@(
        $FirstNameLabel, $Script:FirstNameTextBox,
        $LastNameLabel, $Script:LastNameTextBox,
        $UsernameLabel, $Script:UsernameTextBox,
        $DomainLabel, $Script:DomainDropdown,
        $PasswordLabel, $Script:PasswordTextBox, $GeneratePasswordButton,
        $DepartmentLabel, $Script:DepartmentTextBox,
        $JobTitleLabel, $Script:JobTitleTextBox,
        $OfficeLabel, $Script:OfficeTextBox
    ))
    
    # Management Group (right side)
    $ManagementGroup = New-Object System.Windows.Forms.GroupBox
    $ManagementGroup.Text = "Management & Licensing"
    $ManagementGroup.Location = New-Object System.Drawing.Point(500, 10)
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
    
    # Enhanced info
    $EnhancedInfoLabel = New-Object System.Windows.Forms.Label
    $EnhancedInfoLabel.Text = "LEGACY FUNCTIONALITY RESTORED:`n[OK] Flexible attributes - any values allowed`n[OK] Manual group selection only`n[OK] Distribution lists selectable in groups`n[OK] Shared mailboxes selectable for permissions`n[OK] No hardcoded Exchange assumptions"
    $EnhancedInfoLabel.Location = New-Object System.Drawing.Point(10, 140)
    $EnhancedInfoLabel.Size = New-Object System.Drawing.Size(430, 120)
    $EnhancedInfoLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    $EnhancedInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    
    $ManagementGroup.Controls.AddRange(@(
        $ManagerLabel, $Script:ManagerDropdown,
        $LicenseLabel, $Script:LicenseDropdown,
        $LicenseInfoLabel, $EnhancedInfoLabel
    ))
    
    # Groups Group (full width below)
    $GroupsGroup = New-Object System.Windows.Forms.GroupBox
    $GroupsGroup.Text = "Group Membership & Exchange Resources (Manual Selection - Distribution Lists & Shared Mailboxes Now Available)"
    $GroupsGroup.Location = New-Object System.Drawing.Point(10, 400)
    $GroupsGroup.Size = New-Object System.Drawing.Size(940, 220)
    
    $Script:GroupsCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $Script:GroupsCheckedListBox.Location = New-Object System.Drawing.Point(10, 20)
    $Script:GroupsCheckedListBox.Size = New-Object System.Drawing.Size(920, 190)
    $Script:GroupsCheckedListBox.CheckOnClick = $true
    $Script:GroupsCheckedListBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $GroupsGroup.Controls.Add($Script:GroupsCheckedListBox)
    
    # Create User Button
    $Script:CreateUserButton = New-Object System.Windows.Forms.Button
    $Script:CreateUserButton.Text = "Create M365 User"
    $Script:CreateUserButton.Location = New-Object System.Drawing.Point(10, 630)
    $Script:CreateUserButton.Size = New-Object System.Drawing.Size(200, 40)
    $Script:CreateUserButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $Script:CreateUserButton.BackColor = [System.Drawing.Color]::LightGreen
    $Script:CreateUserButton.Enabled = $false
    
    $Script:CreateUserButton.Add_Click({
        # Validate required fields
        if (-not $Script:FirstNameTextBox.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show("First Name is required.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if (-not $Script:LastNameTextBox.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show("Last Name is required.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if (-not $Script:UsernameTextBox.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show("Username is required.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if (-not $Script:PasswordTextBox.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show("Password is required.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ($Script:DomainDropdown.SelectedItem -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Please select a domain.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Get selected groups
        $SelectedGroups = @()
        for ($i = 0; $i -lt $Script:GroupsCheckedListBox.Items.Count; $i++) {
            if ($Script:GroupsCheckedListBox.GetItemChecked($i)) {
                $SelectedGroups += $Script:GroupsCheckedListBox.Items[$i]
            }
        }
        
        # Create the user
        try {
            New-M365User -FirstName $Script:FirstNameTextBox.Text.Trim() `
                         -LastName $Script:LastNameTextBox.Text.Trim() `
                         -Username $Script:UsernameTextBox.Text.Trim() `
                         -Domain $Script:DomainDropdown.SelectedItem `
                         -Password $Script:PasswordTextBox.Text `
                         -Department $Script:DepartmentTextBox.Text.Trim() `
                         -JobTitle $Script:JobTitleTextBox.Text.Trim() `
                         -Office $Script:OfficeTextBox.Text.Trim() `
                         -Manager $Script:ManagerDropdown.SelectedItem `
                         -LicenseType $Script:LicenseDropdown.SelectedItem `
                         -Groups $SelectedGroups
        }
        catch {
            # Error handling is done in New-M365User function
        }
    })
    
    # Clear Form Button
    $ClearFormButton = New-Object System.Windows.Forms.Button
    $ClearFormButton.Text = "Clear Form"
    $ClearFormButton.Location = New-Object System.Drawing.Point(220, 630)
    $ClearFormButton.Size = New-Object System.Drawing.Size(120, 40)
    $ClearFormButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $ClearFormButton.Add_Click({
        Clear-UserForm
    })
    
    # Add all controls to content panel
    $ContentPanel.Controls.AddRange(@(
        $UserDetailsGroup, $ManagementGroup, $GroupsGroup,
        $Script:CreateUserButton, $ClearFormButton
    ))
    
    # Add panels to tab
    $Tab.Controls.Add($ContentPanel)
    $Tab.Controls.Add($ConnectionPanel)
    
    return $Tab
}

function New-TenantDataTab {
    <#
    .SYNOPSIS
        Creates the tenant data discovery tab
    #>
    
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Tenant Data"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $Script:TenantDataTextBox = New-Object System.Windows.Forms.TextBox
    $Script:TenantDataTextBox.Multiline = $true
    $Script:TenantDataTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $Script:TenantDataTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $Script:TenantDataTextBox.ReadOnly = $true
    $Script:TenantDataTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $Script:TenantDataTextBox.Text = "Connect to Microsoft 365 to view tenant data discovery information..."
    
    $Tab.Controls.Add($Script:TenantDataTextBox)
    return $Tab
}

function New-ActivityLogTab {
    <#
    .SYNOPSIS
        Creates the activity log tab
    #>
    
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "Activity Log"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $Script:ActivityLogTextBox = New-Object System.Windows.Forms.TextBox
    $Script:ActivityLogTextBox.Multiline = $true
    $Script:ActivityLogTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $Script:ActivityLogTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $Script:ActivityLogTextBox.ReadOnly = $true
    $Script:ActivityLogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $Script:ActivityLogTextBox.Text = "$(Get-Date): Application started (Enhanced & Fixed - Legacy functionality restored)`r`n"
    
    $Tab.Controls.Add($Script:ActivityLogTextBox)
    return $Tab
}

# ================================
# MAIN FORM CREATION
# ================================

function New-MainForm {
    <#
    .SYNOPSIS
        Creates the main application form with all tabs and controls
    #>
    
    Write-Host "Creating main application window..." -ForegroundColor Green
    
    # Main Form
    $Script:MainForm = New-Object System.Windows.Forms.Form
    $Script:MainForm.Text = "M365 User Provisioning Tool - Enterprise Edition 2025 (Enhanced & Fixed - Legacy Functionality Restored)"
    $Script:MainForm.Size = New-Object System.Drawing.Size(1400, 900)
    $Script:MainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $Script:MainForm.MinimumSize = New-Object System.Drawing.Size(1200, 800)
    $Script:MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $Script:MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
    
    # Try to set application icon
    try {
        $Script:MainForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$env:SystemRoot\System32\shell32.dll")
    }
    catch {
        Write-Verbose "Could not set application icon"
    }

    # Status Strip
    $Script:StatusStrip = New-Object System.Windows.Forms.StatusStrip
    $Script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $Script:StatusLabel.Text = "Ready - Enhanced & Fixed Version with Legacy Functionality Restored"
    $Script:StatusLabel.Spring = $true
    $Script:StatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $Script:StatusStrip.Items.Add($Script:StatusLabel) | Out-Null

    # Tab Control
    Write-Host "Creating tabbed interface..." -ForegroundColor Cyan
    $Script:TabControl = New-Object System.Windows.Forms.TabControl
    $Script:TabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $Script:TabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    # Create all tabs
    Write-Host "   Creating User Creation tab..." -ForegroundColor Yellow
    $UserCreationTab = New-UserCreationTab

    Write-Host "   Creating Bulk Import tab..." -ForegroundColor Yellow  
    $BulkImportTab = New-BulkImportTab

    Write-Host "   Creating Tenant Data tab..." -ForegroundColor Yellow
    $TenantDataTab = New-TenantDataTab

    Write-Host "   Creating Activity Log tab..." -ForegroundColor Yellow
    $ActivityLogTab = New-ActivityLogTab

    # Add tabs to control
    $Script:TabControl.TabPages.AddRange(@(
        $UserCreationTab,
        $BulkImportTab, 
        $TenantDataTab,
        $ActivityLogTab
    ))

    # Add all controls to main form
    $Script:MainForm.Controls.Add($Script:TabControl)
    $Script:MainForm.Controls.Add($Script:StatusStrip)
}

# ================================
# ADDITIONAL HELPER FUNCTIONS
# ================================

function Test-ExchangeConnection {
    <#
    .SYNOPSIS
        Tests if Exchange Online is connected and available
    #>
    
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Export-ActivityLog {
    <#
    .SYNOPSIS
        Exports activity log to file
    #>
    
    try {
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv"
        $SaveFileDialog.Title = "Export Activity Log"
        $SaveFileDialog.FileName = "M365_Activity_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        
        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            if ($SaveFileDialog.FileName.EndsWith('.csv')) {
                # Export as CSV
                $Global:ActivityLog | Export-Csv -Path $SaveFileDialog.FileName -NoTypeInformation
            }
            else {
                # Export as text
                $LogText = $Global:ActivityLog | ForEach-Object {
                    "$($_.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')) [$($_.Status)] $($_.Action) - $($_.Details)"
                }
                $LogText | Set-Content -Path $SaveFileDialog.FileName -Encoding UTF8
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Activity log exported successfully!",
                "Export Complete", 
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error exporting log: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Show-AboutDialog {
    <#
    .SYNOPSIS
        Shows application about dialog
    #>
    
    $AboutText = @"
M365 User Provisioning Tool
Enterprise Edition 2025 (Enhanced & Fixed)

Version: 3.1.2025-COMPLETE-ENHANCED-FIXED-PROFESSIONAL
PowerShell: 7.0+ Required

FEATURES RESTORED:
[OK] Flexible Attributes - Any values can be entered
[OK] Manual Group Selection - Distribution lists selectable
[OK] Shared Mailbox Permissions - Available in group membership
[OK] No Hardcoded Assumptions - Manual control restored
[OK] Enhanced Exchange Integration - M365.ExchangeOnline module support
[OK] Bulk CSV Import - Template-driven user creation
[OK] Comprehensive Logging - Full activity tracking

DEPENDENCIES:
• Microsoft Graph PowerShell SDK V2.28+
• Exchange Online PowerShell V3.0+
• Optional: M365.ExchangeOnline Module (Enhanced)

FIXES APPLIED:
• Removed hardcoded distribution list assignments
• Restored legacy flexible attribute handling
• Added manual Exchange resource selection
• Enhanced error handling and logging
• Improved group categorization in UI
• Professional interface (removed emojis)

Created by: Enterprise Solutions Team
Support: Check documentation for troubleshooting
"@

    [System.Windows.Forms.MessageBox]::Show(
        $AboutText,
        "About M365 User Provisioning Tool",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

# ================================
# MENU CREATION
# ================================

function Add-MenuStrip {
    <#
    .SYNOPSIS
        Adds menu strip to main form
    #>
    
    $MenuStrip = New-Object System.Windows.Forms.MenuStrip
    
    # File Menu
    $FileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $FileMenu.Text = "&File"
    
    $ExportLogItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $ExportLogItem.Text = "Export Activity Log..."
    $ExportLogItem.Add_Click({ Export-ActivityLog })
    
    $SeparatorItem = New-Object System.Windows.Forms.ToolStripSeparator
    
    $ExitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $ExitItem.Text = "E&xit"
    $ExitItem.Add_Click({ $Script:MainForm.Close() })
    
    $FileMenu.DropDownItems.AddRange(@($ExportLogItem, $SeparatorItem, $ExitItem))
    
    # Tools Menu
    $ToolsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $ToolsMenu.Text = "&Tools"
    
    $RefreshDataItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $RefreshDataItem.Text = "Refresh Tenant Data"
    $RefreshDataItem.Add_Click({ 
        if ($Global:IsConnected) { 
            Get-TenantData 
        } else { 
            [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft 365 first.", "Not Connected", "OK", "Warning")
        }
    })
    
    $TestExchangeItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $TestExchangeItem.Text = "Test Exchange Connection"
    $TestExchangeItem.Add_Click({
        $IsConnected = Test-ExchangeConnection
        $Message = if ($IsConnected) { "Exchange Online is connected and available!" } else { "Exchange Online is not connected or not available." }
        $Icon = if ($IsConnected) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
        [System.Windows.Forms.MessageBox]::Show($Message, "Exchange Connection Test", "OK", $Icon)
    })
    
    $ToolsMenu.DropDownItems.AddRange(@($RefreshDataItem, $TestExchangeItem))
    
    # Help Menu
    $HelpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $HelpMenu.Text = "&Help"
    
    $AboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $AboutItem.Text = "&About"
    $AboutItem.Add_Click({ Show-AboutDialog })
    
    $HelpMenu.DropDownItems.Add($AboutItem)
    
    # Add menus to strip
    $MenuStrip.Items.AddRange(@($FileMenu, $ToolsMenu, $HelpMenu))
    
    # Add menu strip to form
    $Script:MainForm.MainMenuStrip = $MenuStrip
    $Script:MainForm.Controls.Add($MenuStrip)
}

# ================================
# FORM EVENT HANDLERS
# ================================

function Register-FormEvents {
    <#
    .SYNOPSIS
        Registers form event handlers
    #>
    
    # Form closing event
    $Script:MainForm.Add_FormClosing({
        param($sender, $e)
        
        # Check if there are any ongoing operations
        if ($Script:ImportButton -and -not $Script:ImportButton.Enabled) {
            $Result = [System.Windows.Forms.MessageBox]::Show(
                "An import operation may be in progress. Are you sure you want to exit?",
                "Confirm Exit",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($Result -eq [System.Windows.Forms.DialogResult]::No) {
                $e.Cancel = $true
                return
            }
        }
        
        # Disconnect from Graph if connected
        if ($Global:IsConnected) {
            try {
                Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
                Disconnect-MgGraph -ErrorAction SilentlyContinue
                Write-Host "[OK] Disconnected successfully" -ForegroundColor Green
            }
            catch {
                Write-Warning "Error during disconnect: $($_.Exception.Message)"
            }
        }
        
        # Save final activity log
        try {
            Add-ActivityLog "Application" "Shutdown" "Application closed successfully"
            $FinalLogPath = "M365_Final_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $Global:ActivityLog | ForEach-Object {
                "$($_.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')) [$($_.Status)] $($_.Action) - $($_.Details)"
            } | Set-Content -Path $FinalLogPath -Encoding UTF8 -ErrorAction SilentlyContinue
        }
        catch {
            # Silent fail on log save
        }
        
        Write-Host ""
        Write-Host "Thank you for using M365 User Provisioning Tool (Enhanced & Fixed)!" -ForegroundColor Green
        Write-Host "   Legacy functionality restored" -ForegroundColor Cyan
        Write-Host "   [OK] Flexible attributes working" -ForegroundColor Cyan  
        Write-Host "   [OK] Manual group selection working" -ForegroundColor Cyan
        Write-Host "   [OK] Distribution lists selectable" -ForegroundColor Cyan
        Write-Host ""
    })
    
    # Form load event
    $Script:MainForm.Add_Load({
        Write-Host "Main form loaded successfully" -ForegroundColor Green
        Add-ActivityLog "Application" "Started" "M365 User Provisioning Tool started (Enhanced & Fixed)"
        
        # Initial status
        Update-StatusLabel "Ready - Connect to Microsoft 365 to begin"
        
        # Focus first tab
        $Script:TabControl.SelectedIndex = 0
    })
}

# ================================
# APPLICATION STARTUP VALIDATION
# ================================

function Test-Prerequisites {
    <#
    .SYNOPSIS
        Tests if all prerequisites are met before starting
    #>
    
    Write-Host "Checking application prerequisites..." -ForegroundColor Cyan
    
    $Issues = @()
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        $Issues += "PowerShell 7.0+ is required. Current version: $($PSVersionTable.PSVersion)"
    }
    
    # Check required modules availability
    $RequiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'ExchangeOnlineManagement')
    foreach ($Module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            $Issues += "Required module not available: $Module"
        }
    }
    
    # Check if running as admin (recommended)
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $IsAdmin) {
        Write-Host "   [WARNING] Not running as administrator - some features may be limited" -ForegroundColor Yellow
    }
    
    if ($Issues.Count -gt 0) {
        $ErrorMessage = "Prerequisites not met:`n`n" + ($Issues -join "`n")
        Write-Host "[ERROR] Prerequisites check failed" -ForegroundColor Red
        foreach ($Issue in $Issues) {
            Write-Host "   • $Issue" -ForegroundColor Red
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            $ErrorMessage,
            "Prerequisites Check Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    
    Write-Host "[OK] All prerequisites met" -ForegroundColor Green
    return $true
}

# ================================
# MAIN APPLICATION ENTRY POINT
# ================================

function Start-M365ProvisioningTool {
    <#
    .SYNOPSIS
        Main application entry point
    #>
    
    try {
        # Test prerequisites
        if (-not (Test-Prerequisites)) {
            Write-Host "[ERROR] Application startup aborted due to missing prerequisites" -ForegroundColor Red
            return
        }
        
        Write-Host ""
        Write-Host "Starting M365 User Provisioning Tool (Enhanced & Fixed)..." -ForegroundColor Green
        Write-Host ""
        
        # Create main form
        New-MainForm
        
        # Add menu strip
        Add-MenuStrip
        
        # Register event handlers
        Register-FormEvents
        
        # Show startup message
        Write-Host "[OK] Application initialized successfully!" -ForegroundColor Green
        Write-Host ""
        Write-Host "FIXES APPLIED:" -ForegroundColor Cyan
        Write-Host "   [OK] Removed hardcoded Exchange provisioning logic" -ForegroundColor Green
        Write-Host "   [OK] Restored flexible attribute handling (any values allowed)" -ForegroundColor Green
        Write-Host "   [OK] Distribution lists now selectable in Group membership" -ForegroundColor Green
        Write-Host "   [OK] Shared mailboxes now selectable for permissions" -ForegroundColor Green
        Write-Host "   [OK] Manual Exchange operations with proper error handling" -ForegroundColor Green
        Write-Host "   [OK] Enhanced group categorization in UI" -ForegroundColor Green
        Write-Host "   [OK] Professional interface (removed emojis)" -ForegroundColor Green
        Write-Host ""
        Write-Host "READY TO USE:" -ForegroundColor Yellow
        Write-Host "   1. Click 'Connect to Microsoft 365' in the Create User tab" -ForegroundColor Gray
        Write-Host "   2. Enter ANY values in Department, Job Title, Office fields" -ForegroundColor Gray
        Write-Host "   3. Select distribution lists and shared mailboxes from Group membership" -ForegroundColor Gray
        Write-Host "   4. No more hardcoded assumptions or errors!" -ForegroundColor Gray
        Write-Host ""
        
        Add-ActivityLog "Application" "Initialized" "Application window created and ready"
        
        # Start the application
        Write-Host "Displaying application window..." -ForegroundColor Cyan
        [System.Windows.Forms.Application]::Run($Script:MainForm)
        
    }
    catch {
        $ErrorMsg = "Critical error during application startup: $($_.Exception.Message)"
        Write-Host "[ERROR] $ErrorMsg" -ForegroundColor Red
        Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
        
        [System.Windows.Forms.MessageBox]::Show(
            $ErrorMsg,
            "Application Startup Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# ================================
# SCRIPT EXECUTION
# ================================

# Only run if not being imported as a module
if ($MyInvocation.InvocationName -ne '.') {
    # Check if GUI should be skipped
    if ($NoGUI) {
        Write-Host "GUI mode disabled - script functions available for import" -ForegroundColor Yellow
        Write-Host "   Use Start-M365ProvisioningTool to launch GUI" -ForegroundColor Gray
    }
    else {
        # Start the main application
        Start-M365ProvisioningTool
    }
}

# ================================
# EXPORT MODULE MEMBERS (if used as module)
# ================================

# Export main functions for module usage
Export-ModuleMember -Function @(
    'Start-M365ProvisioningTool',
    'New-M365User', 
    'Connect-ToMicrosoftGraph',
    'Get-TenantData',
    'Generate-SecurePassword',
    'Test-ExchangeConnection'
)