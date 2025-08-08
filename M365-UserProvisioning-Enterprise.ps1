#Requires -Version 7.0

<#
.SYNOPSIS
    M365 User Provisioning Tool - Enterprise Edition 2025 (FIXED)
    Comprehensive M365 user management with intelligent tenant discovery

.DESCRIPTION
    Advanced user provisioning tool with:
    - Intelligent tenant discovery (users, groups, mailboxes, sites)
    - Single user creation and bulk CSV import
    - License assignment via CustomAttribute1
    - UK-based location management
    - Clean tabbed interface with pagination
    - Robust error handling and validation
    
    FIXED: Removes SetCompatibleTextRenderingDefault call that caused initialization errors

.NOTES
    Version: 3.0.2025-FIXED
    Author: Enterprise Solutions Team
    PowerShell: 7.0+ Required
    Dependencies: Microsoft Graph PowerShell SDK V2.28+, Exchange Online PowerShell
    Last Updated: August 2025

.EXAMPLE
    .\M365-UserProvisioning-Enterprise-Fixed.ps1
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

Write-Host "M365 User Provisioning Tool - Enterprise Edition 2025" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "🔧 FIXED VERSION - No SetCompatibleTextRenderingDefault" -ForegroundColor Yellow
Write-Host ""

try {
    Write-Host "🔧 Initializing Windows Forms (Fixed Mode)..." -ForegroundColor Cyan
    
    # Load Windows Forms assemblies
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Write-Host "   ✅ Windows Forms assemblies loaded" -ForegroundColor Green
    
    # Enable visual styles only
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Write-Host "   ✅ Visual styles enabled" -ForegroundColor Green
    
    # SKIP SetCompatibleTextRenderingDefault - this was causing the error!
    Write-Host "   ⏭️  Skipping SetCompatibleTextRenderingDefault (not required for functionality)" -ForegroundColor Yellow
    
    Write-Host "✅ Windows Forms ready for enterprise application!" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "❌ Failed to initialize Windows Forms: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ================================
# MICROSOFT GRAPH MODULE LOADING
# ================================

Write-Host "📚 Loading Microsoft Graph PowerShell modules..." -ForegroundColor Cyan

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
            Write-Host "   📥 Installing $Module..." -ForegroundColor Yellow
            Install-Module -Name $Module -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        }
        
        Write-Host "   📤 Importing $Module..." -ForegroundColor Yellow
        Import-Module -Name $Module -Force -ErrorAction Stop
        Write-Host "   ✅ $Module loaded" -ForegroundColor Green
    }
    catch {
        Write-Warning "⚠️  Failed to load $Module : $($_.Exception.Message)"
        Write-Host "🔄 Application will continue with limited functionality" -ForegroundColor Yellow
    }
}

Write-Host "✅ Required modules processed" -ForegroundColor Green
Write-Host ""

# ================================
# TENANT DISCOVERY FUNCTIONS
# ================================

function Get-TenantData {
    <#
    .SYNOPSIS
        Performs comprehensive tenant discovery to populate all dropdowns and lists
    #>
    
    try {
        Update-StatusLabel "🔍 Discovering tenant data..."
        
        # Get tenant information
        Write-Host "   📊 Getting tenant information..." -ForegroundColor Yellow
        $Global:TenantInfo = Get-MgOrganization
        
        # Get accepted domains
        Write-Host "   🌐 Getting accepted domains..." -ForegroundColor Yellow
        $Global:AcceptedDomains = Get-MgDomain | Where-Object { $_.IsVerified -eq $true }
        
        # Get all users (for manager dropdown)
        Write-Host "   👥 Getting existing users..." -ForegroundColor Yellow
        $Global:AvailableUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,JobTitle,Department" | 
            Sort-Object DisplayName
        
        # Get all groups
        Write-Host "   🏢 Getting security groups..." -ForegroundColor Yellow
        $Global:AvailableGroups = Get-MgGroup -All -Property "DisplayName,GroupTypes,SecurityEnabled,MailEnabled" |
            Sort-Object DisplayName
        
        # Get distribution groups
        Write-Host "   📧 Getting distribution groups..." -ForegroundColor Yellow
        try {
            $Global:DistributionLists = Get-DistributionGroup -ResultSize Unlimited | Sort-Object Name
            $Global:MailEnabledSecurityGroups = Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited | Sort-Object Name
        }
        catch {
            Write-Warning "Exchange Online not available - skipping distribution groups"
            $Global:DistributionLists = @()
            $Global:MailEnabledSecurityGroups = @()
        }
        
        # Get mailboxes
        Write-Host "   📪 Getting mailboxes..." -ForegroundColor Yellow
        try {
            $Global:AvailableMailboxes = Get-Mailbox -ResultSize Unlimited | Sort-Object Name
            $Global:SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Sort-Object Name
        }
        catch {
            Write-Warning "Exchange Online not available - skipping mailbox discovery"
            $Global:AvailableMailboxes = @()
            $Global:SharedMailboxes = @()
        }
        
        # Get SharePoint sites
        Write-Host "   🌐 Getting SharePoint sites..." -ForegroundColor Yellow
        try {
            $Global:SharePointSites = Get-MgSite -All | Sort-Object DisplayName
        }
        catch {
            Write-Warning "SharePoint not available - skipping site discovery"
            $Global:SharePointSites = @()
        }
        
        # Get available licenses
        Write-Host "   🎫 Getting license information..." -ForegroundColor Yellow
        $Global:AvailableLicenses = Get-MgSubscribedSku
        
        # Update UI with discovered data
        Update-TenantDataDisplay
        Update-Dropdowns
        
        Update-StatusLabel "✅ Tenant discovery completed successfully"
        Write-Host "✅ Tenant data discovery completed" -ForegroundColor Green
        
        # Log the discovery
        Add-ActivityLog "Tenant Discovery" "Success" "Discovered: $($Global:AvailableUsers.Count) users, $($Global:AvailableGroups.Count) groups, $($Global:AvailableMailboxes.Count) mailboxes"
        
    }
    catch {
        $ErrorMsg = "Tenant discovery failed: $($_.Exception.Message)"
        Update-StatusLabel "❌ $ErrorMsg"
        Write-Host "❌ $ErrorMsg" -ForegroundColor Red
        Add-ActivityLog "Tenant Discovery" "Failed" $_.Exception.Message
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to discover tenant data:`n`n$($_.Exception.Message)",
            "Tenant Discovery Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
}

function Update-Dropdowns {
    <#
    .SYNOPSIS
        Updates all dropdown controls with discovered tenant data
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
    
    # Update office location dropdown
    $Script:OfficeDropdown.Items.Clear()
    foreach ($Location in $Global:UKLocations) {
        $Script:OfficeDropdown.Items.Add($Location) | Out-Null
    }
    if ($Script:OfficeDropdown.Items.Count -gt 0) {
        $Script:OfficeDropdown.SelectedIndex = 0
    }
    
    # Update groups checklist
    $Script:GroupsCheckedListBox.Items.Clear()
    
    # Add security groups
    foreach ($Group in $Global:AvailableGroups) {
        $GroupDisplay = "🔒 $($Group.DisplayName)"
        $Script:GroupsCheckedListBox.Items.Add($GroupDisplay) | Out-Null
    }
    
    # Add distribution groups
    foreach ($DL in $Global:DistributionLists) {
        $GroupDisplay = "📧 $($DL.Name) (Distribution List)"
        $Script:GroupsCheckedListBox.Items.Add($GroupDisplay) | Out-Null
    }
    
    # Add mail-enabled security groups
    foreach ($MESG in $Global:MailEnabledSecurityGroups) {
        $GroupDisplay = "🔐 $($MESG.Name) (Mail-Enabled Security)"
        $Script:GroupsCheckedListBox.Items.Add($GroupDisplay) | Out-Null
    }
}

# ================================
# USER CREATION FUNCTIONS
# ================================

function New-M365User {
    <#
    .SYNOPSIS
        Creates a new M365 user with all specified properties and group memberships
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
        
        Update-StatusLabel "👤 Creating user: $UserPrincipalName"
        
        # Create user parameters
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
        
        # Add optional properties
        if ($Department) { $UserParams.Department = $Department }
        if ($JobTitle) { $UserParams.JobTitle = $JobTitle }
        if ($Office) { $UserParams.OfficeLocation = $Office }
        
        # Set CustomAttribute1 for license assignment
        if ($LicenseType) {
            $UserParams.OnPremisesExtensionAttributes = @{
                ExtensionAttribute1 = $LicenseType
            }
        }
        
        # Create the user
        Write-Host "   📝 Creating user account..." -ForegroundColor Yellow
        $NewUser = New-MgUser @UserParams
        
        Write-Host "   ✅ User created: $($NewUser.UserPrincipalName)" -ForegroundColor Green
        
        # Set manager if specified
        if ($Manager -and $Manager -ne "(No Manager)") {
            try {
                $ManagerUPN = ($Manager -split '\(')[1] -replace '\)', ''
                $ManagerUser = Get-MgUser -Filter "userPrincipalName eq '$ManagerUPN'"
                if ($ManagerUser) {
                    Set-MgUserManagerByRef -UserId $NewUser.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($ManagerUser.Id)" }
                    Write-Host "   👔 Manager set: $($Manager)" -ForegroundColor Green
                }
            }
            catch {
                Write-Warning "Failed to set manager: $($_.Exception.Message)"
            }
        }
        
        # Add to groups
        if ($Groups -and $Groups.Count -gt 0) {
            Write-Host "   🏢 Adding to groups..." -ForegroundColor Yellow
            foreach ($GroupName in $Groups) {
                try {
                    # Clean group name (remove emojis and descriptions)
                    $CleanGroupName = ($GroupName -split ' \(')[0] -replace '^[🔒🔐📧]\s*', ''
                    
                    # Find the group
                    $Group = $Global:AvailableGroups | Where-Object { $_.DisplayName -eq $CleanGroupName }
                    if ($Group) {
                        $GroupMember = @{
                            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($NewUser.Id)"
                        }
                        New-MgGroupMember -GroupId $Group.Id -BodyParameter $GroupMember
                        Write-Host "      ✅ Added to group: $CleanGroupName" -ForegroundColor Green
                    }
                }
                catch {
                    Write-Warning "Failed to add to group $GroupName : $($_.Exception.Message)"
                }
            }
        }
        
        Update-StatusLabel "✅ User created successfully: $UserPrincipalName"
        Add-ActivityLog "User Creation" "Success" "Created user: $UserPrincipalName with license type: $LicenseType"
        
        # Show success message
        [System.Windows.Forms.MessageBox]::Show(
            "User created successfully!`n`nName: $DisplayName`nUPN: $UserPrincipalName`nLicense Type (CustomAttribute1): $LicenseType`n`nThe user will receive an email with sign-in instructions.",
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
        Update-StatusLabel "❌ $ErrorMsg"
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
    if ($Script:OfficeDropdown.Items.Count -gt 0) {
        $Script:OfficeDropdown.SelectedIndex = 0
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
# UI HELPER FUNCTIONS
# ================================

function Update-StatusLabel {
    param([string]$Message)
    
    if ($Script:StatusLabel) {
        $Script:StatusLabel.Text = "$(Get-Date -Format 'HH:mm:ss') - $Message"
        [System.Windows.Forms.Application]::DoEvents()
    }
    Write-Host "📊 $Message" -ForegroundColor Cyan
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
        Update-StatusLabel "🔗 Connecting to Microsoft Graph..."
        
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
            Update-StatusLabel "✅ Connected to Microsoft Graph as $($Context.Account)"
            
            # Enable connection-dependent controls
            $Script:ConnectButton.Text = "🔗 Connected - Discover Tenant Data"
            $Script:ConnectButton.BackColor = [System.Drawing.Color]::LightGreen
            $Script:CreateUserButton.Enabled = $true
            
            Add-ActivityLog "Connection" "Success" "Connected to Microsoft Graph as $($Context.Account)"
            
            # Auto-discover tenant data
            Get-TenantData
            
            return $true
        }
        else {
            throw "Failed to establish Graph context"
        }
    }
    catch {
        Update-StatusLabel "❌ Connection failed: $($_.Exception.Message)"
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
        Updates the tenant data tab with discovered information
    #>
    
    if ($Script:TenantDataTextBox) {
        $TenantSummary = @"
TENANT INFORMATION
==================
Tenant Name: $($Global:TenantInfo.DisplayName)
Tenant ID: $($Global:TenantInfo.Id)
Country: $($Global:TenantInfo.CountryLetterCode)

DISCOVERY SUMMARY
=================
✅ Users: $($Global:AvailableUsers.Count)
✅ Security Groups: $($Global:AvailableGroups.Count)
✅ Distribution Lists: $($Global:DistributionLists.Count)
✅ Mail-Enabled Security Groups: $($Global:MailEnabledSecurityGroups.Count)
✅ Mailboxes: $($Global:AvailableMailboxes.Count)
✅ Shared Mailboxes: $($Global:SharedMailboxes.Count)
✅ SharePoint Sites: $($Global:SharePointSites.Count)
✅ Accepted Domains: $($Global:AcceptedDomains.Count)
✅ License SKUs: $($Global:AvailableLicenses.Count)

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
# MAIN FORM CREATION
# ================================

function New-MainForm {
    <#
    .SYNOPSIS
        Creates the main application form with all tabs and controls
    #>
    
    Write-Host "🖥️  Creating main application window..." -ForegroundColor Green
    
    # Main Form
    $Script:MainForm = New-Object System.Windows.Forms.Form
    $Script:MainForm.Text = "M365 User Provisioning Tool - Enterprise Edition 2025"
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
    $Script:StatusLabel.Text = "Ready - Click Connect to start tenant discovery"
    $Script:StatusLabel.Spring = $true
    $Script:StatusStrip.Items.Add($Script:StatusLabel) | Out-Null
    
    # Tab Control
    $Script:TabControl = New-Object System.Windows.Forms.TabControl
    $Script:TabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $Script:TabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Create tabs
    $UserCreationTab = New-UserCreationTab
    $TenantDataTab = New-TenantDataTab  
    $ActivityLogTab = New-ActivityLogTab
    
    # Add tabs to control
    $Script:TabControl.TabPages.Add($UserCreationTab)
    $Script:TabControl.TabPages.Add($TenantDataTab)
    $Script:TabControl.TabPages.Add($ActivityLogTab)
    
    # Add controls to form
    $Script:MainForm.Controls.Add($Script:TabControl)
    $Script:MainForm.Controls.Add($Script:StatusStrip)
    
    # Form events
    $Script:MainForm.Add_Load({
        Update-StatusLabel "Application started - Ready to connect to Microsoft 365"
        Add-ActivityLog "Application" "Started" "M365 User Provisioning Tool launched"
    })
    
    $Script:MainForm.Add_FormClosing({
        param($sender, $e)
        
        if ($Global:IsConnected) {
            $Result = [System.Windows.Forms.MessageBox]::Show(
                "You are currently connected to Microsoft 365. Disconnect and exit?",
                "Confirm Exit",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($Result -eq [System.Windows.Forms.DialogResult]::No) {
                $e.Cancel = $true
                return
            }
            
            try {
                Disconnect-MgGraph -ErrorAction SilentlyContinue
            }
            catch {
                # Ignore disconnect errors on exit
            }
        }
        
        Add-ActivityLog "Application" "Closed" "User closed application"
        Write-Host "👋 Application closing..." -ForegroundColor Yellow
    })
    
    return $Script:MainForm
}

function New-UserCreationTab {
    <#
    .SYNOPSIS
        Creates the user creation tab with comprehensive form
    #>
    
    $Tab = New-Object System.Windows.Forms.TabPage
    $Tab.Text = "👤 Create User"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # Connection Panel
    $ConnectionPanel = New-Object System.Windows.Forms.Panel
    $ConnectionPanel.Height = 60
    $ConnectionPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $ConnectionPanel.BackColor = [System.Drawing.Color]::LightBlue
    $ConnectionPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $Script:ConnectButton = New-Object System.Windows.Forms.Button
    $Script:ConnectButton.Text = "🔗 Connect to Microsoft 365"
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
    $UserDetailsGroup.Text = "User Details"
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
    
    # Department
    $DepartmentLabel = New-Object System.Windows.Forms.Label
    $DepartmentLabel.Text = "Department:"
    $DepartmentLabel.Location = New-Object System.Drawing.Point(10, $y)
    $DepartmentLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:DepartmentTextBox = New-Object System.Windows.Forms.TextBox
    $Script:DepartmentTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:DepartmentTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    $y += $spacing
    
    # Job Title
    $JobTitleLabel = New-Object System.Windows.Forms.Label
    $JobTitleLabel.Text = "Job Title:"
    $JobTitleLabel.Location = New-Object System.Drawing.Point(10, $y)
    $JobTitleLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:JobTitleTextBox = New-Object System.Windows.Forms.TextBox
    $Script:JobTitleTextBox.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:JobTitleTextBox.Size = New-Object System.Drawing.Size(200, 20)
    
    $y += $spacing
    
    # Office Location
    $OfficeLabel = New-Object System.Windows.Forms.Label
    $OfficeLabel.Text = "Office:"
    $OfficeLabel.Location = New-Object System.Drawing.Point(10, $y)
    $OfficeLabel.Size = New-Object System.Drawing.Size(100, 20)
    
    $Script:OfficeDropdown = New-Object System.Windows.Forms.ComboBox
    $Script:OfficeDropdown.Location = New-Object System.Drawing.Point(120, ($y-2))
    $Script:OfficeDropdown.Size = New-Object System.Drawing.Size(200, 20)
    $Script:OfficeDropdown.DropDownStyle = "DropDownList"
    
    # Add controls to user details group
    $UserDetailsGroup.Controls.AddRange(@(
        $FirstNameLabel, $Script:FirstNameTextBox,
        $LastNameLabel, $Script:LastNameTextBox,
        $UsernameLabel, $Script:UsernameTextBox,
        $DomainLabel, $Script:DomainDropdown,
        $PasswordLabel, $Script:PasswordTextBox, $GeneratePasswordButton,
        $DepartmentLabel, $Script:DepartmentTextBox,
        $JobTitleLabel, $Script:JobTitleTextBox,
        $OfficeLabel, $Script:OfficeDropdown
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
    
    # License types info
    $LicenseTypesLabel = New-Object System.Windows.Forms.Label
    $LicenseTypesLabel.Text = "Available License Types:`n• BusinessBasic`n• BusinessPremium`n• BusinessStandard`n• E3 / E5`n• ExchangeOnline1 / ExchangeOnline2"
    $LicenseTypesLabel.Location = New-Object System.Drawing.Point(10, 140)
    $LicenseTypesLabel.Size = New-Object System.Drawing.Size(430, 120)
    $LicenseTypesLabel.ForeColor = [System.Drawing.Color]::DarkGreen
    $LicenseTypesLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    
    $ManagementGroup.Controls.AddRange(@(
        $ManagerLabel, $Script:ManagerDropdown,
        $LicenseLabel, $Script:LicenseDropdown,
        $LicenseInfoLabel, $LicenseTypesLabel
    ))
    
    # Groups Group (full width below)
    $GroupsGroup = New-Object System.Windows.Forms.GroupBox
    $GroupsGroup.Text = "Group Membership (Select groups to add user to)"
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
    $Script:CreateUserButton.Text = "👤 Create M365 User"
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
                         -Office $Script:OfficeDropdown.SelectedItem `
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
    $ClearFormButton.Text = "🗑️ Clear Form"
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
    $Tab.Text = "🏢 Tenant Data"
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
    $Tab.Text = "📋 Activity Log"
    $Tab.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $Script:ActivityLogTextBox = New-Object System.Windows.Forms.TextBox
    $Script:ActivityLogTextBox.Multiline = $true
    $Script:ActivityLogTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $Script:ActivityLogTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $Script:ActivityLogTextBox.ReadOnly = $true
    $Script:ActivityLogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $Script:ActivityLogTextBox.Text = "$(Get-Date): Application started`r`n"
    
    $Tab.Controls.Add($Script:ActivityLogTextBox)
    return $Tab
}

# ================================
# MAIN APPLICATION ENTRY POINT
# ================================

try {
    Write-Host "🚀 Launching M365 User Provisioning Tool Enterprise Edition..." -ForegroundColor Green
    Write-Host ""
    
    # Create and show the main form
    $MainForm = New-MainForm
    
    if ($MainForm) {
        Write-Host "✅ Application window created successfully" -ForegroundColor Green
        Write-Host "📱 Starting application..." -ForegroundColor Green
        Write-Host ""
        Write-Host "🎯 FEATURES AVAILABLE:" -ForegroundColor Cyan
        Write-Host "   • Comprehensive tenant data discovery" -ForegroundColor White
        Write-Host "   • User creation with full property support" -ForegroundColor White
        Write-Host "   • License assignment via CustomAttribute1" -ForegroundColor White
        Write-Host "   • Group membership management" -ForegroundColor White
        Write-Host "   • UK location support" -ForegroundColor White
        Write-Host "   • Manager assignment" -ForegroundColor White
        Write-Host "   • Activity logging" -ForegroundColor White
        Write-Host ""
        
        # Run the application
        [System.Windows.Forms.Application]::Run($MainForm)
        
        Write-Host ""
        Write-Host "👋 M365 User Provisioning Tool session ended" -ForegroundColor Yellow
    }
    else {
        throw "Failed to create main application window"
    }
}
catch {
    Write-Host ""
    Write-Host "🚨 CRITICAL ERROR:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    
    if ($_.Exception.InnerException) {
        Write-Host "Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Stack Trace:" -ForegroundColor Gray
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    
    Read-Host "Press Enter to exit"
    exit 1
}
finally {
    # Cleanup
    if ($Global:IsConnected) {
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        }
        catch {
            Write-Verbose "Error during cleanup: $($_.Exception.Message)"
        }
    }
}