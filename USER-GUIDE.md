# M365 User Provisioning Tool - Complete User Guide

## Overview
The M365 User Provisioning Tool is a comprehensive PowerShell-based application designed for first-line support teams to easily create and manage Microsoft 365 users without navigating multiple admin portals. This tool provides both single user creation and bulk import capabilities with an intuitive GUI interface.

## Table of Contents
- [System Requirements](#system-requirements)
- [Download and Setup](#download-and-setup)
- [First-Time Installation](#first-time-installation)
- [Quick Start Guide](#quick-start-guide)
- [Using the Tool](#using-the-tool)
- [Troubleshooting](#troubleshooting)
- [Advanced Usage](#advanced-usage)
- [Support](#support)

## System Requirements

### Mandatory Requirements
- **Windows 10/11** or **Windows Server 2019/2022**
- **PowerShell 7.0 or later** (will NOT work with Windows PowerShell 5.1)
- **Internet connection** for downloading modules and connecting to Microsoft 365
- **Microsoft 365 Global Administrator** or **User Administrator** permissions

### Automatic Dependencies (installed by the tool)
- Microsoft Graph PowerShell SDK V2.0+
- Exchange Online PowerShell V3.0+
- Additional Microsoft Graph modules for users, groups, and sites

## Download and Setup

### Step 1: Download from GitHub
1. Go to the GitHub repository: `[https://github.com/iceedd/M365-UserProvisioning-Tool]`
2. Click the green **"Code"** button
3. Select **"Download ZIP"**
4. Extract the ZIP file to a folder (e.g., `C:\M365-Tool\`)

### Step 2: Install PowerShell 7
**CRITICAL: This tool requires PowerShell 7. Windows PowerShell 5.1 will not work.**

#### Option A: Using Windows Package Manager (Recommended)
1. Open **Command Prompt** as Administrator
2. Run: `winget install Microsoft.PowerShell`
3. Restart your command prompt/terminal

#### Option B: Manual Download
1. Visit: https://github.com/PowerShell/PowerShell/releases
2. Download the latest Windows installer (.msi file)
3. Run the installer with default settings
4. Restart your computer

#### Verify Installation
Open a new command prompt and run:
```cmd
pwsh --version
```
You should see version 7.0 or higher.

## First-Time Installation

### Easy Installation (Recommended for First-Line Teams)
1. **Navigate to the tool folder** in File Explorer
2. **Double-click `Start-Tool.bat`**
3. The script will automatically:
   - Check PowerShell 7 is installed
   - Install all required modules
   - Launch the M365 tool

### Manual Installation (Advanced Users)
1. **Open PowerShell 7** (not Windows PowerShell):
   ```cmd
   pwsh
   ```

2. **Navigate to the tool directory**:
   ```powershell
   cd "C:\M365-Tool"
   ```

3. **Install prerequisites**:
   ```powershell
   .\Install-Prerequisites.ps1
   ```

4. **Launch the tool**:
   ```powershell
   .\M365-UserProvisioning-Enterprise.ps1
   ```

## Quick Start Guide

### First Launch
1. **Double-click `Start-Tool.bat`** or run the PowerShell command
2. **Connect to Microsoft 365**:
   - Click **"Connect"** button
   - Sign in with Global Administrator credentials
   - Grant permissions when prompted
3. **Wait for tenant discovery** (this may take 1-2 minutes)
4. **Start creating users!**

### Creating Your First User
1. Go to the **"User Creation"** tab
2. Fill in the required fields:
   - **First Name**: User's first name
   - **Last Name**: User's last name
   - **Username**: Will auto-generate, but can be modified
   - **Password**: Click "Generate" for secure password
   - **Department**: User's department
   - **Job Title**: User's role
   - **Office**: Select from dropdown (populated from your tenant)
3. **Select Manager** (optional): Choose from existing users
4. **Choose License**: Select appropriate license type
5. **Select Groups** (optional): Add user to security/distribution groups
6. Click **"Create User"**

### Bulk Import Users
1. Go to the **"Bulk Import"** tab
2. **Download the template**:
   - Template file: `Templates\M365_BulkImport_Template.csv`
   - Open in Excel or text editor
3. **Fill in user details** following the template format
4. **Upload and process**:
   - Select your completed CSV file
   - Review the preview
   - Click "Import Users"
   - Monitor progress in real-time

## Using the Tool

### Main Interface Tabs

#### 1. User Creation Tab
- **Single user creation** with full property support
- **Real-time validation** of usernames and emails
- **Password generation** with secure random passwords
- **Manager assignment** from existing users
- **Group membership** selection (security and distribution groups)
- **License assignment** based on available licenses

#### 2. Bulk Import Tab
- **CSV-based mass user creation**
- **Progress tracking** with detailed status updates
- **Error reporting** for failed user creations
- **Rollback capability** if issues occur
- **Template-based** data entry

#### 3. Tenant Data Tab
- **Live tenant information** display
- **User statistics** and counts
- **Available licenses** and usage
- **Group listings** (security, distribution, mail-enabled)
- **SharePoint sites** inventory
- **Exchange mailbox** information

#### 4. Activity Log Tab
- **Real-time operation logging**
- **Success/failure tracking**
- **Timestamp records**
- **Export capability** for audit purposes

### Connection Management
- **Connect Button**: Establishes connection to Microsoft 365
- **Switch Tenant Button**: Disconnect from current tenant and connect to a different one
- **Refresh Data**: Updates tenant information
- **Status Indicator**: Shows current connected tenant name

### Multi-Tenant Management
The tool supports switching between different Microsoft 365 tenants without restarting:

1. **Connect to initial tenant** using the Connect button
2. **Work with users** in that tenant as needed
3. **Click "Switch Tenant"** when you need to access a different organization
4. **Confirm disconnection** - all current data will be cleared
5. **Click Connect** to authenticate to the new tenant
6. **Continue working** with the new tenant's data

**Perfect for:**
- **Managed Service Providers (MSPs)** managing multiple client tenants
- **Organizations with multiple M365 environments** (dev/test/prod)
- **Consultants** working across different client organizations

## Troubleshooting

### Common Issues and Solutions

#### "PowerShell 7 is not available"
**Problem**: Script says PowerShell 7 is missing
**Solution**: 
1. Install PowerShell 7 using `winget install Microsoft.PowerShell`
2. Restart your terminal/command prompt
3. Verify with `pwsh --version`

#### "Module installation failed"
**Problem**: Prerequisites script can't install modules
**Solutions**:
1. **Run as Administrator**: Right-click Command Prompt → "Run as administrator"
2. **Check internet connection**: Ensure access to PowerShell Gallery
3. **Corporate firewall**: Contact IT to allow PowerShell Gallery access
4. **Manual installation**: Install modules individually:
   ```powershell
   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
   Install-Module ExchangeOnlineManagement -Scope CurrentUser
   ```

#### "Authentication failed"
**Problem**: Can't connect to Microsoft 365
**Solutions**:
1. **Check credentials**: Ensure Global Administrator permissions
2. **Multi-factor authentication**: Complete MFA prompts
3. **Conditional Access**: May block automated tools - contact IT
4. **Try different browser**: Clear cache or try incognito mode

#### "User creation failed"
**Problem**: Individual users fail to create
**Common causes**:
- Username already exists
- Invalid email format
- Insufficient licenses available
- Missing required fields
- Network connectivity issues

#### "Bulk import errors"
**Problem**: CSV import fails
**Solutions**:
1. **Check CSV format**: Use provided template exactly
2. **Validate data**: Ensure no empty required fields
3. **Character encoding**: Save CSV as UTF-8
4. **Line endings**: Use Windows line endings (CRLF)

### Getting Help
1. **Check Activity Log tab** for detailed error messages
2. **Review CSV template** for proper format
3. **Verify tenant permissions** with your IT administrator
4. **Test with single user** before bulk operations

## Advanced Usage

### Command-Line Mode
For automated scenarios or scripting:
```powershell
# No GUI mode
pwsh .\M365-UserProvisioning-Enterprise.ps1 -NoGUI

# Test mode (no actual changes)
pwsh .\M365-UserProvisioning-Enterprise.ps1 -TestMode
```

### Configuration Files
- **Tenant settings**: Stored automatically after first connection
- **Default values**: Can be pre-configured for your organization
- **Templates**: Customize CSV template for your needs

### Integration
- **PowerShell scripts**: Can be called from other automation tools
- **Scheduled tasks**: Set up regular bulk imports
- **Audit logs**: Export activity logs for compliance

## CSV Template Format

The bulk import uses this CSV structure:

```csv
DisplayName,UserPrincipalName,FirstName,LastName,Department,JobTitle,Office,Manager,LicenseType,Groups,Password,ForcePasswordChange
John Smith,john.smith@company.com,John,Smith,IT,Developer,United Kingdom - London,manager@company.com,BusinessPremium,"IT Team,Developers",,true
```

### Required Fields
- **DisplayName**: Full name as it appears in directory
- **UserPrincipalName**: Complete email address/login
- **FirstName**: Given name
- **LastName**: Surname

### Optional Fields
- **Department**: User's department
- **JobTitle**: Role/position
- **Office**: Physical location (must match available offices)
- **Manager**: Email address of manager (must exist in tenant)
- **LicenseType**: License SKU name
- **Groups**: Comma-separated list of group names
- **Password**: Leave empty for auto-generated secure passwords
- **ForcePasswordChange**: true/false for password change on first login

## Security Notes

### Permissions Required
- **Global Administrator** or **User Administrator** role
- **Exchange Administrator** (for mailbox operations)
- **Groups Administrator** (for group assignments)

### Best Practices
1. **Use dedicated service account** with minimal required permissions
2. **Enable audit logging** for all user creation activities
3. **Regular password rotation** for service accounts
4. **Monitor activity logs** for unusual patterns
5. **Test in development tenant** before production use

### Data Protection
- **No passwords stored**: Tool generates secure passwords only
- **Encrypted connections**: All API calls use HTTPS
- **Local logs only**: No data sent to third parties
- **Session management**: Automatic disconnection on inactivity

## Support

### Self-Help Resources
1. **Activity Log**: Check the tool's Activity Log tab for detailed error messages
2. **CSV Template**: Use the provided template exactly as formatted
3. **PowerShell 7**: Ensure you're using PowerShell 7, not Windows PowerShell 5.1

### Getting Technical Support
1. **Check GitHub Issues**: Search existing issues for solutions
2. **Create GitHub Issue**: Report bugs or feature requests
3. **Include Information**:
   - PowerShell version (`pwsh --version`)
   - Error messages from Activity Log
   - Steps to reproduce the issue
   - Operating system version

### Contact Information
- **GitHub Repository**: `[YOUR-GITHUB-URL]`
- **Issues/Bug Reports**: `[YOUR-GITHUB-URL]/issues`
- **Documentation**: This guide and CLAUDE.md in the repository

---

## Quick Reference Commands

```cmd
REM Install PowerShell 7
winget install Microsoft.PowerShell

REM Launch tool (easy method)
Start-Tool.bat

REM Manual PowerShell commands
pwsh .\Install-Prerequisites.ps1
pwsh .\M365-UserProvisioning-Enterprise.ps1
```

**Remember**: Always use PowerShell 7 (`pwsh`) not Windows PowerShell (`powershell`)!