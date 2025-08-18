# M365 User Provisioning Tool - Claude Code Configuration

## Project Overview

This is a comprehensive PowerShell-based application designed for first-line IT support teams to create and manage Microsoft 365 users without navigating multiple admin portals. The tool provides both single user creation and bulk import capabilities with an intuitive GUI interface.

## Quick Start for 1st Line Support

### Launch the Tool
```bash
# Easy method - just double-click
Start-Tool.bat

# Or from PowerShell 7
pwsh .\M365-UserProvisioning-Enterprise.ps1
```

### First Time Setup
1. Ensure PowerShell 7 is installed: `winget install Microsoft.PowerShell`
2. Run `Install-Prerequisites.ps1` to install required modules
3. Use `Start-Tool.bat` for automatic setup

### 🔄 NEW: Multi-Tenant Support
- **Orange "🔄 Switch Tenant" button** next to Connect button
- **No application restart** required for tenant switching
- **Perfect for MSPs** managing multiple client tenants

## Common Tasks for 1st Line Support

### Creating Single Users
- Use the "User Creation" tab in the GUI
- Fill required fields (FirstName, LastName, Username)
- Generate secure passwords using the built-in generator
- Assign licenses and groups as needed
- Click "Create User" and confirm

### Bulk User Import
- Use the "Bulk Import" tab
- Download template from `Templates\M365_BulkImport_Template.csv`
- Fill in user details following the template format
- Upload CSV and process users in batch

### 🔄 Tenant Management & Switching
- **Click "🔄 Switch Tenant"** - Orange button next to Connect for easy access
- **Complete disconnection** - Clears all cached authentication and tenant data  
- **Fresh authentication** - Forces new login for different tenant
- **Data isolation** - No bleeding of data between tenants
- **MSP workflow** - Perfect for consultants managing multiple clients
- **Refresh tenant data** to get latest licenses, groups, and users
- **Monitor activity** in the "Activity Log" tab with tenant-specific logs

## Key Files & Structure

### Main Application
- `M365-UserProvisioning-Enterprise.ps1` - Main GUI application
- `Start-Tool.bat` - One-click launcher for support teams
- `Install-Prerequisites.ps1` - Automatic module installer

### Modules (PowerShell Libraries)
- `Modules\M365.Authentication\` - Handles M365 connections
- `Modules\M365.GUI\` - User interface components  
- `Modules\M365.UserManagement\` - User creation logic
- `Modules\M365.ExchangeOnline\` - Exchange mailbox operations

### Templates & Data
- `Templates\M365_BulkImport_Template.csv` - Bulk import template
- `Config\` - Configuration files (auto-generated)
- `Logs\` - Application logs and activity history (M365_Final_Log_*.txt)
- `Tests\` - Testing scripts for validation and debugging

### 🔄 Switch Tenant Feature Components
- **Main Script** - Switch Tenant button and disconnection logic
- **M365.Authentication Module** - Enhanced disconnect with cache clearing
- **M365.ExchangeOnline Module** - Exchange-specific cache clearing
- **Comprehensive Data Clearing** - Both Microsoft Graph and Exchange Online

## Troubleshooting Commands

### Check PowerShell Version
```powershell
pwsh --version  # Should be 7.0 or higher
```

### Test Module Installation
```powershell
Get-Module Microsoft.Graph.Authentication -ListAvailable
Get-Module ExchangeOnlineManagement -ListAvailable
```

### Manual Module Installation (if needed)
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
```

### Check M365 Connection
```powershell
Get-MgContext  # Should show connected tenant
```

## Security & Permissions

### Required M365 Roles
- Global Administrator OR
- User Administrator + Exchange Administrator + Groups Administrator

### Authentication
- Uses Microsoft Graph API with OAuth2
- No passwords stored locally
- Session-based authentication with automatic timeout

## CSV Template Format

### Required Fields
- DisplayName, UserPrincipalName, FirstName, LastName

### Optional Fields  
- Department, JobTitle, Office, Manager, LicenseType, Groups, Password, ForcePasswordChange

### Example Row
```csv
John Smith,john.smith@company.com,John,Smith,IT,Developer,United Kingdom - London,manager@company.com,BusinessPremium,"IT Team,Developers",,true
```

## Testing & Validation

### Run Tests
```powershell
# Unit tests
Invoke-Pester .\Tests\Unit\

# Integration tests  
Invoke-Pester .\Tests\Integration\
```

### Test Mode (No Changes)
```powershell
pwsh .\M365-UserProvisioning-Enterprise.ps1 -TestMode
```

## Common Error Solutions

### "PowerShell 7 not found"
Install PowerShell 7: `winget install Microsoft.PowerShell`

### "Module installation failed"
Run as Administrator or use `-Scope CurrentUser` parameter

### "Authentication failed"
Check Global Admin permissions and complete MFA prompts

### "User creation failed"
- Username already exists
- Invalid email format  
- Insufficient licenses
- Check Activity Log for details

## Multi-Tenant Support

The tool supports switching between different M365 tenants:
1. Connect to first tenant
2. Work with users as needed
3. Click "Switch Tenant" to change organizations
4. Authenticate to new tenant
5. Continue working with new tenant data

Perfect for Managed Service Providers (MSPs) and consultants.

## Log Files

### Activity Logs
- Real-time operation logging in GUI
- Export capability for audit purposes
- Located in `Logs\` directory

### Error Logs
- Detailed error information for troubleshooting
- Timestamp records for tracking issues
- Format: `error_YYYYMMDD_HHMMSS.log`

## Integration Notes

- Can be called from other PowerShell scripts
- Supports command-line parameters for automation
- Compatible with scheduled tasks for bulk operations
- Audit logs can be exported for compliance systems

## Best Practices for 1st Line Support

1. **Always use the template** for bulk imports
2. **Test with single user** before bulk operations  
3. **Check Activity Log** for any errors
4. **Verify tenant permissions** before starting
5. **Use generated passwords** for security
6. **Confirm operations** using the built-in dialogs

## Support Resources

- **USER-GUIDE.md** - Complete documentation
- **Activity Log tab** - Real-time error tracking
- **GitHub Issues** - Bug reports and feature requests
- **CSV Template** - Proper format examples