# M365 User Provisioning Tool

## Overview

A comprehensive PowerShell-based application designed to simplify Microsoft 365 user creation and management for IT support teams. This tool eliminates the need to navigate multiple M365 admin portals by providing a single, intuitive interface for all user provisioning tasks.

## Who Is This For?

### Primary Audience: **First-Line IT Support Teams**
- **Help desk technicians** who need to create users quickly and accurately
- **Junior IT staff** who may not be familiar with PowerShell or M365 admin centers
- **Service desk teams** handling user onboarding requests
- **IT departments** looking to standardize and streamline user creation processes

### Use Cases:
- **New employee onboarding** - Create users with proper licenses, groups, and permissions
- **Bulk user imports** - Handle multiple new hires from HR systems via CSV
- **Contractor/temporary user setup** - Quick user creation with appropriate access
- **User management tasks** - Standardized approach across the organization

## What This Tool Does

### 🎯 **Core Functionality**
- **Single User Creation** - Complete user setup through an intuitive GUI
- **Bulk CSV Import** - Create multiple users from spreadsheet data
- **Real-time Validation** - Instant username availability checking
- **Safety Features** - Confirmation dialogs prevent accidental user creation

### 🔧 **Key Features**
- **No M365 Portal Navigation** - Everything from one interface
- **Intelligent Tenant Discovery** - Automatically finds available licenses, groups, and domains
- **License Assignment** - Assigns appropriate M365 licenses during user creation
- **Group Membership** - Add users to security and distribution groups
- **Exchange Integration** - Mailbox setup and distribution list management
- **Activity Logging** - Track all user creation activities for audit purposes

### 💼 **Business Benefits**
- **Reduces Training Time** - Simple GUI instead of complex admin portals
- **Prevents Errors** - Built-in validation and confirmation dialogs
- **Increases Efficiency** - Bulk operations and standardized workflows
- **Ensures Consistency** - Same process every time, reducing variations
- **Improves Audit** - Complete activity logging for compliance

## Technical Requirements

- **PowerShell 7.0+** (Critical - Windows PowerShell 5.1 will not work)
- **Windows 10/11** or **Windows Server 2019/2022**
- **Microsoft 365 Global Administrator** or **User Administrator** permissions
- **Internet connection** for M365 API access

## Quick Start

### For First-Line Teams (Easiest)
1. **Download** the ZIP file from GitHub
2. **Extract** to a folder (e.g., `C:\M365-Tool\`)
3. **Install PowerShell 7**: `winget install Microsoft.PowerShell`
4. **Double-click** `Start-Tool.bat`
5. **Connect** to Microsoft 365 and start creating users!

### What You Get
- **Automated setup** - All required modules install automatically
- **User-friendly interface** - No PowerShell knowledge required
- **Step-by-step guidance** - Clear instructions and error messages
- **Professional appearance** - Looks and feels like enterprise software

## Repository Structure

```
M365-UserProvisioning-Tool/
├── M365-UserProvisioning-Enterprise.ps1    # Main application
├── Install-Prerequisites.ps1               # Automatic setup script
├── Start-Tool.bat                          # One-click launcher
├── USER-GUIDE.md                          # Complete documentation
├── Modules/                                # PowerShell modules
│   ├── M365.Authentication/               # M365 connection handling
│   ├── M365.GUI/                         # User interface
│   ├── M365.UserManagement/              # User creation logic
│   └── M365.ExchangeOnline/              # Exchange operations
├── Templates/                             # CSV templates for bulk import
└── Tests/                                # Unit tests
```

## Why Choose This Tool?

### Instead of M365 Admin Center:
- ❌ **Multiple portals** to navigate (Users, Exchange, Groups, Licenses)
- ❌ **Complex interface** requiring training
- ❌ **Easy to make mistakes** with manual entry
- ❌ **No bulk operations** or automation

### With M365 User Provisioning Tool:
- ✅ **Single interface** for all user operations
- ✅ **Intuitive GUI** anyone can use
- ✅ **Built-in validation** prevents errors
- ✅ **Bulk operations** and CSV import
- ✅ **Real-time feedback** and confirmation dialogs
- ✅ **Activity logging** for audit trails

## Support & Documentation

- **📖 Complete User Guide**: See `USER-GUIDE.md` for detailed instructions
- **🚀 Quick Setup**: Use `Start-Tool.bat` for automatic installation
- **❓ Issues & Questions**: Check GitHub Issues for support
- **🔧 Technical Details**: See `.claude/CLAUDE.md` for development info

## Security & Compliance

- **Secure Authentication** - Uses Microsoft's official Graph API
- **No Password Storage** - Generates secure passwords on-demand
- **Audit Logging** - Complete activity tracking
- **Minimal Permissions** - Only requires necessary M365 admin rights
- **Local Processing** - No data sent to third-party services

---

**Perfect for organizations looking to streamline M365 user management while maintaining security and reducing support burden on IT teams.**