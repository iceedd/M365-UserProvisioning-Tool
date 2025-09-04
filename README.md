# M365 User Provisioning Tool

[![PowerShell](https://img.shields.io/badge/PowerShell-7.0+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://github.com/PowerShell/PowerShell)

> **Enterprise-grade Microsoft 365 user provisioning tool with multi-tenant support designed for IT support teams and Managed Service Providers (MSPs).**

## 🚀 What This Tool Does

A comprehensive PowerShell-based GUI application that **eliminates the complexity** of Microsoft 365 user management by providing a single, intuitive interface for all user provisioning tasks across multiple tenants.

### ✨ Key Features

- **🎯 Single User Creation** - Complete user setup with all M365 properties
- **📊 Bulk CSV Import** - Create multiple users from spreadsheet data  
- **🔄 Multi-Tenant Support** - Switch between M365 tenants without restarting
- **⚡ Real-time Validation** - Instant username availability checking
- **🛡️ Security First** - Uses official Microsoft Graph API with OAuth
- **📝 Activity Logging** - Complete audit trail for compliance
- **🏢 MSP-Ready** - Perfect for managing multiple client tenants
- **💡 Dynamic Groups** - Works with your existing M365 dynamic group setup
- **🔍 Exchange Integration** - Mailbox, distribution list, and shared mailbox management

## 🎯 Who This Is For

### Primary Users
- **IT Support Teams** - First-line technicians creating users daily
- **Managed Service Providers (MSPs)** - Managing multiple client tenants  
- **Help Desk Staff** - Non-PowerShell users needing simple user creation
- **System Administrators** - Looking to standardize user provisioning processes

### Perfect For
- **Multi-client environments** (up to 150 users per tenant)
- **Organizations using dynamic groups** for license assignment
- **Teams wanting to eliminate M365 admin portal complexity**
- **Compliance-focused environments** requiring audit trails

## 🚀 Quick Start Guide

### Option 1: One-Click Setup (Recommended)
1. **Download** this repository as ZIP
2. **Extract** to any folder (e.g., `C:\M365-Tool\`)  
3. **Double-click** `Start-Tool.bat`
4. **Follow prompts** to install PowerShell 7 and required modules
5. **Launch** and connect to your first Microsoft 365 tenant!

### Option 2: Manual Setup
```powershell
# Install PowerShell 7 (if needed)
winget install Microsoft.PowerShell

# Install required modules
.\Install-Prerequisites.ps1

# Launch the tool
.\M365-UserProvisioning-Enterprise.ps1
```

### First Connection
1. Click **"Connect to Microsoft 365"**
2. **Sign in** with Global Admin or User Admin credentials
3. **Wait** for tenant discovery to complete (finds users, groups, licenses)
4. **Start creating users** with the intuitive interface!

## 🔄 Multi-Tenant Management (MSP Feature)

### Seamless Tenant Switching
Perfect for MSPs managing multiple client environments:

1. **Connect** to first client tenant
2. **Work** with users as normal  
3. **Click "🔄 Switch Tenant"** (orange button)
4. **Confirm** disconnection (clears all cached data)
5. **Connect** to different client tenant
6. **Continue** working with complete data isolation

### Security & Isolation
- ✅ **Aggressive cache clearing** prevents data bleeding between tenants
- ✅ **Fresh authentication** required for each tenant
- ✅ **Separate audit logs** per tenant for compliance
- ✅ **No credential storage** - always prompts for authentication

## 📋 Core Functionality

### Single User Creation
- **Complete user profiles** with all M365 attributes
- **Real-time username validation** 
- **Secure password generation** or custom passwords
- **Dynamic group assignment** (works with your existing setup)
- **License assignment** via CustomAttribute1
- **Manager assignment** and organizational hierarchy

### Bulk Operations  
- **CSV import** with template generation
- **Progress tracking** with detailed status updates
- **Dry-run testing** before actual user creation
- **Error handling** with detailed reporting
- **Rollback capabilities** for failed operations

### Exchange Integration
- **Mailbox provisioning** 
- **Distribution list management**
- **Shared mailbox permissions**
- **Mail-enabled security groups**
- **Real-time Exchange data discovery**

## 📁 Repository Structure

```
M365-UserProvisioning-Tool/
├── M365-UserProvisioning-Enterprise.ps1    # 🎯 Main application  
├── Start-Tool.bat                          # 🚀 One-click launcher
├── Install-Prerequisites.ps1               # ⚙️ Automated setup
├── Modules/                                # 📦 PowerShell modules
│   ├── M365.Authentication/               #   🔐 Authentication & tenant switching  
│   ├── M365.GUI/                         #   🖥️ User interface components
│   ├── M365.UserManagement/              #   👤 User creation logic
│   └── M365.ExchangeOnline/              #   📧 Exchange operations
├── Templates/                             # 📄 CSV templates for bulk import
├── Tests/                                 # 🧪 Testing and validation scripts  
├── Logs/                                  # 📊 Activity logs and audit trails
└── Documentation/                         # 📚 Additional guides and help
```

## ⚙️ System Requirements

### Required
- **PowerShell 7.0+** (⚠️ Windows PowerShell 5.1 will NOT work)
- **Windows 10/11** or **Windows Server 2019/2022**
- **Microsoft 365 tenant** with admin permissions
- **Internet connection** for Microsoft Graph API access

### Permissions Needed
- **Global Administrator** (full functionality)
- **User Administrator** (user creation only)
- **Exchange Administrator** (for Exchange features)

### Auto-Installed Dependencies
The tool automatically installs these Microsoft modules:
- Microsoft.Graph.Authentication
- Microsoft.Graph.Users  
- Microsoft.Graph.Groups
- Microsoft.Graph.Identity.DirectoryManagement
- ExchangeOnlineManagement

## 🛠️ Advanced Usage

### For MSPs Managing Multiple Clients
- Use **Switch Tenant** for seamless client switching
- **Standardized workflows** across all client environments
- **Separate audit logs** for each client
- **Dynamic discovery** eliminates client-specific configuration

### Integration with Dynamic Groups
- Works perfectly with **M365 dynamic groups**
- **License assignment** via CustomAttribute1 triggers dynamic group rules
- **No hardcoded client configurations** needed
- **Flexible group assignment** based on user attributes

### Bulk Import Features
- **CSV template generation** for consistent data format
- **Data validation** before import starts
- **Progress tracking** with real-time updates
- **Error reporting** with detailed failure information
- **Dry-run mode** for testing before production

## 📊 Enterprise Features

### Security & Compliance
- **Official Microsoft Graph API** - No third-party dependencies
- **OAuth authentication** - No password storage
- **Complete audit logging** - Track all user creation activities  
- **Data isolation** - Aggressive tenant data clearing
- **Local processing** - No data sent to external services

### Reliability & Performance  
- **Robust error handling** - Graceful failure recovery
- **Connection management** - Automatic token refresh
- **Memory management** - Efficient handling of large datasets
- **Modular architecture** - Easy to maintain and extend

## 🆘 Support & Troubleshooting

### Getting Help
- **📖 User Guide**: See `USER-GUIDE.md` for detailed instructions
- **🐛 Issues**: Report bugs via GitHub Issues
- **💡 Questions**: Check existing issues or create new ones

### Common Issues
- **PowerShell Version**: Ensure you're using PowerShell 7.0+, not Windows PowerShell 5.1
- **Module Installation**: Run `Install-Prerequisites.ps1` as administrator if modules fail to install
- **Tenant Permissions**: Verify you have User Administrator or Global Administrator role

## 🤝 Contributing

This tool is designed for enterprise use. If you encounter issues or have suggestions:

1. **Check existing issues** on GitHub
2. **Create detailed bug reports** with logs from the `Logs/` folder
3. **Test thoroughly** before submitting pull requests
4. **Follow PowerShell best practices** in any contributions

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

### 🌟 Why Choose This Tool Over M365 Admin Portal?

| M365 Admin Portal | This Tool |
|-------------------|-----------|
| ❌ Multiple portals to navigate | ✅ Single unified interface |
| ❌ Complex setup requiring training | ✅ Intuitive GUI anyone can use |
| ❌ Manual entry prone to errors | ✅ Built-in validation and safety checks |
| ❌ No bulk operations | ✅ CSV import with progress tracking |
| ❌ Tenant switching requires new browser sessions | ✅ Seamless tenant switching |
| ❌ Limited audit capabilities | ✅ Complete activity logging |

**Transform your M365 user management from complex to simple. Perfect for IT teams who want enterprise-grade functionality without the enterprise-level complexity.**