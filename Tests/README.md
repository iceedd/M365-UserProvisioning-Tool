# Testing Scripts for M365 User Provisioning Tool

This folder contains various testing and debugging scripts to help validate and troubleshoot the M365 User Provisioning Tool functionality.

## 🧪 Testing Scripts

### Switch Tenant Testing
- **`Test-SwitchTenantButton.ps1`** - Isolated test of the Switch Tenant button UI component
- **`Debug-TenantSwitching.ps1`** - Interactive monitoring tool for tenant switching process
- **`Test-TenantDataClearing.ps1`** - Validates that tenant data clearing functions work correctly

### Application Testing
- **`Quick-Test-MainScript.ps1`** - Quick validation test for the main application script
- **`Debug-GUI-Module.ps1`** - Debug script to examine GUI module loading and button creation
- **`Force-Reload-And-Run.ps1`** - Forces complete module reload and runs main application (for cache issues)

## 📁 Organized Testing

### Unit Tests (`Unit/`)
- **`M365.Authentication.Tests.ps1`** - Unit tests for authentication module
- **`M365.UserManagement.Tests.ps1`** - Unit tests for user management functions

### Integration Tests (`Integration/`)
- Ready for end-to-end testing scenarios

## 🚀 How to Use These Scripts

### For Switch Tenant Validation:
1. **Run the main application**
2. **In separate PowerShell window**: `.\Tests\Debug-TenantSwitching.ps1`
3. **Follow the step-by-step prompts** to monitor tenant switching

### For UI Button Testing:
```powershell
# Test button creation in isolation
.\Tests\Test-SwitchTenantButton.ps1

# Debug GUI module loading
.\Tests\Debug-GUI-Module.ps1
```

### For Data Clearing Validation:
```powershell
# Test tenant data clearing functions
.\Tests\Test-TenantDataClearing.ps1
```

### For Module Cache Issues:
```powershell
# Force reload all modules and run application
.\Tests\Force-Reload-And-Run.ps1
```

## 🔍 Debugging Workflow

### Problem: Switch Tenant button not visible
1. Run `Debug-GUI-Module.ps1` to check GUI module loading
2. Run `Test-SwitchTenantButton.ps1` to test button creation in isolation
3. Use `Force-Reload-And-Run.ps1` to clear module cache

### Problem: Previous tenant data still showing
1. Run `Debug-TenantSwitching.ps1` to monitor data clearing
2. Run `Test-TenantDataClearing.ps1` to validate clearing functions
3. Check console output for cache clearing messages

### Problem: Module loading issues
1. Use `Force-Reload-And-Run.ps1` to clear PowerShell module cache
2. Run `Quick-Test-MainScript.ps1` to validate basic functionality

## 📊 Expected Outputs

### Successful Switch Tenant Test:
```
📊 Current Tenant Data State
   AcceptedDomains: 0 items
   AvailableUsers: 0 items
   AvailableGroups: 0 items
   SharedMailboxes: 0 items
   DistributionLists: 0 items
   IsConnected: False
✅ Disconnected and data cleared
```

### Successful Button Test:
```
✅ FOUND Switch Tenant button!
   Size: {Width=160, Height=35}
   Location: {X=200,Y=15}
   Enabled: False
   BackColor: Color [Orange]
```

## 🛠️ Development Notes

These scripts were created during the development of the Switch Tenant functionality to:
- Validate UI component creation
- Monitor data clearing processes
- Debug module loading issues
- Provide interactive testing capabilities

They serve as both validation tools and examples for future testing development.