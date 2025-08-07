# Test-LiveConnection.ps1 - Live M365 Connection Test
# Run this to test actual connection to your M365 tenant

Write-Host "🔗 M365 User Provisioning Tool - Live Connection Test" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan

try {
    # Import modules
    Write-Host "📦 Importing modules..." -ForegroundColor Yellow
    Import-Module .\Modules\M365.Authentication\M365.Authentication.psd1 -Force
    Import-Module .\Modules\M365.UserManagement\M365.UserManagement.psd1 -Force
    Write-Host "   ✅ Modules imported successfully" -ForegroundColor Green

    # Test connection
    Write-Host "`n🔐 Attempting to connect to Microsoft Graph..." -ForegroundColor Yellow
    Write-Host "   📱 Browser window will open for authentication..." -ForegroundColor Cyan
    
    $ConnectionResult = Connect-ToMicrosoftGraph
    
    if ($ConnectionResult.Success) {
        Write-Host "   ✅ Successfully connected to Microsoft Graph!" -ForegroundColor Green
        Write-Host "   🏢 Tenant ID: $($ConnectionResult.TenantId)" -ForegroundColor Cyan
        Write-Host "   👤 Account: $($ConnectionResult.Account)" -ForegroundColor Cyan
        Write-Host "   🌍 Environment: $($ConnectionResult.Environment)" -ForegroundColor Cyan
        
        # Show tenant data summary
        Write-Host "`n📊 Tenant Discovery Results:" -ForegroundColor Yellow
        $Summary = $ConnectionResult.TenantData
        if ($Summary) {
            Write-Host "   👥 Users: $($Summary.AvailableUsers.Count)" -ForegroundColor White
            Write-Host "   🏘️ Groups: $($Summary.AvailableGroups.Count)" -ForegroundColor White
            Write-Host "   📧 Distribution Lists: $($Summary.DistributionLists.Count)" -ForegroundColor White
            Write-Host "   📄 Licenses: $($Summary.AvailableLicenses.Count)" -ForegroundColor White
            Write-Host "   🌐 Domains: $($Summary.AcceptedDomains.Count)" -ForegroundColor White
        }
        
        # Show Exchange Online status
        if ($ConnectionResult.ExchangeOnlineConnected) {
            Write-Host "   📬 Exchange Online: ✅ Connected" -ForegroundColor Green
        } else {
            Write-Host "   📬 Exchange Online: ⚠️ Not connected (manual tasks only)" -ForegroundColor Yellow
        }
        
        # Test creating a sample password
        Write-Host "`n🔑 Testing password generation:" -ForegroundColor Yellow
        $SamplePasswords = @()
        for ($i = 1; $i -le 3; $i++) {
            $SamplePasswords += New-SecurePassword
        }
        $SamplePasswords | ForEach-Object { Write-Host "   🔐 Sample password: $_" -ForegroundColor Cyan }
        
        # Show some sample domains for user creation
        $TenantData = Get-M365TenantData
        if ($TenantData.AcceptedDomains.Count -gt 0) {
            Write-Host "`n📧 Available domains for user creation:" -ForegroundColor Yellow
            $TenantData.AcceptedDomains | ForEach-Object { 
                $DefaultText = if ($_.IsDefault) { " (Default)" } else { "" }
                Write-Host "   🌐 $($_.DomainName)$DefaultText" -ForegroundColor Cyan 
            }
        }
        
        # Show some sample groups
        if ($TenantData.AvailableGroups.Count -gt 0) {
            Write-Host "`n👥 Sample groups (first 5):" -ForegroundColor Yellow
            $TenantData.AvailableGroups | Select-Object -First 5 | ForEach-Object {
                Write-Host "   📁 $($_.DisplayName) [$($_.GroupType)]" -ForegroundColor Cyan
            }
        }
        
        Write-Host "`n✅ CONNECTION TEST SUCCESSFUL!" -ForegroundColor Green
        Write-Host "🎯 Your M365 User Provisioning Tool is ready for production use!" -ForegroundColor Magenta
        
        # Offer to disconnect
        $Disconnect = Read-Host "`n🔌 Disconnect from M365 now? (y/N)"
        if ($Disconnect -eq 'y' -or $Disconnect -eq 'Y') {
            Write-Host "🔌 Disconnecting..." -ForegroundColor Yellow
            $DisconnectResult = Disconnect-FromMicrosoftGraph
            if ($DisconnectResult.Success) {
                Write-Host "   ✅ Disconnected successfully" -ForegroundColor Green
            } else {
                Write-Host "   ⚠️ Disconnect completed with warnings: $($DisconnectResult.Message)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "   ℹ️ Staying connected for further testing" -ForegroundColor Blue
        }
        
    } else {
        Write-Host "   ❌ Connection failed: $($ConnectionResult.Message)" -ForegroundColor Red
        Write-Host "   🔧 Troubleshooting:" -ForegroundColor Yellow
        Write-Host "      1. Check internet connection" -ForegroundColor White
        Write-Host "      2. Verify your account has proper permissions" -ForegroundColor White
        Write-Host "      3. Ensure Microsoft Graph PowerShell SDK is up to date" -ForegroundColor White
    }
    
} catch {
    Write-Host "❌ Critical error during connection test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "📍 Error location: Line $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Gray
    Write-Host "🔧 Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Gray
}

Write-Host "`n🎯 Next Steps:" -ForegroundColor Yellow
Write-Host "1. If connection successful: Extract GUI functions from legacy script" -ForegroundColor White
Write-Host "2. If connection failed: Review error messages and fix authentication" -ForegroundColor White
Write-Host "3. Complete the GUI module to have full functionality" -ForegroundColor White