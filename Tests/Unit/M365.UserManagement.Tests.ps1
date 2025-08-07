BeforeAll {
    Import-Module "..\..\Modules\M365.UserManagement"
}

Describe "New-M365User" {
    It "Creates user with minimal params" {
        Mock New-MgUser { return @{Id="test123"} }
        
        $user = New-M365User -DisplayName "Test" `
            -UserPrincipalName "test@domain.com" `
            -Password "P@ssw0rd!"
            
        $user.Id | Should -Be "test123"
    }
}