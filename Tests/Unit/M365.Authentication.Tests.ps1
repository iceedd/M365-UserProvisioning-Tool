Describe "M365.Authentication Module" {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\Modules\M365.Authentication" -Force
    }

    It "Exports all 9 functions" {
        (Get-Command -Module M365.Authentication).Count | Should -Be 9
    }

    It "Connects to Graph without errors" -Skip:(!$env:CI) {
        { Connect-ToMicrosoftGraph -ErrorAction Stop } | Should -Not -Throw
    }

    It "Gets tenant data after connection" -Skip:(!$env:CI) {
        $data = Get-M365TenantData
        $data.AcceptedDomains | Should -Not -BeNullOrEmpty
    }
}
