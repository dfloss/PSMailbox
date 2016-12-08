$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
Ipmo PSMailBox -Force

Describe "Get-MailItem Unit" -Tag "Unit" {
}

Describe "Get-MailItem Integration" -Tag "Integration" {
    It "Should get the first item in the inbox by default"{
        Get-MailItem | Should not BeNullOrEmpty
    }
}