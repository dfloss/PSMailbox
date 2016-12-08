$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
ipmo PSMailBox -Force

Describe "Get-ExchangeService Unit" -Tag "Unit" {
}

Describe "Get-ExchangeService Integration" -Tag "Integration" {
    InModuleScope PSMailBox{
        Context "Creates Default Exchange Service Successfully"{
            It "Should Return a connected Exchange Service"{
                $DefaultService | Should Not BeNullOrEmpty
            }
        }
    }
}