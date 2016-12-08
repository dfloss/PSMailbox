$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
Ipmo PSMailBox -Force

InModuleScope PSMailBox{
    Describe "Get-EwsFolderID Unit" -Tag "Unit" {
    }

    Describe "Get-EwsFolderID Integration" -Tag "Integration" {
        It "Should Find the Inbox"{
            Get-EWsFolderID -FolderName "Inbox"
        }
    }
}