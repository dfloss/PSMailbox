$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
Ipmo $here/../PSMailBox.psd1 -Force
Set-DefaultMailbox
Describe "Get-EwsFolder Unit" -Tag "Unit" {
}

Describe "Get-EwsFolder Integration" -Tag "Integration" {
    Context "Well Known Folders"{
        It "Should return the inbox"{
            $Folder = Get-EWSFolder
        }
    }
}