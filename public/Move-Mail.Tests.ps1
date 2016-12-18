$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
Ipmo PSMailBox -Force
Set-DefaultMailbox
Describe "Move-Mail" {
    It "Moves An Email" {
        $Email = Get-MailItem
        Move-Mail -Email $Email -FolderName "UREC"
    }
}
