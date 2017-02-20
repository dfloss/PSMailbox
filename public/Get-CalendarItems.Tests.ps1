$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Get-CalendarItems" {
    It "Successfully Calls Calendar Get Items" {
        {Get-CalendarItems} | Should Not Throw
    }
}
