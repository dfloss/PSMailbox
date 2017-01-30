$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
ipmo "$PSScriptRoot\..\PSMAILBOX.psd1"
Set-DefaultMailbox

Describe "Send-CalendarInvite Unit" -Tag "Unit" {
    It "Has no way to unit test with .NET objects"{}
}
Describe "Send-CalendarInvite Integration" -Tag "Integration"{
    Send-CalendarInvite -Subject "Test Event" -Body "Test Body" -Location "Test Location" -Start (Get-Date) -End ((Get-Date).AddHours(2))
}