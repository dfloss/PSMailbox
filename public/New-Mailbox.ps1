function New-Mailbox{
    [cmdletbinding()]
    param(
        [Parameter(Mandatory)]
            [PSCredential]$Credentials,
        [Parameter()]
            [String]$Name = "",
        [Parameter()]
            [switch]$Force,
        [Parameter(dontshow)]$MailboxLocation = $Script:DefaultMailboxLocation
    )
    if ($Name -eq ""){
        $Name = ($Credentials.UserName) -replace "@.+",""
    }

}