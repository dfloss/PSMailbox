function New-Mailbox{
    [cmdletbinding()]
    param(
        [Parameter()]
            [PSCredential]$Credential = (Get-Credential -Message "Enter the Credentials for the new Mailbox, please include the full email address"),
        [Parameter()]
            [String]$Name = "",
        [Parameter()]
            [switch]$Force,
        [Parameter()]
            $url,
        [Parameter(dontshow)]$MailboxLocation = $Script:DefaultMailboxLocation
    )
    if ($Name -eq ""){
        $Name = ($Credentials.UserName)
    }
    $MailboxFile = "$MailboxLocation\$Name.xml"
    $Params = @{
        EmailAddress = $Credential.UserName
        Credential = ($Credential.GetNetworkCredential())
    }
    if ($url){$Params.Add('url',$url)}
    $Service = Get-ExchangeService @Params
    
    $MailboxItem = [PSCustomObject]@{
        Credential = $Credential
        Url = $Service.Url
    }

    $Params = @{
        Path = $MailboxFile
        Force = $Force
        InputObject = $MailboxItem
    }
    Export-Clixml @Params
}