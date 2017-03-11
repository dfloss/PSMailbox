function Set-DefaultMailbox {
    param(
        [Parameter()]$MailboxFile = "C:\users\$env:UserName\PSMAILBOX.xml",
        [Parameter()][Switch]$Force
    )
    #Exit if Mailbox is already setup and force isn't specificed, prevents extra load times
    If ($Script:DefaultService -ne $null -and -not $Force){Return}

    #setup creds, thanks office 365
    If ((Test-Path -Path $MailboxFile)){
        Write-Verbose "Loading File $Mailboxfile"
        $ServiceInfo = Import-Clixml $MailboxFile
        $Service = Get-ExchangeService -emailaddress $ServiceInfo.Credential.Username -Credential $ServiceInfo.Credential.GetNetworkCredential() -Url $ServiceInfo.url
    }
    Else{
        $Credential = Get-Credential -Message "Please Input the Email address and password for the account you'd like to setup as you default PSMAILBOX account"
        $Service = Get-ExchangeService -emailaddress $Credential.UserName -Credential $Credential.GetNetworkCredential()
        $ServiceInfo = [PSCustomObject]@{
            url = $Service.Url
            credential = $Credential
        }
        Export-Clixml -Path $MailboxFile -InputObject $ServiceInfo 
    }

    $script:DefaultService = $Service
    Write-Verbose "Service Set"
    $script:DefaultInbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Script:DefaultService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
}
