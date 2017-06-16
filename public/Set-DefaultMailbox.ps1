function Set-DefaultMailbox {
    param(
        [Parameter()]$MailboxFile = $Script:DefaultMailboxFile,
        [Parameter()][Switch]$Force,
        [Parameter()][Switch]$Save

    )
    #Exit if Mailbox is already setup and force isn't specificed, prevents extra load times
    If ($Script:DefaultService -ne $null -and -not $Force){Return}

    #setup creds, thanks office 365
    If ($MailboxFile -ne $null -and (Test-Path -Path $MailboxFile) -and -not $Force){
        Write-Verbose "Loading File $Mailboxfile"
        $ServiceInfo = Import-Clixml $MailboxFile
        $Service = Get-ExchangeService -emailaddress $ServiceInfo.Credential.Username -Credential $ServiceInfo.Credential.GetNetworkCredential() -Url $ServiceInfo.url
    }
    Else{
        $Credential = Get-Credential -Message "Please Input the Email address and password for the account you'd like to setup as you default PSMAILBOX account at $MailBoxFile" -UserName $Script:DefaultEmailAddress 
        $Service = Get-ExchangeService -emailaddress $Credential.UserName -Credential $Credential.GetNetworkCredential() -url $Script:DefaultUri
        $ServiceInfo = [PSCustomObject]@{
            url = $Service.Url
            credential = $Credential
        }
        if ($Save){
            Export-Clixml -Path $MailboxFile -InputObject $ServiceInfo
        }
    }

    $script:DefaultService = $Service
    Write-Verbose "Service Set"
    $script:DefaultInbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Script:DefaultService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
}
