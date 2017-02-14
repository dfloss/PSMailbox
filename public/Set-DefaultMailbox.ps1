function Set-DefaultMailbox {
    param(
        [Parameter()]$CredFile = "C:\users\$env:UserName\PSMAILBOX.txt"
    )
    If ($Script:DefaultService -ne $null){Return}
    $Script:DefaultEmailAddress = "$($env:USERNAME)@$($env:USERDNSDOMAIN)"
    $DefaultUri = 'https://outlook.office365.com/EWS/Exchange.asmx'

    #setup creds, thanks office 365
    Try{
        $DefaultCreds = Import-Clixml -Path $CredFile
    }
    Catch {
        $DefaultCreds = Get-Credential -UserName $DefaultEmailAddress -Message "Enter your email password to setup the default credentials"
        if ((Read-Host "Would you like to Store your credentials(y/N)") -like "y"){
            $DefaultCreds | Export-CliXml -Depth 6 -Path $CredFile
        }
    }
    $script:DefaultService = Get-ExchangeService -emailaddress $DefaultEmailAddress -Credential $DefaultCreds.GetNetworkCredential() -Uri $DefaultUrl
    $script:DefaultInbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($DefaultService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
}
