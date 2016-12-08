function Set-DefaultMailbox {
    $Script:DefaultEmailAddress = "$($env:USERNAME)@$($env:USERDNSDOMAIN)"
    $DefaultUri = 'https://outlook.office365.com/EWS/Exchange.asmx'

    #setup creds, thanks office 365
    Try{
        $passfile = "C:\users\$env:UserName\network.txt"
        $securepass = Get-Content -Path $passfile | ConvertTo-SecureString
        $Defaultcreds = new-object -typename System.Management.Automation.PSCredential `
                        -argumentlist $DefaultEmailaddress, $securepass
    }
    Catch {
        $DefaultCreds = Get-Credential -UserName $DefaultEmailAddress -Message "Enter your email password to setup the default credentials"
    }
    $script:DefaultService = Get-ExchangeService -emailaddress $DefaultEmailAddress -Credential $DefaultCreds.GetNetworkCredential() -Uri $DefaultUrl
    $script:DefaultInbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($DefaultService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
}
