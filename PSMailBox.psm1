#Import the EWS dll
[void][Reflection.Assembly]::LoadFile("$PSSCriptRoot\lib\Microsoft.Exchange.WebServices.dll")

$DefaultEmailAddress = "$($env:USERNAME)@$($env:USERDNSDOMAIN)"
$DefaultUri = 'https://outlook.office365.com/EWS/Exchange.asmx'
<#
#setup creds, thanks office 365
Try{
    $passfile = "C:\users\$env:UserName\network.txt"
    $securepass = Get-Content -Path $passfile | ConvertTo-SecureString
    $DefaultEmailAddress = "$($env:USERNAME)@$($env:USERDNSDOMAIN)"
    $Defaultcreds = new-object -typename System.Management.Automation.PSCredential `
                    -argumentlist $DefaultEmailaddress, $securepass
}
Catch {
    $DefaultCreds = Get-Credential -UserName $DefaultEmailAddress -Message "Enter your email password to setup the default credentials"
}
#>

#Get public and private function definition files while excluding tests
    $Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue | Where Name -NotMatch "\.Tests\.ps1$")
    $Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue | Where Name -NotMatch "\.Tests\.ps1$")

#Dot source the files
    Foreach($import in @($Public + $Private))
    {
        Try
        {
            . $import.fullname
        }
        Catch
        {
            Write-Error -Message "Failed to import function $($import.fullname): $_"
        }
    }

#Export all public modules
Export-ModuleMember -Function $Public.Basename

#Create Default Service and inbox for EZ mode
#$DefaultService = Get-ExchangeService -emailaddress $DefaultEmailAddress -Credential $DefaultCreds.GetNetworkCredential() -Uri $DefaultUrl
#$DefaultInbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($DefaultService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)