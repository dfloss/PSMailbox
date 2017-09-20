#Import the EWS dll
[void][Reflection.Assembly]::LoadFile("$PSSCriptRoot\lib\Microsoft.Exchange.WebServices.dll")
#dot source our config to import config variables into script scope
. "$PsScriptRoot\config.ps1"

#Setup the builtin Streaming subscription service
#$Script:SubscriptionService = new-object Microsoft.Exchange.WebServices.Data.StreamingSubscriptionConnection($service, 30)

#Get public and private function definition files while excluding tests
    $Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue | Where Name -NotMatch "\.Tests\.ps1$")
    $Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue | Where Name -NotMatch "\.Tests\.ps1$")

#Dot source the function files
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
Set-DefaultMailbox -Save