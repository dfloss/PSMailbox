function Get-ExchangeService {
    [cmdletbinding()]
    param(
        [parameter(Mandatory)][string]$emailaddress,
        [parameter(Mandatory)]$Credential,
        [parameter()][string]$uri = $null
        
    )

    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService    
    $Service.Credentials = $Credential
    if ($uri -eq ""){
        $Service.AutodiscoverUrl($emailaddress,{$true})
    }
    else{
        $Service.Url = [uri]$Uri
    }
    
    Return $Service
}
