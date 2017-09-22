function Get-ExchangeService {
    [OutputType([Microsoft.Exchange.WebServices.Data.ExchangeService])]
    [cmdletbinding()]
    param(
        [parameter(Mandatory)][string]$emailaddress,
        [parameter(Mandatory)]$Credential,
        [parameter()][uri]$url = $null
        
    )

    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService    
    $Service.Credentials = $Credential
    write-Verbose "Url: $url"
    Write-Verbose "Url -eq Null: $($url -eq $null)"
    if ($url -eq $null -or $url -eq ""){
        Write-Verbose "Auto Discovering"
        Write-Debug "ABOUT TO AUTODISCOVER GET DOWN"
        $Service.AutodiscoverUrl($emailaddress,{$true})
    }
    else{
        Write-Verbose "Setting Url"
        $Service.Url = $Url
    }
    Write-Verbose "returning Service"
    Return $Service
}
