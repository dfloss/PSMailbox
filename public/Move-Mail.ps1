function Move-Mail {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory)]
            [Object]$Email,
        [Parameter(Mandatory)]
            [String]$FolderName,
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $DefaultService
    )
    $FolderID = Get-EWSFolderID -FolderName $FolderName -Service $Service
    $Email.Move($FolderID)
}
