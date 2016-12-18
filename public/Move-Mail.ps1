function Move-Mail {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory)]
            [Microsoft.Exchange.WebServices.Data.EmailMessage]$Email,
        [Parameter(Mandatory)]
            [String]$FolderName,
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $DefaultService
    )
    $FolderID = Get-EWSFolderID -FolderName $FolderName -Service $Service
    $Email.Move($FolderID)
}
