﻿function Get-MailItem {
    [cmdletbinding(DefaultParameterSetName="FolderObject")]
    param(
        [Parameter(Mandatory,ParameterSetName="FolderName")]
            [string]$FolderName,
        [Parameter(ParameterSetName="FolderObject")]
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder = $DefaultInbox,
        [Parameter()][ValidateRange(1,[int]::MaxValue)]
            [int]$ItemNumber = 0,
        [Parameter()][ValidateRange(1,[int]::MaxValue)]
            [int]$NumberOfItems = 1,
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $DefaultService
    )

    if ($PSCmdlet.ParameterSetName -eq "FolderName"){
        $Folder = Get-EwsFolder -FolderName $FolderName -Service $Service
    }
    $PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

    $items = $Folder.FindItems([Microsoft.Exchange.WebServices.Data.ItemView]::new($NumberOfItems,$ItemNumber))
    $items.Load($PropertySet)
    return $items
}