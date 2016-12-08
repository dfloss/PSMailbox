function Get-EWSFolderID {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory)][string]$FolderName,
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $DefaultService
    )

    $Account = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
    $Account.Load()
    $FolderView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderView -ArgumentList 100
    $FolderViewId = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
    $FolderViewName = [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName
    $FolderView.PropertySet = New-Object -TypeName Microsoft.Exchange.WebServices.Data.PropertySet -ArgumentList $FolderViewID,$FolderViewName
    $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    foreach($Folder in $Account.FindFolders($FolderView)){
        If($Folder.DisplayName -like $FolderName){
            Return ($Folder.Id)
        }
    }
    Throw [System.Management.Automation.ItemNotFoundException] "Unable to locate Folder with Name $FolderName"
}
