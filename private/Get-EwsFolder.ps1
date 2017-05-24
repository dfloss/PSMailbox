function Get-EwsFolder {
    [cmdletbinding()]
    param(
        [Parameter(ValueFromPipeline)][string]$FolderName = "Inbox",
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $ScriptDefaultService
    )
    Switch($FolderName){
        "Inbox"{
            $FolderID = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
        }
        "Sent"{
            $FolderID = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems
        }
        "Sent Items"{
            $FolderID = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems
        }
        "Outbox"{
            $FolderID = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Outbox
        }
        "Drafts"{
            $FolderID = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Drafts
        }
        ""{
            $FolderID = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail
        }
        "Deleted"{
            $FolderID = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems
        }
        "Deleted Items"{
            $FolderID = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems
        }
        ""{
            $FolderID = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]
        }
        Default{
            $FolderID = Get-EWSFolderID -FolderName $FolderName -Service $Service
        }
    }
    $PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$FolderId)
    Return $Folder
}
