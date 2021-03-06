﻿function Get-MailSubscription {
    [cmdletbinding()]
    param(
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:DefaultService,
        [Parameter()]
            [switch]$subonly
    )
         $FolderIds = new-object Microsoft.Exchange.WebServices.Data.FolderId[] 1
         $Inboxid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$Service.Credentials.Credentials.UserName)
         $FolderIds[0] = $Inboxid

        $stmSubscription = $service.SubscribeToStreamingNotifications($FolderIds, [Microsoft.Exchange.WebServices.Data.EventType]::NewMail)  
        if ($subonly){
            Return $stmSubscription
        }
        else{
            $stmConnection = new-object Microsoft.Exchange.WebServices.Data.StreamingSubscriptionConnection($service, 30);
            $stmConnection.AddSubscription($stmSubscription)
            return $stmConnection
        }
}
