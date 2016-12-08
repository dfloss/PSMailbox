function Set-MailListener {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({$_.Ast.ParamBlock.Parameters.Count -eq 1})]
            [ScriptBlock]$MailAction,

<#        [Parameter(ParameterSetName="FolderIds")]
            [Microsoft.Exchange.WebServices.Data.FolderId[]]$FolderIds = @($DefaultInbox.id),
        [Parameter(ParameterSetName="FolderNames")]
            [Array]$FolderNames,
#>
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $DefaultService
    )
         $FolderIds = new-object Microsoft.Exchange.WebServices.Data.FolderId[] 1
         $Inboxid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$service.Credentials.Credentials.UserName)
         $FolderIds[0] = $Inboxid

    #create our script block for our mail action action, loops through each email from the event and executes the passed code block once per email
    $MailScriptBlock = {
        param($MailAction)
        Try{
            Add-Content -Value "Event Recieved" -Path "C:\users\dgharri\desktop\textemail.txt"
            $PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
            $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
            foreach($notEvent in $event.SourceEventArgs.Events){      
                $itmId = $notEvent.ItemId  
                $message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($Service,$itmId)
                $message.load($PropertySet)
                & $MailAction $message
            }
        }
        Catch{
            Add-Content -Value "$_" -Path "C:\users\dgharri\desktop\errors.txt"
            throw $_
        }
    }
    $Action = {& $MailScriptBlock -MailAction $MailAction}
    #Write-Verbose $FolderIds
    $stmSubscription = $service.SubscribeToStreamingNotifications($FolderIds, [Microsoft.Exchange.WebServices.Data.EventType]::NewMail)  
    $stmConnection = new-object Microsoft.Exchange.WebServices.Data.StreamingSubscriptionConnection($service, 30);  
    $stmConnection.AddSubscription($stmSubscription)  


    Register-ObjectEvent -inputObject $stmConnection `
                         -eventName "OnNotificationEvent" `
                         -MessageData $service `
                         -Action $Action 

    Register-ObjectEvent -inputObject $stmConnection -eventName "OnDisconnect" -Action {$event.MessageData.Open()} -MessageData $stmConnection  
    $stmConnection.Open()
}
