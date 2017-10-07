function Get-MailAction {
    param(
        [Parameter(Mandatory)]
        [ValidateScript({$_.Ast.ParamBlock.Parameters.Count -eq 1})]
        [Alias("MailController")]
            [ScriptBlock]$ScriptBlock
    )
    $MailAction =  {
        "MailAction Start" | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Force
        $Service = $event.MessageData.Service
        $ScriptBlock = $event.MessageData.ScriptBlock
        Try{
            "Try Block" | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Append
            foreach($notEvent in [array]$event.SourceEventArgs.Events){   
                "Foreach" | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Append
                $notEvent | FL | Out-String | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Append
                $itemId = $notEvent.ItemId
                $PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
                "Service: $Service" | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Append
                "itemid: $ItemId" | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Append
                "PropertySet: $PropertySet"
                $message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($Service,$itemId,$PropertySet)
                "Invoking Command" | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Append
                Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $message -ErrorAction Stop
            }
        }
        Catch{
            "Mail lisener failed with the following error: $_" | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Force -Append
            $_ | Out-String | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Force -Append
            Throw $_
        }
    }.GetNewClosure()

    Return $MailAction
}