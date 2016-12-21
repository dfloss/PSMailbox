function Get-MailAction {
    param(
        [Parameter(Mandatory)]
        [ValidateScript({$_.Ast.ParamBlock.Parameters.Count -eq 1})]
            [ScriptBlock]$MailController
    )

    $MailAction =  {
        param($MailController, $event)
        Try{
            foreach($notEvent in [array]$event.SourceEventArgs.Events){      
                $itmId = $notEvent.ItemId
                $PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
                $message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($Service,$itmId,$PropertySet)
                & $MailController $message
            }
        }
        Catch{
            Write-Error "Mail lisener failed with the following error: "
            Write-Error $_
        }
    }

    Return $MailAction
}