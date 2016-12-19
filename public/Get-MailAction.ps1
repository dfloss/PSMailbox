function Get-MailAction {
    param(
        [Parameter(Mandatory)]
        [ValidateScript({$_.Ast.ParamBlock.Parameters.Count -eq 1})]
            [ScriptBlock]$ScriptBlock
    )

    $MailAction =  {
        param($ScriptBlock)
        Try{
            foreach($notEvent in [array]$event.SourceEventArgs.Events){      
                Write-Host "Actually trying"
                $itmId = $notEvent.ItemId  
                $message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($Service,$itmId)
                & $ScriptBlock $message
            }
        }
        Catch{
            Write-Error $_
        }
        Write-Host "Mail recieved: $($event.SourceEventArgs.Events.Gettype() | Out-String)"
    }
    $SB = {& $MailAction $ScriptBlock}
    Return $MailAction
}