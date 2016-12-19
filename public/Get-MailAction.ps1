function Get-MailAction {
    param(
        [Parameter(Mandatory)]
        [ValidateScript({$_.Ast.ParamBlock.Parameters.Count -eq 1})]
            [ScriptBlock]$ScriptBlock
    )

    $MailAction =  {
        param($ScriptBlock)
        Try{
            foreach($notEvent in $event.SourceEventArgs.Events){      
                $itmId = $notEvent.ItemId  
                $message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($Service,$itmId)
                & $ScriptBlock $message
                Write-Host "Actually trying"
            }
        }
        Catch{
            Write-Error $_
        }
        Write-Host "Mail recieved: $($event.SourceEventArgs.Events | Out-String)"
    }
    $SB = {& $MailAction $ScriptBlock}
    Return $MailAction
}