function Set-MailAction{
    param(
        [Parameter(Mandatory)]
        [ValidateScript({$_.Ast.ParamBlock.Parameters.Count -eq 1})]
            [ScriptBlock]$ScriptBlock,
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:DefaultService
    )

    $StreamingService = Get-MailSubscription -Service $Service
    $MailAction =  {
        $Service = $event.MessageData.Service
        $ScriptBlock = $event.MessageData.ScriptBlock
        Try{
            foreach($SubEvent in [array]$event.SourceEventArgs.Events){   
                $itemId = $SubEvent.ItemId
                $PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
                $message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($Service,$itemId,$PropertySet)
                Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $message -ErrorAction Stop
            }
        }
        Catch{
            "Mail lisener failed with the following error: $_" | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Force -Append
            $_ | Out-String | Out-File -FilePath 'C:\Users\dgharri\Desktop\MailTest.txt' -Force -Append
            Throw $_
        }
    }   
    $ConnectionRefreshSB = {
        Try{
            $event.MessageData.Open()
            write-host "[$(Get-Date)]Refreshing Mail Listener"
        }
        Catch{
            #Nested TryCatch because "MethodInvocationException" is not enough to go off of =(
            Try{
                $replacementsub = Get-MailSubscription -Service $Service -subonly
                $stmConnection.AddSubscription($replacementsub)
                $event.MessageData.Open()
            }
            Catch{
                Write-Error "Unable to refresh Mail Listener"
                Throw $_
            }
        }
    }
    $MessageData = [PSCustomObject]@{
        Service = $Service
        ScriptBlock = $ScriptBlock
    }

    Register-ObjectEvent -inputObject $StreamingService -eventName "OnNotificationEvent" -Action $MailAction -MessageData $MessageData
    Register-ObjectEvent -inputObject $StreamingService -eventName "OnDisconnect" -Action $ConnectionRefreshSB -MessageData $StreamingService
    $StreamingService.open()
}