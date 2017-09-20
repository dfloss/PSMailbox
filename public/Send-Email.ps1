function Send-Email {
    Param(
        [Parameter(Mandatory)]
            [string[]]$To,
        [Parameter(Mandatory)]
            [String]$Subject,
        [Parameter(Mandatory)]
            [String]$Body,
        [Parameter(Mandatory = $false)]
            [string[]]$Bcc = $null,
        [Parameter()]
            [switch]$Html,
        [Parameter()]
            [ValidateScript({Test-Path $_})]
            [string[]]$Attachments,
        [Parameter(ParameterSetName="InlineImage")]
            [HashTable]$InlineImages,
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $DefaultService
    )
    
    $Email = [Microsoft.Exchange.WebServices.Data.EmailMessage]::new($Service)
    $To | %{$Email.ToRecipients.Add($_) | Out-Null}
    $Email.Subject = $Subject
    If ($Bcc -ne $null){$Bcc | %{$Email.BccRecipients.Add($_) | Out-Null}}
    
    If ($Html){$BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML}
    Else {$BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text}
    $Email.Body = [Microsoft.Exchange.WebServices.Data.MessageBody]::new($BodyType,$Body)

    #AttachMents
    Foreach($Attachment in $Attachments){
        $Email.Attachments.AddFileAttachment($Attachment) | Out-Null
    }

    #Inline Images
    If ($PSCmdlet.ParameterSetName -eq "InlineImage"){
        $i=0
        $InlineImages.GetEnumerator() | %{
            $Key = $_.Key
            $Email.Attachments.AddFileAttachment($Key,$_.Value)
            $Email.Attachments[$i].IsInline = $true
            $Email.Attachments[$i].ContentId = $Key
            Write-Verbose ($Email.Attachments | Out-String)
        }
    }
    Try{
        Return $Email.SendAndSaveCopy()
    }
    Catch {
        Throw $_
    }
}
