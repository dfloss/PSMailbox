﻿function Send-Email {
    Param(
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $DefaultService,
        [Parameter(Mandatory)]
            [String]$Body,
        [Parameter(Mandatory)]
            [String]$Subject,
        [Parameter(Mandatory)]
            [string[]]$To,
        [Parameter(Mandatory = $false)]
            [string[]]$Bcc = $null,
        [Parameter()]
            [switch]$Html,
        [Parameter(ParameterSetName="InlineImage")]
            [HashTable]$InlineImages
    )
    
    $Email = [Microsoft.Exchange.WebServices.Data.EmailMessage]::new($Service)
    $To | %{$Email.ToRecipients.Add($_) | Out-Null}
    $Email.Subject = $Subject
    $Bcc | %{$Email.BccRecipients.Add($_) | Out-Null}
    
    If ($Html){$BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML}
    Else {$BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text}
    $Email.Body = [Microsoft.Exchange.WebServices.Data.MessageBody]::new($BodyType,$Body)

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