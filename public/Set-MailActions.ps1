function Set-MailAction{
    param(
        [ScriptBlock]$Filter,
        [Parameter(Mandatory)]
        [ValidateScript({$_.Ast.ParamBlock.Parameters.Count -eq 1})]
        [ScriptBlock]$MailAction
    )
}