$MailboxScriptBlock = {
    Get-ChildItem $Script:DefaultMailboxLocation | Select -ExcludeProperty Name |Foreach{
        Try{
        $Name = $_
        [System.Management.Automation.CompletionResult]::new("'$_'",$Name,'ParameterValue',$Name)
        }
        Catch{
            write-host $_
        }
    }
}
Register-ArgumentCompleter -CommandName 'Set-DefaultMailbox' -ParameterName 'Mailbox' -ScriptBlock $MailboxScriptBlock