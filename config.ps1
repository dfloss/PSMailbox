$Script:DefaultEmailAddress = "$($env:USERNAME)@$($env:USERDNSDOMAIN)"
$Script:DefaultUri = 'https://outlook.office365.com/EWS/Exchange.asmx'
$Script:DefaultMailboxLocation = "$env:USERPROFILE\PSMailBox"
$Script:DefaultMailboxFile = "$Script:DefaultMailboxLocation\$env:Username.xml"