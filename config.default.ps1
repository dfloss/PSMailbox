$Script:DefaultEmailAddress = "$($env:USERNAME)@$($env:USERDNSDOMAIN)"
$Script:DefaultUri = 'https://outlook.office365.com/EWS/Exchange.asmx'
$Script:DefaultMailboxFile = "$env:USERPROFILE\PSMailBox\$env:Username.xml"
$Script:DefaultMailboxLocation = "$env:USERPROFILE\PSMailBox"