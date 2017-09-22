[cmdletbinding()]
Param(
    [Parameter()]
        $Force
)
$Parameters = @{
    Path = "$env:USERPROFILE\PSMAILbox"
    ItemType = 'Directory'
    Force = $Force
}
New-Item @Parameters