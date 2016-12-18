function Send-CalendarInvite {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory)]
            [string]$Subject,
        [Parameter(Mandatory)]
            [string]$Body,
        [Parameter(Mandatory)]
            [string]$Location,
        [Parameter(Mandatory)]
            [DateTime]$Start,
        [Parameter(Mandatory)]
            [DateTime]$End,
        [Parameter()]
            [String[]]$Attendees = $null,
        [Parameter()]
            [int] $ReminderMinutes = 15,
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $DefaultService
    )
    $Appointment = New-Object -TypeName Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $Service  #[Microsoft.Exchange.WebServices.Data.Appointment]::new($service)

    $Appointment.Subject = $Subject
    $Appointment.Body = $Body
    $Appointment.Location = $Location
    $Appointment.Start = $Start
    $Appointment.End = $End
    $Appointment.ReminderMinutesBeforeStart = $ReminderMinutes

    if ($Attendees -eq $null){
      Return $appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
    }
    else{
        foreach($Attendee in $Attendees){
            $appointment.RequiredAttendees.Add($Attendee)
        }
        return $Appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
    }
}