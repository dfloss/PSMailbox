function Get-CalendarItems {
    param(
        [Parameter()]
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $DefaultService,
        [Parameter()]
            [DateTime]$StartDate,
        [Parameter()]
            [DateTime]$EndDate,
        [Parameter()]
            [int]$MaxAppointments = 100,
        [Parameter()]
            [switch]$LoadProperties
    )

    $Calendar = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($Service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
    $CalendarView = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($StartDate,$EndDate,$MaxAppointments)
    $CalendarView.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Id,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Categories)
    $Appointments = $Calendar.FindAppointments($CalendarView)
    If ($LoadProperties){
        #$BodyProperty = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Body,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject,[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Categories)
        $BodyProperty = [Microsoft.Exchange.WebServices.Data.PropertySet]::FirstClassProperties
        foreach ($Appointment in $Appointments){
            $Appointment.Load($BodyProperty)
        }
    }
    Return $Appointments
}