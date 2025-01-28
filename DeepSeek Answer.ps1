# Required Module: Microsoft.Graph (Install if missing)
if (-not (Get-Module Microsoft.Graph -ListAvailable)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Calendars.Read", "User.Read.All", "Calendar.ReadBasic"

# Resource Mailbox Configuration
$resourceMailboxUPN = "resource@domain.com"  # Replace with your resource mailbox UPN

# Get Resource Mailbox Calendar
$resourceUser = Get-MgUser -UserId $resourceMailboxUPN
$calendar = Get-MgUserCalendar -UserId $resourceUser.Id -Filter "Name eq 'Calendar'"

# Get Users with Calendar Access
$permissions = Get-MgUserCalendarPermission -UserId $resourceUser.Id -CalendarId $calendar.Id
$usersWithAccess = @()
foreach ($perm in $permissions) {
    if ($perm.Role -match "read|owner|editor" -and $perm.EmailAddress.Address -ne "Default") {
        $user = Get-MgUser -Filter "mail eq '$($perm.EmailAddress.Address)'" -ErrorAction SilentlyContinue
        if ($user) { $usersWithAccess += $user }
    }
}

# Retrieve Calendar Events
$events = Get-MgUserCalendarEvent -UserId $resourceUser.Id -CalendarId $calendar.Id -All -Property "Subject,Start,End,Organizer"

# Prepare Event Data
$report = foreach ($event in $events) {
    $organizer = $event.Organizer.EmailAddress.Address
    $organizerUser = Get-MgUser -Filter "mail eq '$organizer'" -ErrorAction SilentlyContinue
    $bookerName = if ($organizerUser) { $organizerUser.DisplayName } else { $organizer }

    foreach ($user in $usersWithAccess) {
        [PSCustomObject]@{
            EmployeeName   = $user.DisplayName
            EmployeeEmail  = $user.Mail
            EventStart     = $event.Start.DateTime
            EventEnd       = $event.End.DateTime
            EventSubject   = $event.Subject
            BookerName     = $bookerName
            BookerEmail    = $organizer
        }
    }
}

# Export Results
$report | Export-Csv -Path "ResourceCalendarReport.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Report exported to ResourceCalendarReport.csv"

# Optional: Disconnect Graph session
Disconnect-MgGraph
