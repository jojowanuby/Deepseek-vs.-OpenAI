# Verbinden Sie sich mit Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName youradmin@yourdomain.com -ShowProgress $true

# Definieren Sie die Ressourcenmailbox und den Suchzeitraum
$ResourceMailbox = "resourcemailbox@yourdomain.com"
$StartDate = (Get-Date).AddDays(-7)  # Beispiel: Letzte 7 Tage
$EndDate = Get-Date

# Holen Sie sich alle Termine der Ressourcenmailbox im definierten Zeitraum
$CalendarItems = Get-MailboxFolderStatistics -Identity $ResourceMailbox | Where-Object { $_.FolderType -eq "Calendar" }
$CalendarItems = Get-CalendarDiagnosticObjects -Identity $ResourceMailbox -MailboxFolderId $CalendarItems.FolderId -StartDate $StartDate -EndDate $EndDate

# Holen Sie sich alle Benutzer mit integriertem Kalender
$Users = Get-Mailbox -RecipientTypeDetails UserMailbox | Where-Object { $_.CalendarIntegrationEnabled -eq $true }

# Fügen Sie die Termine der Ressourcenmailbox in den Kalender der Benutzer ein
foreach ($User in $Users) {
    foreach ($Item in $CalendarItems) {
        $Subject = $Item.Subject
        $Organizer = $Item.Organizer
        $StartTime = $Item.Start
        $EndTime = $Item.End

        # Erstellen Sie einen neuen Kalendereintrag im Kalender des Benutzers
        New-MailboxFolder -Mailbox $User.Identity -Name "ResourceCalendar" -Parent "\Calendar"
        $UserCalendar = Get-MailboxFolderStatistics -Identity $User.Identity | Where-Object { $_.FolderType -eq "Calendar" }
        $Appointment = New-CalendarDiagnosticObject -Mailbox $User.Identity -MailboxFolderId $UserCalendar.FolderId -Subject $Subject -Start $StartTime -End $EndTime -Organizer $Organizer

        # Fügen Sie den Kalendereintrag hinzu
        Add-MailboxFolderPermission -Identity $UserCalendar.FolderId -User $User.Identity -AccessRights Owner
        Set-CalendarProcessing -Identity $User.Identity -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false
    }
}

# Trennen Sie die Verbindung zu Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
