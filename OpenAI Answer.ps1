# Verbindung zu Exchange Online herstellen
Connect-ExchangeOnline -UserPrincipalName Admin@deinedomain.de

# Definiere die Ressourcen-Mailbox und die gewünschten Einstellungen
$resourceMailbox = "raum1@deinedomain.de"
$defaultAccessLevel = "Reviewer" # Benutzer können Termine und Details lesen
$resourceDisplayDetails = @("Organizer", "Subject") # Felder, die angezeigt werden sollen

# Zugriffsebene für die Ressourcen-Mailbox auf "Reviewer" setzen
Set-MailboxFolderPermission -Identity "$resourceMailbox:\Kalender" -User Default -AccessRights $defaultAccessLevel

# Konfiguration der Ressourcen-Mailbox anpassen, um Bucher und Betreff darzustellen
Set-CalendarProcessing -Identity $resourceMailbox `
    -AddOrganizerToSubject $true `
    -DeleteComments $false `
    -DeleteSubject $false

Write-Host "Die Ressourcen-Mailbox $resourceMailbox wurde erfolgreich konfiguriert."

# Sitzung trennen
Disconnect-ExchangeOnline -Confirm:$false
