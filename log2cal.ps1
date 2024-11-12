param (
    [switch]$filterToday
)

# Outlook Kalender hinzuf�gen
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendarFolder = $namespace.GetDefaultFolder(9)  # 9 entspricht dem Kalender

# Einlesen der Logfile
$logFile = "C:\Users\mcapellari\AppData\Roaming\TeamViewer\Connections2.txt"
$entries = Get-Content -Path $logFile | Where-Object { $_ -match '\d' }

# Das heutige Datum als Referenz
$currentDate = Get-Date

foreach ($line in $entries) {
    # Trennen der Felder nach beliebigen Leerzeichen
    $fields = $line -split '\s+'  # Splittet nach einem oder mehr Leerzeichen

    # Sicherstellen, dass wir genug Felder haben
    if ($fields.Length -ge 8) {
        $teamViewerId = $fields[0].Trim()
        $startDate = $fields[1].Trim()
        $startTime = $fields[2].Trim()
        $endDate = $fields[3].Trim()
        $endTime = $fields[4].Trim()
        $userId = $fields[5].Trim()
        $activity = $fields[6].Trim()
        $sessionId = $fields[7].Trim()

        # Wenn der Benutzername mit %username% �bereinstimmt
        if ($userId -eq $env:USERNAME) {
            try {
                # Umwandlung der Datums- und Uhrzeitangaben
                $startDateTime = [datetime]::ParseExact("$startDate $startTime", "dd-MM-yyyy HH:mm:ss", $null)
                $endDateTime = [datetime]::ParseExact("$endDate $endTime", "dd-MM-yyyy HH:mm:ss", $null)

                # Nur Eintr�ge des aktuellen Tages verarbeiten, wenn der Schalter gesetzt ist
                if ($filterToday -and $startDateTime.Date -eq $currentDate.Date) {
                    # Pr�fen, ob der Eintrag bereits hinzugef�gt wurde
                    $existingAppointment = $calendarFolder.Items | Where-Object {
                        $_.Body -match "TeamViewer ID: $teamViewerId" -and 
                        $_.Body -match "Session-ID: $sessionId" -and 
                        $_.Body -match "Startzeitpunkt: $startDateTime" -and 
                        $_.Body -match "Endzeitpunkt: $endDateTime"
                    }

                    if ($existingAppointment.Count -eq 0) {
                        # Berechnung der Sitzungsdauer
                        $duration = $endDateTime - $startDateTime
                        $durationStr = "{0} Stunden, {1} Minuten, {2} Sekunden" -f $duration.Hours, $duration.Minutes, $duration.Seconds

                        # Titel des Kalendereintrags: "TeamViewer Session: <TeamViewer ID>"
                        $appointment = $calendarFolder.Items.Add("IPM.Appointment")
                        $appointment.Subject = "TeamViewer Session: $teamViewerId"
                        $appointment.Start = $startDateTime
                        $appointment.End = $endDateTime
                        $appointment.Body = @"
TeamViewer ID: $teamViewerId
Startzeitpunkt: $startDateTime
Endzeitpunkt: $endDateTime
Benutzer-ID: $userId
T�tigkeit: $activity
Session-ID: $sessionId
Sitzungsdauer: $durationStr
"@
                        $appointment.Save()

                        Write-Host "Kalendereintrag erstellt f�r TeamViewer ID: $teamViewerId"
                    } else {
                        Write-Host "Kalendereintrag bereits vorhanden f�r TeamViewer ID: $teamViewerId ($sessionId) mit Startzeit $startDateTime und Endzeit $endDateTime"
                    }
                } elseif (-not $filterToday) {
                    # Wenn der Schalter nicht gesetzt ist, alle Eintr�ge verarbeiten
                    # Pr�fen, ob der Eintrag bereits hinzugef�gt wurde
                    $existingAppointment = $calendarFolder.Items | Where-Object {
                        $_.Body -match "TeamViewer ID: $teamViewerId" -and 
                        $_.Body -match "Session-ID: $sessionId" -and 
                        $_.Body -match "Startzeitpunkt: $startDateTime" -and 
                        $_.Body -match "Endzeitpunkt: $endDateTime"
                    }

                    if ($existingAppointment.Count -eq 0) {
                        # Berechnung der Sitzungsdauer
                        $duration = $endDateTime - $startDateTime
                        $durationStr = "{0} Stunden, {1} Minuten, {2} Sekunden" -f $duration.Hours, $duration.Minutes, $duration.Seconds

                        # Titel des Kalendereintrags: "TeamViewer Session: <TeamViewer ID>"
                        $appointment = $calendarFolder.Items.Add("IPM.Appointment")
                        $appointment.Subject = "TeamViewer Session: $teamViewerId"
                        $appointment.Start = $startDateTime
                        $appointment.End = $endDateTime
                        $appointment.Body = @"
TeamViewer ID: $teamViewerId
Startzeitpunkt: $startDateTime
Endzeitpunkt: $endDateTime
Benutzer-ID: $userId
T�tigkeit: $activity
Session-ID: $sessionId
Sitzungsdauer: $durationStr
"@
                        $appointment.Save()

                        Write-Host "Kalendereintrag erstellt f�r TeamViewer ID: $teamViewerId"
                    } else {
                        Write-Host "Kalendereintrag bereits vorhanden f�r TeamViewer ID: $teamViewerId ($sessionId) mit Startzeit $startDateTime und Endzeit $endDateTime"
                    }
                }
            } catch {
                Write-Host "Fehler beim Verarbeiten der Zeile: $line"
            }
        }
    } else {
        Write-Host "Ung�ltige Zeile �bersprungen: $line"
    }
}

Write-Host "Fertig!"
