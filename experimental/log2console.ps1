# Datei einlesen
$filePath = ".\Connections2.txt"
$entries = Get-Content $filePath

# Verarbeiten jeder Zeile
foreach ($line in $entries) {
    # Überspringen von leeren Zeilen
    if (-not $line.Trim()) {
        continue
    }

    # Entfernen überflüssiger Leerzeichen und Aufteilen in Felder
    $columns = $line -replace '\s+', ' ' -split '\s+', 8

    # Zuweisung der Spalten
    $teamViewerID = $columns[0].Trim()
    $startDate = $columns[1].Trim()
    $startTime = $columns[2].Trim()
    $endDate = $columns[3].Trim()
    $endTime = $columns[4].Trim()
    $userID = $columns[5].Trim()
    $activity = $columns[6].Trim()
    $sessionID = $columns[7].Trim()

    # Kombinieren von Datum und Zeit für Start- und Endzeitpunkte
    try {
        $startDateTime = [datetime]::ParseExact("$startDate $startTime", "dd-MM-yyyy HH:mm:ss", $null)
    } catch {
        Write-Output "Fehler beim Verarbeiten des Startzeitpunkts für TeamViewer ID: $teamViewerID"
        continue
    }

    try {
        $endDateTime = [datetime]::ParseExact("$endDate $endTime", "dd-MM-yyyy HH:mm:ss", $null)
    } catch {
        Write-Output "Fehler beim Verarbeiten des Endzeitpunkts für TeamViewer ID: $teamViewerID"
        $endDateTime = $null
    }

    # Berechnung der Sitzungsdauer, wenn beide Zeitpunkte vorhanden sind
    if ($endDateTime -ne $null) {
        $sessionDuration = $endDateTime - $startDateTime
        $durationString = "$($sessionDuration.Hours) Stunden, $($sessionDuration.Minutes) Minuten, $($sessionDuration.Seconds) Sekunden"
    } else {
        $durationString = "Endzeitpunkt fehlt oder ungültig"
    }

    # Ausgabe der Details
    Write-Output "TeamViewer ID: $teamViewerID"
    Write-Output "Startzeitpunkt: $startDateTime"
    Write-Output "Endzeitpunkt: $endDateTime"
    Write-Output "Benutzer-ID: $userID"
    Write-Output "Tätigkeit: $activity"
    Write-Output "Session-ID: $sessionID"
    Write-Output "Sitzungsdauer: $durationString"
    Write-Output "-----------------------------"
}
