# log2out
Ein kleines Powershell Script um das Teamviewer Connections Log welches unter %appdata%\Teamviewer} Connections.txt liegt für Zeiterfassungszwecke in einen Outlook Kalender zu schreiben

Haftungsausschluss für log2cal.ps1

Haftungsbeschränkung:
Der Autor haftet nicht für direkte, indirekte, zufällige oder Folgeschäden, die aus der Nutzung, Installation, Änderung oder Weitergabe dieses Skripts entstehen. Der Anwender verwendet dieses Skript auf eigene Verantwortung und trägt das Risiko möglicher Schäden an Daten, Systemen oder Hardware.

Dieses PowerShell-Skript, log2cal.ps1, wird kostenlos und „wie besehen“ zur Verfügung gestellt. Der Autor übernimmt keinerlei Garantie, weder ausdrücklich noch stillschweigend, bezüglich der Funktionalität, Genauigkeit oder Eignung dieses Skripts für einen bestimmten Zweck.

Verwendung: es muss die Variable $logFile in Zeile 11 des Scripts nach eigenen Bedürfnissen angepasst werden, anschließend kann das Script aufgerufen werden, standardmässig geht das Script alle! Einträge des Logfiles durch, das kann etwas Zeit in Anspruch nehmen, sollen lediglich die Einträge des aktuellen Tages verarbeitet ist es möglich mittels Command Line Switch -filtertoday lediglich die Logfile Einträge des aktuellen Tages zu verarbeiten - d.h. Aufruf .\log2cal.ps1 -filtertoday

Zum Testen des Scripts empfiehlt es sich ebenfalls den Command Line Switch -filtertoday zu verwenden, da sonst möglicherweise haufenweise Kalendereinträge erstellt werden.

Aktuell hat das Script Probleme wenn der Username im Logfile ein Leerzeichen enthält, da das in meinem Anwendungsfall keine Rolle spielt, werde ich es vorläufig dabei belassen.
