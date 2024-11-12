# tv_log2outlook
Ein kleines Powershell Script um das Teamviewer Connections Log welches unter %appdata%\Teamviewer} Connections.txt liegt für Zeiterfassungszwecke in einen Outlook Kalender zu schreiben

Haftungsausschluss für log2cal.ps1

Dieses PowerShell-Skript, log2cal.ps1, wird kostenlos und „wie besehen“ zur Verfügung gestellt. Der Autor übernimmt keinerlei Garantie, weder ausdrücklich noch stillschweigend, bezüglich der Funktionalität, Genauigkeit oder Eignung dieses Skripts für einen bestimmten Zweck.

Haftungsbeschränkung:
Der Autor haftet nicht für direkte, indirekte, zufällige oder Folgeschäden, die aus der Nutzung, Installation, Änderung oder Weitergabe dieses Skripts entstehen. Der Anwender verwendet dieses Skript auf eigene Verantwortung und trägt das Risiko möglicher Schäden an Daten, Systemen oder Hardware.

Das Script ist relativ langsam, daher bietet es den Command Line Switch -filtertoday an, womit ausschließlich die Log-Einträge des aktuellen Tages in den Outlook Kalender übertragen werden, zum testen des Scripts sollte dieser Switch unbedingt verwendet werden, da sonst möglicherweise haufenweise Kalendereinträge erstellt werden.
