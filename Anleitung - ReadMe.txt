#Einstellung:
Die Automatisierte Übertragung der Noten erfolgt über ein PowerShell Programm ("Noten Eintragung.ps1"), dies kann auf jedem Windows System ausgeführt werden, sofern Excel installiert ist.
Das Programm gleicht die Matrikelnummern in beiden Tabellen ab und überträgt die Noten wo die Matrikelnummern übereinstimmen.

Damit dieses Programm richtig funktioniert müssen in der config.csv Datei die richtigen Werte eingetragen werden:
"Dateipfad_Quelle" - der komplette Dateipfad zur Excel Datei in der Noten und Matrikelnummern eingetragen sind. 
"Worksheet_Quelle" - In einer Excel Datei kann es mehrere Arbeitsblätter geben. Hier muss der Name des richtigen Arbeitsblatt eingetragen werden. Die Schreibweise muss exakt stimmen.
"Spalte_Matrikelnummer_Quelle" - Die Spalte in der die Matrikelnummer eingetragen ist. (z.B. A)
"Spalte_Note_Quelle" - Die Spalte in der die Noten eingetragen sind. (z.B. B)
"Dateipfad_Ziel" - Der komplette Dateipfad (z.B. von C:\ ausgehend) zur Excel Datei in der die Noten eingetragen werden sollen.
"Worksheet_Ziel" - In einer Excel Datei kann es mehrere Arbeitsblätter geben. Hier muss der Name des richtigen Arbeitsblatt eingetragen werden. Die Schreibweise muss exakt stimmen.
"Spalte_Matrikelnummer_Ziel" - Die Spalte in der die Matrikelnummer eingetragen ist. (z.B. A)
"Spalte_Note_Ziel" - Die Spalte in der die Noten eingetragen werden sollen. (z.B. D)

Die config.csv am besten mit dem Texteditor von Windows öffnen:
Die Konfigurationsdatei muss in dem gleichen Verzeichnis liegen in welchem auch die Programm-Datei ("Noten Eintragung.ps1") zu finden ist.
Die Zieldatei darf nicht schreibgeschützt sein.

#Sicherheit:
Das Programm öffnet, speichert und schließt die Dateien selbstständig und im Hintergrund. Ein Datenverlust sollte nicht auftreten. 
Um dennoch einen Datenverlust auszuschließen bitte vorher eine Kopie von beiden Dateien anlegen und separat speichern.

Nach Ausführung des Programms sollte die Richtigkeit der Übertragung anhand von einigen Stichproben kontrolliert werden.

#Start:
Auf die Programm-Datei mit Rechtsklick gehen und auf die Option "mit PowerShell starten" klicken.


DISCLAIMER:
Dieses Programm soll helfen die Übertragung der Noten zu vereinfachen. Trotz ausführlicher Tests kann die Richtigkeit des Ergebnis nicht garantiert werden. 
Die Verantwortung dafür die Richtigkeit des Ergebnis zu Kontrollieren liegt beim Nutzer.
Wenn Fehler auffallen bitte Hinweise per E-Mail an maik.sube@gmail.com 