# Konfigurationsdatei einlesen, basierend auf dem Verzeichnis, in dem das Skript liegt
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
$configFilePath = Join-Path $scriptDirectory "config.csv"
$configData = Import-Csv -Path $configFilePath -Delimiter ','

# Konfigurationsparameter in Variablen speichern
$Dateipfad_Quelle = ($configData | Where-Object { $_.Parameter -eq "Dateipfad_Quelle" }).Wert
$Worksheet_Quelle = ($configData | Where-Object { $_.Parameter -eq "Worksheet_Quelle" }).Wert
$Spalte_Matrikelnummer_Quelle = ($configData | Where-Object { $_.Parameter -eq "Spalte_Matrikelnummer_Quelle" }).Wert
$Spalte_Note_Quelle = ($configData | Where-Object { $_.Parameter -eq "Spalte_Note_Quelle" }).Wert
$Dateipfad_Ziel = ($configData | Where-Object { $_.Parameter -eq "Dateipfad_Ziel" }).Wert
$Worksheet_Ziel = ($configData | Where-Object { $_.Parameter -eq "Worksheet_Ziel" }).Wert
$Spalte_Matrikelnummer_Ziel = ($configData | Where-Object { $_.Parameter -eq "Spalte_Matrikelnummer_Ziel" }).Wert
$Spalte_Note_Ziel = ($configData | Where-Object { $_.Parameter -eq "Spalte_Note_Ziel" }).Wert

# Excel-Anwendung starten
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Datei A (Quelle) öffnen
$workbookSource = $excel.Workbooks.Open($Dateipfad_Quelle)
$worksheetSource = $workbookSource.Worksheets.Item($Worksheet_Quelle)

# Datei B (Ziel) öffnen
$workbookTarget = $excel.Workbooks.Open($Dateipfad_Ziel)
$worksheetTarget = $workbookTarget.Worksheets.Item($Worksheet_Ziel)

# Funktion zum Ermitteln des letzten benutzten Zeilenindex
function Get-LastRow($worksheet, $column) {
    $row = 1
    while ($worksheet.Cells.Item($row, $column).Value2 -ne $null) {
        $row++
    }
    return $row - 1
}

# Funktion zur Überprüfung, ob eine Zelle eine Zahl enthält
function Is-Number($value) {
    return [double]::TryParse($value, [ref]0)
}

# Matrikelnummern und Noten aus Datei A (Quelle) auslesen (nur Zahlenwerte)
$lastRowSource = Get-LastRow $worksheetSource ([int][char]::Parse($Spalte_Matrikelnummer_Quelle) - 64)
$sourceData = @{}
for ($i = 2; $i -le $lastRowSource; $i++) {
    $matrikelnummer = $worksheetSource.Cells.Item($i, [int][char]::Parse($Spalte_Matrikelnummer_Quelle) - 64).Text

    # Nur weiterverarbeiten, wenn es sich um eine Zahl handelt
    if (Is-Number($matrikelnummer)) {
        $note = $worksheetSource.Cells.Item($i, [int][char]::Parse($Spalte_Note_Quelle) - 64).Value2
        $sourceData[$matrikelnummer] = $note
    }
}

# Matrikelnummern aus Datei B (Ziel) abgleichen und Noten einfügen (nur Zahlenwerte)
$lastRowTarget = Get-LastRow $worksheetTarget ([int][char]::Parse($Spalte_Matrikelnummer_Ziel) - 64)
for ($i = 2; $i -le $lastRowTarget; $i++) {
    $matrikelnummer = $worksheetTarget.Cells.Item($i, [int][char]::Parse($Spalte_Matrikelnummer_Ziel) - 64).Text

    # Überspringe leere Zeilen oder Überschriften
    if (Is-Number($matrikelnummer)) {
        # Wenn die Matrikelnummer in den Quelldaten vorhanden ist, trage die Note ein
        if ($sourceData.ContainsKey($matrikelnummer)) {
            $worksheetTarget.Cells.Item($i, [int][char]::Parse($Spalte_Note_Ziel) - 64).Value2 = $sourceData[$matrikelnummer]
        }
    }
}

# Änderungen speichern und Dateien schließen
$workbookTarget.Save()
$workbookSource.Close()
$workbookTarget.Close()

# Excel beenden
$excel.Quit()

# COM-Objekte freigeben
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheetSource) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookSource) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheetTarget) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookTarget) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Daten erfolgreich abgeglichen und gespeichert."
