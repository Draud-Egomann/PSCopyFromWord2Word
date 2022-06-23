<# 
Dieses Powershell Skript ersetzt den Header und Footer mit einem Template basierten Word
Authors: Dominic Tosku, Justin Urbanek, Kristyan Usarz, Luciano Zehnder
Alias: TgN
Version: 1.0
Letzte Änderung: 23.06.22 23:41
#>

<# Nutzerpfade #>
$rootSrc = "<Geben-Sie-Ihren-Dateipfad-an>"
$templateFilesName = "template.docx"
<# Src Variabeln #>
$sourceSrc = ($rootSrc + "\sourceFiles\")
$templateSrc = ($rootSrc + "\templates\")
$exportSrc = ($rootSrc + "\exportedFiles\")

#Geht die Word Dateien im Quel verzeichnis durch
Get-ChildItem -Path $sourceSrc -Recurse  | ForEach-Object {
    try {
        # Erstellt neue Word Instanz
        $WordAPI = New-Object -ComObject Word.Application;
        # Stellt ein, ob die Datei sichtbar ist während der bearbeitung
        $WordAPI.Visible = $true;
        # Fügt eine neue Word Datei hinzu
        $ExportedDoc = $WordAPI.Documents.Add($sourceSrc + $_);
        # Öffnet das Template Dokument
        $TemplateDoc = $WordAPI.Documents.Add($templateSrc + $templateFilesName);
        #Speichert die erste Sektion der Template Word Datei
        $TemplateSection = $TemplateDoc.Sections.Item(1);
        #Speichert die erste Sektion der Ziel Word Datei
        $ExportedSection = $ExportedDoc.Sections.Item(1);
        
        #Kopiert den Header des Templates und fügt ihn der Ziel Datei ein
        $TemplateSection.Headers.Item(1).Range.copy($ExportedSection.Headers.Item(1).selection.range)
        $ExportedSection.Headers.Item(1).range.PasteSpecial()
        
        #Kopiert den Footer des Templates und fügt ihn der Ziel Datei ein
        $TemplateSection.footers.Item(1).Range.copy($ExportedSection.footers.Item(1).selection.range)
        $ExportedSection.footers.Item(1).range.PasteSpecial()
        
        #Speichert die File unter dem selben Namen
        $ExportedDoc.SaveAs($exportSrc + $_);
        $ExportedDoc.Close();
        
        #Schliesst alle Instanzen
        $WordAPI.Quit();
    }
    catch {
        Write-Error "Das Dokument ist schreibgesch�tzt. Bitte �ffen Sie die Dokumente manuell und aktivieren Sie die Bearbeitung."
        #Schliesst alle Instanzen, wenn ein Fehler auftritt
        $ExportedDoc.Close();
        $WordAPI.Quit();
    }
}

Write-Host "Das Skript ist fertig." -ForegroundColor Black -BackgroundColor Green