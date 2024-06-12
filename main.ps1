# Pfade definieren
$exportPath = Join-Path -Path $PSScriptRoot -ChildPath "export"
$sourcePath = Join-Path -Path $PSScriptRoot -ChildPath "source"
$templatePath = Join-Path -Path $PSScriptRoot -ChildPath "template"
$bckPath = Join-Path -Path $PSScriptRoot -ChildPath "bck"

# Überprüfen und Erstellen von Verzeichnissen
$directories = @($exportPath, $sourcePath, $templatePath, $bckPath)
foreach ($dir in $directories) {
    if (-not (Test-Path -Path $dir)) {
        New-Item -ItemType Directory -Path $dir
    }
}

# Backup erstellen
$timestamp = (Get-Date).ToString("yyyy-MM-dd HH-mm-ss")
$backupName = "Backup_$timestamp.zip"
$backupPath = Join-Path -Path $bckPath -ChildPath $backupName
Compress-Archive -Path $sourcePath\* -DestinationPath $backupPath

# Template Datei überprüfen
$templateFiles = Get-ChildItem -Path $templatePath -Filter *.docx

if ($templateFiles.Count -eq 0) {
    Write-Error "Keine Dateien im Template Ordner gefunden."
    exit
} elseif ($templateFiles.Count -gt 1) {
    Write-Warning "Es gibt mehrere Dateien im Template Ordner. Die erste Datei wird verwendet."
}

$templateFile = $templateFiles[0]

# Sicherstellen, dass die Datei freigegeben wird, bevor wir sie kopieren
Start-Sleep -Seconds 2

$templateZipPath = Join-Path -Path $env:TEMP -ChildPath "template.zip"
Copy-Item -Path $templateFile.FullName -Destination $templateZipPath -Force

# Source Dateien überprüfen
$sourceFiles = Get-ChildItem -Path $sourcePath -Filter *.docx

if ($sourceFiles.Count -eq 0) {
    Write-Error "Keine Word Dateien im Source Ordner gefunden."
    exit
}

# Ignoriere nicht-Word Dateien im source Verzeichnis
$nonWordFiles = Get-ChildItem -Path $sourcePath | Where-Object { $_.Extension -ne ".docx" }
if ($nonWordFiles.Count -gt 0) {
    Write-Warning "Es gibt nicht-Word Dateien im Source Verzeichnis. Diese werden ignoriert."
}

# Hilfsfunktion zum Kopieren von Dateien in eine ZIP-Datei und Bearbeiten von [Content_Types].xml
function Copy-ZipContent {
    param (
        [string]$sourceZip,
        [string]$targetZip,
        [string[]]$filesToCopy
    )
    $sourceTempPath = Join-Path -Path $env:TEMP -ChildPath "sourceZip"
    $targetTempPath = Join-Path -Path $env:TEMP -ChildPath "targetZip"
    
    Remove-Item -Recurse -Force -Path $sourceTempPath, $targetTempPath -ErrorAction SilentlyContinue
    Expand-Archive -Path $sourceZip -DestinationPath $sourceTempPath
    Expand-Archive -Path $targetZip -DestinationPath $targetTempPath

    foreach ($file in $filesToCopy) {
        $sourceFilePath = Join-Path -Path $sourceTempPath -ChildPath $file
        $targetFilePath = Join-Path -Path $targetTempPath -ChildPath $file
        if (Test-Path -Path $sourceFilePath) {
            Write-Output "Kopiere $sourceFilePath nach $targetFilePath"
            Copy-Item -Path $sourceFilePath -Destination $targetFilePath -Force
        } else {
            Write-Output "Datei $sourceFilePath existiert nicht und wird daher nicht kopiert."
        }
    }

    # Bearbeiten oder Erstellen der Datei [Content_Types].xml im Root-Verzeichnis
    $contentTypesFile = Join-Path -Path $targetTempPath -ChildPath "[Content_Types].xml"
    if (-not (Test-Path -Path $contentTypesFile)) {
        Write-Output "Die Datei [Content_Types].xml wurde nicht gefunden. Erstelle eine neue Datei."
        $xmlContent = @"
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
    <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
    <Default Extension='xml' ContentType='application/xml'/>
</Types>
"@
        if (-not (Test-Path -Path $targetTempPath)) {
            New-Item -ItemType Directory -Path $targetTempPath -Force
        }
        $xmlContent | Out-File -FilePath $contentTypesFile -Encoding UTF8 -Force
    }

    Write-Output "Bearbeite Datei $contentTypesFile"
    [xml]$xml = Get-Content -Path $contentTypesFile
    $mediaTypes = @("jpeg", "png", "gif")  # Beispiel-Medientypen
    foreach ($mediaType in $mediaTypes) {
        $existing = $xml.Types.Default | Where-Object { $_.Extension -eq $mediaType }
        if (-not $existing) {
            Write-Output "Füge Tag für $mediaType zu $contentTypesFile hinzu"
            $newElement = $xml.CreateElement("Default", $xml.Types.NamespaceURI)
            $newElement.SetAttribute("Extension", $mediaType)
            $newElement.SetAttribute("ContentType", "image/$mediaType")
            $xml.Types.AppendChild($newElement)
        }
    }
    $xml.Save($contentTypesFile)

    $sourceMediaPath = Join-Path -Path $sourceTempPath -ChildPath "word\media"
    $targetMediaPath = Join-Path -Path $targetTempPath -ChildPath "word\media"

    if (Test-Path -Path $targetMediaPath) {
        Copy-Item -Path $sourceMediaPath\* -Destination $targetMediaPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    Compress-Archive -Path $targetTempPath\* -DestinationPath $targetZip -Update
}

# Bearbeitung der Word-Dateien
$filesToReplace = @(
    "word\_rels\header1.xml.rels",
    "word\_rels\footer1.xml.rels",
    "word\footer1.xml",
    "word\header1.xml",
    "[Content_Types].xml"
)

foreach ($sourceFile in $sourceFiles) {
    $sourceZipPath = Join-Path -Path $env:TEMP -ChildPath "$($sourceFile.BaseName).zip"
    $tempSourceDir = Join-Path -Path $env:TEMP -ChildPath "$($sourceFile.BaseName)"

    Remove-Item -Recurse -Force -Path $tempSourceDir -ErrorAction SilentlyContinue

    # Benenne die .docx Datei in .zip um
    $tempDocxPath = Join-Path -Path $env:TEMP -ChildPath "$($sourceFile.BaseName).docx"
    Copy-Item -Path $sourceFile.FullName -Destination $tempDocxPath -Force
    Rename-Item -Path $tempDocxPath -NewName "$($sourceFile.BaseName).zip"
    
    # Kopiere und ersetze Inhalte
    Copy-ZipContent -sourceZip $templateZipPath -targetZip $sourceZipPath -filesToCopy $filesToReplace

    # Erstelle die bearbeitete Datei
    Expand-Archive -Path $sourceZipPath -DestinationPath $tempSourceDir -Force

    $exportFilePath = Join-Path -Path $exportPath -ChildPath $sourceFile.Name

    # Verwenden Sie den Parameter -Force, um die vorhandene Archivdatei zu überschreiben
    Compress-Archive -Path "$tempSourceDir\*" -DestinationPath "$exportFilePath.zip" -Force

    # Benenne die ZIP-Datei wieder in .docx um
    if (Test-Path -Path "$exportFilePath.zip") {
        Rename-Item -Path "$exportFilePath.zip" -NewName $exportFilePath -Force
    }

    # Entferne nur existierende Dateien/Verzeichnisse
    if (Test-Path -Path $sourceZipPath) { Remove-Item -Recurse -Force -Path $sourceZipPath }
    if (Test-Path -Path $tempSourceDir) { 
        # Lassen Sie das tempSourceDir für Debugging-Zwecke bestehen
        # Remove-Item -Recurse -Force -Path $tempSourceDir 
        Write-Output "Temporäres Verzeichnis für Debugging-Zwecke beibehalten: $tempSourceDir"
    }
    if (Test-Path -Path $tempDocxPath) { Remove-Item -Recurse -Force -Path $tempDocxPath }
}

# Entferne temporäre Template ZIP-Datei
if (Test-Path -Path $templateZipPath) { 
    # Lassen Sie die templateZipPath für Debugging-Zwecke bestehen
    # Remove-Item -Recurse -Force -Path $templateZipPath 
    Write-Output "Temporäre Template-ZIP-Datei für Debugging-Zwecke beibehalten: $templateZipPath"
}

Write-Output "Skript erfolgreich abgeschlossen."
