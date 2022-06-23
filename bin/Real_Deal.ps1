$rootSrc = "C:\Users\Dominic\Documents\ExportInfoScript\"
$sourceSrc = ($rootSrc + "sourceFiles\")
$templateSrc = ($rootSrc + "templates\")
$exportSrc = ($rootSrc + "exportedFiles\")

Get-ChildItem -Path $sourceSrc -Recurse  | ForEach-Object {
    <# https://stackoverflow.com/questions/10727919/add-headers-and-footers-to-word-document-with-power-shell #>
    <# https://techblog.dorogin.com/generate-word-documents-with-powershell-cda654b9cb0e #>
    # Create a new Word application COM object
    $WordAPI = New-Object -ComObject Word.Application;
    # Make the Word application visible
    $WordAPI.Visible = $true;
    # Add a new document to the application
    $ExportedDoc = $WordAPI.Documents.Add($sourceSrc + $_);
    # Header and footer document
    $TemplateDoc = $WordAPI.Documents.Add($templateSrc + "BBWZ.docx");
    $TemplateSection = $TemplateDoc.Sections.Item(1);
    # Get the first Section of the Document object
    $ExportedSection = $ExportedDoc.Sections.Item(1);
    
    $TemplateSection.Headers.Item(1).Range.copy($ExportedSection.Headers.Item(1).selection.range)
    $ExportedSection.Headers.Item(1).range.PasteSpecial()

    $TemplateSection.footers.Item(1).Range.copy($ExportedSection.footers.Item(1).selection.range)
    $ExportedSection.footers.Item(1).range.PasteSpecial()

    $ExportedDoc.SaveAs($exportSrc + $_);
    $ExportedDoc.Close();

    $WordAPI.Quit();
}
