Add-type -AssemblyName Microsoft.Office.Interop.Word

$folder = "\\vmware-host\Shared Folders\Downloads\pdf-temp"
$files = Get-ChildItem $folder'\*.docx'


$wdExportFormat = [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF
$wdOpenAfterExport = $false
$wdExportOptimizeFor = [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint
$wdExportRange = [Microsoft.Office.Interop.Word.WdExportRange]::wdExportAllDocument
$wdStartPage = 0
$wdEndPage = 0
$wdExportItem = [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent
$wdIncludeDocProps = $false
$wdKeepIRM = $true
$wdCreateBookmarks = [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateWordBookmarks
$wdDocStructureTags = $true
$wdBitmapMissingFonts = $true
$wdUseISO19005_1 = $false

$wdApplication = $null;
$wdDocument = $null;

foreach ($file in $files){

       $wdApplication = New-Object -ComObject "Word.Application"

       $wdDocument = $wdApplication.Documents.Open($file.FullName, 0, $true)
       $name = ($wdDocument.Fullname).replace("docx","pdf")

       echo $wdDocument.Fullname
       $wdExportFile = $name
       echo $wdExportFile

       $wdDocument.ExportAsFixedFormat(
       $wdExportFile,
       $wdExportFormat,
       $wdOpenAfterExport,
       $wdExportOptimizeFor,
       $wdExportRange,
       $wdStartPage,
       $wdEndPage,
       $wdExportItem,
       $wdIncludeDocProps,
       $wdKeepIRM,
       $wdCreateBookmarks,
       $wdDocStructureTags,
       $wdBitmapMissingFonts,
       $wdUseISO19005_1
       )



}

if ($wdDocument)
{
        $wdDoNotSaveChanges = $false
        $wdDocument.Close($wdDoNotSaveChanges)
        $wdDocument = $null
}
if ($wdApplication)
{
        $wdApplication.Quit()
        $wdApplication = $null
}
