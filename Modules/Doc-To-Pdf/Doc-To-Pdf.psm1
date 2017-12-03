function Doc-To-Pdf
{
    Add-Type -AssemblyName Office -ErrorAction SilentlyContinue
    Add-Type -AssemblyName Microsoft.Office.Interop.Word `
                -ErrorAction SilentlyContinue

    $word = New-Object -COM Word.Application

    # Do not show window:
    $word.Visible = $false

    $exportFormat = [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF
    $openAfterExport = $false
    $exportOptimizeFor = [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint
    $exportRange = [Microsoft.Office.Interop.Word.WdExportRange]::wdExportAllDocument
    $startPage = 0
    $endPage = 0
    $exportItem = [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent
    $includeDocProps = $true
    $keepIRM = $true
    $createBookmarks = [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateHeadingBookmarks
    $docStructureTags = $true
    $bitmapMissingFonts = $true
    $useISO19005_1 = $false

    $close = [Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges

    function As-Pdf([string]$ifile, [string]$ofile)
    {
        $doc = $word.Documents.Open($ifile)
        $doc.ExportAsFixedFormat($ofile,
                                 $exportFormat,
                                 $openAfterExport,
                                 $exportOptimizeFor,
                                 $exportRange,
                                 $startPage,
                                 $endPage,
                                 $exportItem,
                                 $includeDocProps,
                                 $keepIRM,
                                 $createBookmarks,
                                 $docStructureTags,
                                 $bitmapMissingFonts,
                                 $useISO19005_1)
        $doc.Close([ref]$close)
        $doc = $null
    }

    foreach ($file in $args)
    {
        $out_file = [System.IO.Path]::ChangeExtension($file, ".pdf")
        As-Pdf -ifile $file -ofile $out_file
    }

    $word.Quit()
    $word = $null

    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}

Export-ModuleMember -Function Doc-To-Pdf