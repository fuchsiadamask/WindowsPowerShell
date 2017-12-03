function Ppt-To-Pdf
{
    Add-Type -AssemblyName Office -ErrorAction SilentlyContinue
    Add-Type -AssemblyName Microsoft.Office.Interop.Powerpoint `
                -ErrorAction SilentlyContinue

    $ppt = New-Object -COM Powerpoint.Application

    $msoFalse = [Microsoft.Office.Core.MsoTriState]::msoFalse

    $pdf = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF

    function As-Pdf([string]$ifile, [string]$ofile)
    {
        # 4th arg -- do not show window:
        $pres = $ppt.Presentations.Open($ifile, $msoFalse, $msoFalse, $msoFalse)
        $pres.SaveAs($ofile, $pdf)
        $pres.Close()
        $pres = $null
    }

    foreach ($file in $args)
    {
        $out_file = [System.IO.Path]::ChangeExtension($file, ".pdf")
        As-Pdf -ifile $file -ofile $out_file
    }

    $ppt.Quit()
    $ppt = $null

    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}

Export-ModuleMember -Function Ppt-To-Pdf