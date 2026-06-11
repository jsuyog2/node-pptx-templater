$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone

function Test-And-Save-Presentation($filePath, $savedPath) {
    $resolvedPath = Resolve-Path $filePath
    $resolvedSavedPath = [System.IO.Path]::GetFullPath($savedPath)
    Write-Host "Opening file: $resolvedPath"
    try {
        $pres = $pptApp.Presentations.Open($resolvedPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
        Write-Host "Opened successfully. Saving to: $resolvedSavedPath"
        $pres.SaveAs($resolvedSavedPath)
        $pres.Close()
        Write-Host "✅ Saved successfully."
    } catch {
        Write-Host "❌ Error during PowerPoint open/save: $_"
    }
}

Test-And-Save-Presentation "scratch/test-repair-output.pptx" "scratch/test-repair-saved.pptx"
Test-And-Save-Presentation "examples/output/basic-output.pptx" "scratch/basic-saved.pptx"

$pptApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
