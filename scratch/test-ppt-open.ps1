$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone

function Test-Presentation($filePath) {
    $resolvedPath = Resolve-Path $filePath
    Write-Host "Testing file: $resolvedPath"
    try {
        $pres = $pptApp.Presentations.Open($resolvedPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
        Write-Host "✅ Success: Opened presentation cleanly without repair errors."
        $pres.Close()
    } catch {
        Write-Host "❌ Error: Failed to open presentation. PowerPoint might have prompted for a repair. Details: $_"
    }
}

Test-Presentation "scratch/test-repair-output.pptx"
Test-Presentation "examples/output/basic-output.pptx"
Test-Presentation "scratch/test-merge-vertical-output.pptx"

$pptApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
