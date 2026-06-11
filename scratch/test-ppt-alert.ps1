$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.DisplayAlerts = 1 # ppAlertsAll

function Test-Presentation($filePath) {
    $resolvedPath = Resolve-Path $filePath
    Write-Host "Opening file with alerts enabled: $resolvedPath"
    try {
        $pres = $pptApp.Presentations.Open($resolvedPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
        Write-Host "✅ Success: Opened presentation cleanly without repair dialog."
        $pres.Close()
    } catch {
        Write-Host "❌ Error: Failed to open presentation. Details: $_"
    }
}

Test-Presentation "examples/output/basic-output.pptx"

$pptApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
