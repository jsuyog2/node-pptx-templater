$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.DisplayAlerts = 1 # ppAlertsAll

$resolvedPath = Resolve-Path "scratch/test-table2-nesting-output.pptx"
$resolvedSavedPath = [System.IO.Path]::GetFullPath("scratch/test-table2-nesting-saved.pptx")

try {
    $pres = $pptApp.Presentations.Open($resolvedPath, 1, 0, 0)
    $pres.SaveAs($resolvedSavedPath)
    $pres.Close()
    Write-Host "✅ Success: Saved test-table2-nesting-saved.pptx"
} catch {
    Write-Host "❌ Error: $_"
}

$pptApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
