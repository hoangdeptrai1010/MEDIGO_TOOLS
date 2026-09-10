param (
    [string]$FilePath
)

if (-not (Test-Path $FilePath)) {
    Write-Host "File not found: $FilePath"
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false

try {
    $fullPath = [System.IO.Path]::GetFullPath($FilePath)
    Write-Host "Opening workbook: $fullPath"
    $wb = $excel.Workbooks.Open($fullPath)
    
    $excel.Calculation = -4105
    $wb.ForceFullCalculation = $true
    Write-Host "Calculating full rebuild on all $($wb.Worksheets.Count) sheets..."
    $excel.CalculateFullRebuild()
    foreach ($sh in $wb.Worksheets) {
        $sh.Calculate()
    }
    
    Write-Host "Saving recalculated workbook..."
    $wb.SaveAs($fullPath, 51)
    $wb.Close($false)
    Write-Host "Recalculation complete and cached values saved successfully!"
} catch {
    Write-Host "Error during calculation: $($_.Exception.ToString())"
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
