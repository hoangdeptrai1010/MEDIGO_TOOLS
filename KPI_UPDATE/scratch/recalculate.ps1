$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

$files = @(
    "d:\MEDIGO\KPI_UPDATE\thang8\bangluong_thang8_hoanthien.xlsx",
    "d:\MEDIGO\KPI_UPDATE\thang8\BANGLUONGTHANG8.xlsx"
)

foreach ($f in $files) {
    if (Test-Path $f) {
        Write-Host "Opening and calculating: $f"
        $wb = $excel.Workbooks.Open($f, 0, $false)
        $excel.CalculateFull()
        $wb.Save()
        $wb.Close($true)
        Write-Host "Done: $f"
    }
}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "All files recalculated successfully!"
