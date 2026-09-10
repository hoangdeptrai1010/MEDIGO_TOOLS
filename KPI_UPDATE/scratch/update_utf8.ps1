$items = Get-ChildItem -Filter "*8 2026.xlsx"
$src = $items[0].FullName
Write-Host "Opening: $src"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($src)
    Write-Host "Opened successfully!"
    for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
        Write-Host ("Sheet " + $i + ": " + $wb.Sheets.Item($i).Name)
    }
    $ws_ds = $wb.Sheets.Item(2)
    Write-Host ("Row 26 name: " + $ws_ds.Cells.Item(26, 2).Text)
    Write-Host ("Row 26 branch before: " + $ws_ds.Cells.Item(26, 1).Text)
    Write-Host ("Row 26 M before: " + $ws_ds.Cells.Item(26, 13).Text)
    
    $ws_ds.Cells.Item(26, 1).Value2 = "Trường Sa"
    $excel.CalculateFull()
    
    Write-Host ("Row 26 branch after: " + $ws_ds.Cells.Item(26, 1).Text)
    Write-Host ("Row 26 M after: " + $ws_ds.Cells.Item(26, 13).Text)
    Write-Host ("Row 26 N after: " + $ws_ds.Cells.Item(26, 14).Text)
    
    $dst1 = (Join-Path (Get-Location) "baocaokpi_thang8_hoanthien.xlsx")
    $dst2 = (Join-Path (Get-Location) "thang8\baocaokpi_thang8_hoanthien.xlsx")
    
    $wb.SaveCopyAs($dst1)
    $wb.SaveCopyAs($dst2)
    Write-Host "Saved both files successfully!"
    $wb.Close($false)
} catch {
    Write-Host "Error: $_"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
