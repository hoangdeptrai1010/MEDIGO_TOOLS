$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $src = "D:\MEDIGO\KPI_UPDATE\NHÀ THUỐC THÁNG 8 2026.xlsx"
    Write-Host "Opening $src..."
    $wb = $excel.Workbooks.Open($src)
    Write-Host "Opened successfully!"

    $ws_ds = $wb.Sheets.Item("kpi dược sĩ")
    Write-Host "Row 26 name: " $ws_ds.Cells.Item(26, 2).Text
    Write-Host "Row 26 branch: " $ws_ds.Cells.Item(26, 1).Text
    Write-Host "Row 26 Col M (before): " $ws_ds.Cells.Item(26, 13).Text

    # Fix branch for Trịnh Thị Phượng to Trường Sa
    $ws_ds.Cells.Item(26, 1).Value2 = "Trường Sa"

    $excel.CalculateFull()

    Write-Host "Row 26 Col M (after): " $ws_ds.Cells.Item(26, 13).Text
    Write-Host "Row 26 Col N (after): " $ws_ds.Cells.Item(26, 14).Text

    $dst1 = "D:\MEDIGO\KPI_UPDATE\baocaokpi_thang8_hoanthien.xlsx"
    $dst2 = "D:\MEDIGO\KPI_UPDATE\thang8\baocaokpi_thang8_hoanthien.xlsx"

    # Save to both paths
    $wb.SaveCopyAs($dst1)
    $wb.SaveCopyAs($dst2)
    Write-Host "Saved successfully to $dst1 and $dst2!"

    $wb.Close($false)
}
catch {
    Write-Host "Error: $_"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
