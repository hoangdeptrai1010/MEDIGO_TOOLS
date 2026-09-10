$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open("D:\MEDIGO\KPI_UPDATE\baocaokpi_thang7_hoanthien.xlsx")
$excel.CalculateFullRebuild()
$ws_ds = $wb.Worksheets.Item("kpi dược sĩ")
for ($r = 3; $r -le 7; $r++) {
    Write-Host "Row $r Target(H): $($ws_ds.Cells.Item($r, 8).Value2) Act(I): $($ws_ds.Cells.Item($r, 9).Value2) Bonus(K): $($ws_ds.Cells.Item($r, 11).Value2)"
}
$ws_da = $wb.Worksheets.Item("Dự án T7")
for ($r = 3; $r -le 7; $r++) {
    Write-Host "DA Row $r CK(F): $($ws_da.Cells.Item($r, 6).Value2) TierBonus(J): $($ws_da.Cells.Item($r, 10).Value2) Tot(M): $($ws_da.Cells.Item($r, 13).Value2)"
}
$wb.Save()
$wb.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
