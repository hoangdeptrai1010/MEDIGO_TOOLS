$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('D:\MEDIGO\KPI_UPDATE\baocaokpi_thang8_hoanthien.xlsx')
$ws = $wb.Worksheets.Item('Dự án T8')
for ($r = 3; $r -le 11; $r++) {
    $nv = $ws.Cells.Item($r, 2).Text
    $da = $ws.Cells.Item($r, 10).Text
    $them = $ws.Cells.Item($r, 11).Text
    $hb = $ws.Cells.Item($r, 12).Text
    $tot = $ws.Cells.Item($r, 13).Text
    Write-Host ('{0,-25} | DA={1,10} | Thêm={2,10} | HotBill={3,10} | Total={4,10}' -f $nv, $da, $them, $hb, $tot)
}
$wb.Close($false)
$excel.Quit()
