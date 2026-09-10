param (
    [string]$TargetName
)

try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    if ($excel) {
        Write-Host "Connected to active Excel instance."
        foreach ($wb in $excel.Workbooks) {
            Write-Host "Open workbook: $($wb.Name)"
            if ($wb.Name -like "*$TargetName*") {
                Write-Host "Closing $($wb.Name) without saving..."
                $wb.Close($false)
                Write-Host "Closed $($wb.Name)."
            }
        }
    }
} catch {
    Write-Host "No active Excel COM instance found or error: $($_.Exception.Message)"
}
