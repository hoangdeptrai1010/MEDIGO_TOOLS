import openpyxl, subprocess, shutil

print("=== BẮT ĐẦU CHUẨN HÓA BÁO CÁO KPI THÁNG 8 TỪ FILE CHUẨN GỐC ===")

# Load the authentic Month 8 file
wb = openpyxl.load_workbook('NHÀ THUỐC THÁNG 8 2026.xlsx', data_only=False)

ws_ds = wb['kpi dược sĩ']
ws_da = wb['Dự án T8']

# 1. Fix Trịnh Thị Phượng branch mismatch:
# In kpi dược sĩ, Trịnh Thị Phượng primarily worked at Trường Sa (matching Dự án T8 row 17)
for r in range(3, ws_ds.max_row + 1):
    nv = ws_ds.cell(r, 2).value
    if nv == 'Trịnh Thị Phượng':
        print(f"Fixing branch for Trịnh Thị Phượng in kpi dược sĩ: {ws_ds.cell(r, 1).value} -> Trường Sa")
        ws_ds.cell(r, 1).value = 'Trường Sa'

# 2. Fix Column M formula in kpi dược sĩ to safely lookup by employee name if branch varies
for r in range(3, ws_ds.max_row + 1):
    nv = ws_ds.cell(r, 2).value
    if nv:
        # Use XLOOKUP by employee name $B{r} in Dự án T8!$B:$B to be 100% robust against multi-branch transfers
        ws_ds.cell(r, 13).value = f'=IFERROR(_xlfn.XLOOKUP($B{r}, \'Dự án T8\'!$B:$B, \'Dự án T8\'!$M:$M, 0)*$L{r}, 0)'
        # Col N: N(K) + M (clean addition)
        ws_ds.cell(r, 14).value = f'=N(K{r})+M{r}'

# 3. Ensure Hot Bill in Dự án T8 uses clean SUMIF
for r in range(3, ws_da.max_row + 1):
    nv = ws_da.cell(r, 2).value
    if nv:
        ws_da.cell(r, 12).value = f'=IFERROR(SUMIF(\'Hot Bill HN\'!$D:$D, $B{r}, \'Hot Bill HN\'!$G:$G), 0)'
        ws_da.cell(r, 13).value = f'=IFERROR($J{r}+$K{r}+$L{r}, 0)'

# Save to baocaokpi_thang8_hoanthien.xlsx
out_file = 'baocaokpi_thang8_hoanthien.xlsx'
wb.save(out_file)
wb.close()
print(f"Saved: {out_file}")

# Copy to thang8 directory
shutil.copy2(out_file, 'thang8/baocaokpi_thang8_hoanthien.xlsx')
print("Copied to thang8/baocaokpi_thang8_hoanthien.xlsx")

# Recalculate via Excel COM to store 100% cached numbers
res = subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-File', 'recalc_workbook.ps1', '-FilePath', out_file], capture_output=True, text=True)
print("Recalc baocaokpi_thang8_hoanthien.xlsx:", "SUCCESS" if "successfully" in res.stdout else "FAILED")

# Also recalculate the copy in thang8
res2 = subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-File', 'recalc_workbook.ps1', '-FilePath', 'thang8/baocaokpi_thang8_hoanthien.xlsx'], capture_output=True, text=True)
print("Recalc thang8/baocaokpi_thang8_hoanthien.xlsx:", "SUCCESS" if "successfully" in res2.stdout else "FAILED")
