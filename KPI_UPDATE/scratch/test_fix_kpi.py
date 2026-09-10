import openpyxl
import subprocess
import os

wb = openpyxl.load_workbook('baocaokpi_thang7_hoanthien.xlsx')
ws_ds = wb['kpi dược sĩ']

# Fix any formula in ws_ds referencing $T$1
for (r, c), cell in ws_ds._cells.items():
    if isinstance(cell.value, str) and '$T$1' in cell.value:
        cell.value = cell.value.replace("'kpi nhà thuốc'!$T$1", "$C$1")

wb.save('test_fixed_kpi.xlsx')
wb.close()

p = os.path.abspath('test_fixed_kpi.xlsx')
res = subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-File', 'recalc_workbook.ps1', '-FilePath', p], capture_output=True, text=True)
print(f'Test fixed KPI return: {res.returncode}')
print('Output:\n', res.stdout)
