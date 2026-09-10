import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=True)
if 'Dự án T8' in wb.sheetnames:
    ws = wb['Dự án T8']
    print(f"Total rows in 'Dự án T8': {ws.max_row}")
    for r in range(1, ws.max_row + 1):
        branch = ws.cell(r, 1).value
        staff = ws.cell(r, 2).value
        role = ws.cell(r, 3).value
        thuong_da = ws.cell(r, 10).value
        thuong_them = ws.cell(r, 11).value
        hotbill = ws.cell(r, 12).value
        total_da = ws.cell(r, 13).value
        if r <= 3 or (total_da is not None and total_da > 0):
            print(f"Row {r:2d}: {branch} | {staff} | {role} | Thưởng DA: {thuong_da} | Thưởng Thêm: {thuong_them} | HotBill: {hotbill} | Total: {total_da}")
