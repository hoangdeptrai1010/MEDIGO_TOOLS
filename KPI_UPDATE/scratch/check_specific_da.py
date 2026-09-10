import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws_d = wb['Dự án']
for r in range(1, ws_d.max_row + 1):
    cn = ws_d.cell(r, 1).value
    name = ws_d.cell(r, 2).value
    tot = ws_d.cell(r, 3).value
    r_name = ws_d.cell(r, 8).value
    r_tot = ws_d.cell(r, 9).value
    if name in ['Trịnh Thị Phượng', 'Phan Công Vũ Tài', 'Vũ Thanh Hằng', 'Hồ Thị Minh Hòa', 'Nguyễn Mạnh Tuấn'] or r_name in ['Trịnh Thị Phượng', 'Phan Công Vũ Tài', 'Vũ Thanh Hằng', 'Hồ Thị Minh Hòa', 'Nguyễn Mạnh Tuấn']:
        print(f"Row {r:2d}: Left ({cn} | {name} | {tot}) | Right ({r_name} | {r_tot})")
