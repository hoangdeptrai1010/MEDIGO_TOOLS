import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

# Let's inspect baocaokpi_thang8_hoanthien.xlsx (formulas vs values) for Phượng and Tài
wb_form = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=False)
wb_val = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=True)

ws_f = wb_form['Dự án T8']
ws_v = wb_val['Dự án T8']

for r in range(1, ws_f.max_row + 1):
    name = ws_f.cell(r, 2).value
    if name in ['Trịnh Thị Phượng', 'Phan Công Vũ Tài', 'Vũ Thanh Hằng']:
        print(f"\n--- Row {r}: {name} ({ws_f.cell(r, 1).value}) ---")
        print("Formulas:")
        print(f"  Col D (CK+Combo+NY3/ngày): {ws_f.cell(r, 4).value}")
        print(f"  Col E (NY3): {ws_f.cell(r, 5).value}")
        print(f"  Col F (CK): {ws_f.cell(r, 6).value}")
        print(f"  Col G (Combo): {ws_f.cell(r, 7).value}")
        print(f"  Col H (CK/ngày): {ws_f.cell(r, 8).value}")
        print(f"  Col I (Combo/ngày): {ws_f.cell(r, 9).value}")
        print(f"  Col J (Thưởng DA): {ws_f.cell(r, 10).value}")
        print(f"  Col K (Thưởng Thêm): {ws_f.cell(r, 11).value}")
        print(f"  Col L (Hot Bill): {ws_f.cell(r, 12).value}")
        print(f"  Col M (Total DA): {ws_f.cell(r, 13).value}")
        print("Cached Values:")
        print(f"  Col D: {ws_v.cell(r, 4).value}")
        print(f"  Col E: {ws_v.cell(r, 5).value}")
        print(f"  Col F: {ws_v.cell(r, 6).value}")
        print(f"  Col G: {ws_v.cell(r, 7).value}")
        print(f"  Col H: {ws_v.cell(r, 8).value}")
        print(f"  Col I: {ws_v.cell(r, 9).value}")
        print(f"  Col J: {ws_v.cell(r, 10).value}")
        print(f"  Col K: {ws_v.cell(r, 11).value}")
        print(f"  Col L: {ws_v.cell(r, 12).value}")
        print(f"  Col M: {ws_v.cell(r, 13).value}")

