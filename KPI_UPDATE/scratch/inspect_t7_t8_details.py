import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

print("=" * 100)
print("CHI TIẾT THÁNG 7: GỐC vs TARGET vs HOÀN THIỆN")
print("=" * 100)

wb7_goc = openpyxl.load_workbook(r'goc\NHÀ THUỐC THÁNG 7 2026.xlsx', data_only=False)
wb7_goc_val = openpyxl.load_workbook(r'goc\NHÀ THUỐC THÁNG 7 2026.xlsx', data_only=True)

for sname in ['Dự án T7', 'Thưởng CK', 'kpi nhà thuốc', 'Chiết khấu - combo - NY3']:
    if sname in wb7_goc.sheetnames:
        ws_f = wb7_goc[sname]
        ws_v = wb7_goc_val[sname]
        print(f"\n--- [GỐC T7] Sheet: {sname} ---")
        for r in range(1, 6):
            f_row = [ws_f.cell(r, c).value for c in range(1, 15)]
            v_row = [ws_v.cell(r, c).value for c in range(1, 15)]
            print(f"Row {r} FORMULA: {f_row}")
            print(f"Row {r} VALUE:   {v_row}")

wb7_bl = openpyxl.load_workbook(r'thang7\target\BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=False)
wb7_bl_val = openpyxl.load_workbook(r'thang7\target\BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=True)

for sname in ['BẢNG LƯƠNG', 'Thưởng CK', 'Dự án']:
    if sname in wb7_bl.sheetnames:
        ws_f = wb7_bl[sname]
        ws_v = wb7_bl_val[sname]
        print(f"\n--- [TARGET BL T7] Sheet: {sname} ---")
        for r in range(1, 6):
            f_row = [ws_f.cell(r, c).value for c in range(1, 15)]
            v_row = [ws_v.cell(r, c).value for c in range(1, 15)]
            print(f"Row {r} FORMULA: {f_row}")
            print(f"Row {r} VALUE:   {v_row}")

print("\n" + "=" * 100)
print("CHI TIẾT THÁNG 8: GỐC vs HOÀN THIỆN")
print("=" * 100)

wb8_goc = openpyxl.load_workbook(r'goc\NHÀ THUỐC THÁNG 8 2026.xlsx', data_only=False)
wb8_goc_val = openpyxl.load_workbook(r'goc\NHÀ THUỐC THÁNG 8 2026.xlsx', data_only=True)

for sname in ['Dự án T8', 'kpi nhà thuốc', 'Chiết khấu - combo - NY3']:
    if sname in wb8_goc.sheetnames:
        ws_f = wb8_goc[sname]
        ws_v = wb8_goc_val[sname]
        print(f"\n--- [GỐC T8] Sheet: {sname} ---")
        for r in range(1, 6):
            f_row = [ws_f.cell(r, c).value for c in range(1, 15)]
            v_row = [ws_v.cell(r, c).value for c in range(1, 15)]
            print(f"Row {r} FORMULA: {f_row}")
            print(f"Row {r} VALUE:   {v_row}")

wb8_bl = openpyxl.load_workbook(r'thang8\bangluong_thang8_hoanthien.xlsx', data_only=False)
wb8_bl_val = openpyxl.load_workbook(r'thang8\bangluong_thang8_hoanthien.xlsx', data_only=True)

for sname in ['BẢNG LƯƠNG', 'Thưởng CK', 'Dự án', 'MiniKat - HN', 'MiniKat - HCM']:
    if sname in wb8_bl.sheetnames:
        ws_f = wb8_bl[sname]
        ws_v = wb8_bl_val[sname]
        print(f"\n--- [BL T8 HOÀN THIỆN] Sheet: {sname} ---")
        for r in range(1, 6):
            f_row = [ws_f.cell(r, c).value for c in range(1, 15)]
            v_row = [ws_v.cell(r, c).value for c in range(1, 15)]
            print(f"Row {r} FORMULA: {f_row}")
            print(f"Row {r} VALUE:   {v_row}")
