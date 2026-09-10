import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=False)
for s in ['Dự án', 'Thưởng CK']:
    ws = wb[s]
    print(f"\n--- Sheet: {s} Left table (Col A, B, C) ---")
    staff_count = {}
    for r in range(2, ws.max_row + 1):
        cn = ws.cell(r, 1).value
        name = ws.cell(r, 2).value
        formula_c = ws.cell(r, 3).value
        if name:
            staff_count[name] = staff_count.get(name, 0) + 1
            if staff_count[name] > 1 or name in ['Trịnh Thị Phượng', 'Phan Công Vũ Tài', 'Vũ Thanh Hằng']:
                print(f"Row {r:2d}: {cn} | {name} | C={formula_c}")

    print(f"Total staff with count > 1 in {s}: {[k for k, v in staff_count.items() if v > 1]}")
