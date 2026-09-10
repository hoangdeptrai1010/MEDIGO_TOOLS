import openpyxl

for p in ['thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', 'thang8/bangluong_thang8_hoanthien.xlsx']:
    wb = openpyxl.load_workbook(p, data_only=False)
    ws = wb['Ngày công']
    print(f"=== {p} -> Ngày công ===")
    print("auto_filter.ref:", ws.auto_filter.ref)
    print("auto_filter.filterColumn:", ws.auto_filter.filterColumn)
    if ws.row_dimensions:
        hidden_rows = [r for r, dim in ws.row_dimensions.items() if dim.hidden]
        print(f"Total hidden rows in Ngày công: {len(hidden_rows)}")
        if hidden_rows:
            print("Sample hidden rows:", hidden_rows[:20])
