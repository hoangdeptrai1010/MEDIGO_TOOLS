import openpyxl

for target_file in ['thang8/bangluong_thang8_hoanthien.xlsx', 'thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx']:
    wb = openpyxl.load_workbook(target_file)
    for ws in wb.worksheets:
        # Unhide all rows
        for dim in ws.row_dimensions.values():
            if dim.hidden:
                dim.hidden = False
        # Clear filter criteria while keeping the filter button if wanted, or clear filter criteria
        if ws.auto_filter:
            ws.auto_filter.filterColumn = []
            
    wb.save(target_file)
    print(f"Successfully unhid all rows and cleared filter criteria in {target_file}!")
