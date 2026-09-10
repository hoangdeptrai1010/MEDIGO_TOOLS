"""
MODULE 9: TÍNH THƯỞNG CHƯƠNG TRÌNH MAPS (ĐÁNH GIÁ GOOGLE MAPS)
- Số lượng đánh giá Maps theo từng Dược sĩ & Chi nhánh
- Thưởng Maps theo quy định
- Đổ dữ liệu và gắn công thức vào sheet 'Map'
"""

import openpyxl

def process_maps_sheet(wb_out):
    if 'Map' not in wb_out.sheetnames:
        return

    print(f"--> [Module Maps] Đang cập nhật sheet 'Map'...")
    ws_map = wb_out['Map']
    for r in range(2, ws_map.max_row + 1):
        nm = ws_map.cell(r, 1).value
        if nm and str(nm).strip() and str(nm).strip() != 'Tổng':
            # Gắn công thức tính thưởng Maps
            ws_map.cell(r, 4, f"=IF(C{r}>0, C{r}*10000, 0)").number_format = '#,##0'

    print("✅ Đã cập nhật xong sheet 'Map'!")
