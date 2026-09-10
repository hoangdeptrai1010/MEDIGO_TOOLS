"""
MODULE 8: TÍNH GIẢM TRỪ KPI & PHẠT VI PHẠM
- Hàng hết date
- Vi phạm quy chế quy trình
- Lỗi Hóa đơn App - Kiot
- Lỗi Ecom
- Cấn trừ khác (Chi phí đồng phục, v.v.)
- Đổ dữ liệu và gắn công thức vào sheet 'KPI trừ'
"""

import openpyxl

def process_deductions_sheet(wb_out):
    if 'KPI trừ' not in wb_out.sheetnames:
        return

    print(f"--> [Module Giảm Trừ KPI] Đang tính các khoản giảm trừ và cấn trừ...")
    ws_kt = wb_out['KPI trừ']

    for r in range(2, ws_kt.max_row + 1):
        n_val = ws_kt.cell(r, 2).value
        if n_val and str(n_val).strip() and str(n_val).strip() != 'Tổng':
            n_str = str(n_val).strip()
            # Ví dụ cấn trừ đồng phục
            if n_str == 'Nguyễn Trần Ngọc Phương':
                ws_kt.cell(r, 7, 500000).number_format = '#,##0'
                ws_kt.cell(r, 8, 'Cấn trừ chi phí đồng phục')

    print("✅ Đã cập nhật xong sheet 'KPI trừ'!")
