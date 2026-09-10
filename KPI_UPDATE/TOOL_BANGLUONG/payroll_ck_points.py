"""
MODULE 7: TÍNH THƯỞNG HÀNG ĐIỂM CHIẾT KHẤU (CK) & CẮT LIỀU
- Cập nhật thưởng hàng điểm 50,000 đ/hộp cho nhân sự Hà Nội
- Áp dụng hệ số đánh giá / thưởng hoa hồng (0.8 - 1.0)
- Đổ dữ liệu và gắn công thức vào sheet 'Thưởng CK'
"""

import openpyxl
from collections import defaultdict

HN_BRANCHES = {'Hàng Bông', 'Đường Láng'}

def process_ck_points_sheet(wb_out):
    if 'Thưởng CK' not in wb_out.sheetnames:
        return

    print(f"--> [Module Thưởng CK] Đang tính thưởng hàng điểm CK & cắt liều...")
    ws_ck = wb_out['Thưởng CK']

    for r in range(2, ws_ck.max_row + 1):
        br = ws_ck.cell(r, 1).value
        nm = ws_ck.cell(r, 2).value
        if br and nm and str(nm).strip() and str(nm).strip() != 'Tổng':
            # Nếu là nhân sự Hà Nội, cập nhật thưởng hàng điểm chuẩn
            if str(br).strip() in HN_BRANCHES:
                ws_ck.cell(r, 5, 50000).number_format = '#,##0'
            ws_ck.cell(r, 3, f"=E{r}*D{r}").number_format = '#,##0'

    print("✅ Đã cập nhật xong sheet 'Thưởng CK'!")
