"""
MODULE 6: TÍNH THƯỞNG KPI DOANH SỐ DƯỢC SĨ & NHÀ THUỐC
- Đổ Doanh thu mục tiêu, Doanh thu thực tế, % Đạt và Thưởng KPI từ Báo Cáo KPI
- Áp dụng các mốc tỷ lệ thưởng theo chính sách KPI
- Đổ dữ liệu và gắn công thức vào sheet 'KPI'
"""

import openpyxl
from collections import defaultdict

BRANCH_MAP = {
    'NT Hàng Bông 247': 'Hàng Bông', 'NT Hàng Bông': 'Hàng Bông', 'Hàng Bông': 'Hàng Bông',
    'NT Đường Láng 247': 'Đường Láng', 'NT Đường Láng': 'Đường Láng', 'Đường Láng': 'Đường Láng',
    'NT 24H Minh Châu 1': 'Minh Châu', 'Minh Châu 1': 'Minh Châu', 'Minh Châu': 'Minh Châu',
    'NT 24H Nam Hòa': 'Nam Hòa', 'Nam Hòa': 'Nam Hòa',
    'NT 24H Nguyễn Chí Thanh': 'Nguyễn Chí Thanh', 'Nguyễn Chí Thanh': 'Nguyễn Chí Thanh',
    'NT 24H Nguyễn Thị Thập': 'Nguyễn Thị Thập', 'Nguyễn Thị Thập': 'Nguyễn Thị Thập',
    'NT 24H Nguyễn Văn Quá': 'Nguyễn Văn Quá', 'Nguyễn Văn Quá': 'Nguyễn Văn Quá',
    'NT 24H Rạch Bùng Binh': 'Rạch Bùng Binh', 'Rạch Bùng Binh': 'Rạch Bùng Binh',
    'NT 24H Trường Sa': 'Trường Sa', 'Trường Sa': 'Trường Sa',
    'NT 24h Đỗ Quang Đẩu': 'Đỗ Quang Đẩu', 'Đỗ Quang Đẩu': 'Đỗ Quang Đẩu',
    'NT 24H Lê Bình': 'Lê Bình', 'Lê Bình': 'Lê Bình',
}

def normalize_branch(b):
    if not b: return 'Trường Sa'
    b_clean = str(b).strip()
    return BRANCH_MAP.get(b_clean, b_clean)

def to_num(v, default=0.0):
    if v is None: return default
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(',', '')
    try: return float(s)
    except: return default

def process_kpi_sheet(wb_out, kpi_file=None):
    if 'KPI' not in wb_out.sheetnames:
        return

    print(f"--> [Module Thưởng KPI] Đang nạp số liệu KPI doanh số...")
    kpi_map = defaultdict(lambda: {'target': 0.0, 'actual': 0.0, 'reward': 0.0})

    if kpi_file:
        try:
            wb_kpi = openpyxl.load_workbook(kpi_file, data_only=True)
            for sname in ['kpi dược sĩ', 'KPI dược sĩ', 'KPI Duoc si', 'kpi duoc si']:
                if sname in wb_kpi.sheetnames:
                    ws_kpi_staff = wb_kpi[sname]
                    for r in range(2, ws_kpi_staff.max_row + 1):
                        br = ws_kpi_staff.cell(r, 1).value or ws_kpi_staff.cell(r, 2).value
                        nm = ws_kpi_staff.cell(r, 2).value or ws_kpi_staff.cell(r, 3).value
                        tgt = to_num(ws_kpi_staff.cell(r, 4).value)
                        act = to_num(ws_kpi_staff.cell(r, 5).value)
                        rw = to_num(ws_kpi_staff.cell(r, 7).value)
                        if br and nm:
                            key = (normalize_branch(br), str(nm).strip())
                            kpi_map[key] = {'target': tgt, 'actual': act, 'reward': rw}
            wb_kpi.close()
        except Exception as e:
            print(f"Warning reading KPI staff rewards: {e}")

    # Đổ dữ liệu vào sheet 'KPI'
    ws_kpi_out = wb_out['KPI']
    for r in range(2, ws_kpi_out.max_row + 1):
        br = ws_kpi_out.cell(r, 1).value
        nm = ws_kpi_out.cell(r, 2).value
        if br and nm and str(nm).strip() and str(nm).strip() != 'Tổng':
            key = (normalize_branch(br), str(nm).strip())
            d = kpi_map[key]
            if d['target'] > 0 or d['actual'] > 0:
                ws_kpi_out.cell(r, 4, d['target']).number_format = '#,##0'
                ws_kpi_out.cell(r, 5, d['actual']).number_format = '#,##0'
            ws_kpi_out.cell(r, 6, f"=IF(D{r}>0, E{r}/D{r}, 0)").number_format = '0.00%'
            ws_kpi_out.cell(r, 7, f"=IF(F{r}>=1.2, 0.05*E{r}, IF(F{r}>=1.0, 0.03*E{r}, IF(F{r}>=0.8, 0.015*E{r}, 0)))").number_format = '#,##0'
            ws_kpi_out.cell(r, 3, f"=G{r}").number_format = '#,##0'

    print("✅ Đã cập nhật xong sheet 'KPI'!")
