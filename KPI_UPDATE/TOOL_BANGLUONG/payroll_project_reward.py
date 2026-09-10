"""
MODULE 4: TÍNH THƯỞNG DỰ ÁN NHÓM HÀNG & THƯỞNG NÓNG (HOT BILL)
- Đổ dữ liệu thưởng dự án nhóm hàng (CK, PartySmart, LadyCare, Combo)
- Đổ thưởng chương trình Hot Bill Hà Nội (15-31/8)
- Đổ dữ liệu và gắn công thức vào sheet 'Dự án'
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

def process_project_reward_sheet(wb_out, kpi_file=None):
    if 'Dự án' not in wb_out.sheetnames:
        return

    print(f"--> [Module Thưởng Dự Án] Đang lấy số liệu Thưởng dự án & Hot Bill từ KPI...")
    proj_map = defaultdict(lambda: {'duan': 0.0, 'hotbill': 0.0})

    if kpi_file:
        try:
            wb_kpi = openpyxl.load_workbook(kpi_file, data_only=True)
            
            # 1. Quét thưởng nhóm hàng từ sheet 'Dự án' hoặc 'Dự án T8' trong file KPI
            for sname in ['Dự án T8', 'Dự án', 'Du an', 'kpi dược sĩ']:
                if sname in wb_kpi.sheetnames:
                    ws_kpi_proj = wb_kpi[sname]
                    for r in range(2, ws_kpi_proj.max_row + 1):
                        br = ws_kpi_proj.cell(r, 1).value or ws_kpi_proj.cell(r, 2).value
                        nm = ws_kpi_proj.cell(r, 2).value or ws_kpi_proj.cell(r, 3).value
                        rw = to_num(ws_kpi_proj.cell(r, 3).value)
                        if br and nm and rw > 0:
                            key = (normalize_branch(br), str(nm).strip())
                            proj_map[key]['duan'] += rw

            # 2. Quét thưởng Hot Bill Hà Nội
            if 'Hot Bill HN' in wb_kpi.sheetnames:
                ws_hb = wb_kpi['Hot Bill HN']
                for r in range(2, ws_hb.max_row + 1):
                    br = ws_hb.cell(r, 1).value
                    nm = ws_hb.cell(r, 2).value
                    rw = to_num(ws_hb.cell(r, 3).value)
                    if br and nm and rw > 0:
                        key = (normalize_branch(br), str(nm).strip())
                        proj_map[key]['hotbill'] += rw

            wb_kpi.close()
        except Exception as e:
            print(f"Warning reading KPI Project rewards: {e}")

    # Đổ dữ liệu vào sheet 'Dự án'
    ws_out_proj = wb_out['Dự án']
    for r in range(2, ws_out_proj.max_row + 1):
        br = ws_out_proj.cell(r, 1).value
        nm = ws_out_proj.cell(r, 2).value
        if br and nm and str(nm).strip() and str(nm).strip() != 'Tổng':
            key = (normalize_branch(br), str(nm).strip())
            d = proj_map[key]
            if d['duan'] > 0 or d['hotbill'] > 0:
                ws_out_proj.cell(r, 4, d['duan']).number_format = '#,##0'
                ws_out_proj.cell(r, 5, d['hotbill']).number_format = '#,##0'
            ws_out_proj.cell(r, 3, f"=SUM(D{r}:E{r})").number_format = '#,##0'

    print("✅ Đã cập nhật xong sheet 'Dự án'!")
