import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import json
import sys
import io
import re
import glob
import calendar
import datetime
from collections import defaultdict
from kpi_styling import apply_full_kpi_styling

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==============================================================================
# PROFESSIONAL COLOR PALETTE & STYLING TOKENS
# ==============================================================================
font_header_white = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
font_header_dark = Font(name='Calibri', size=11, bold=True, color='1E293B')
font_header_green = Font(name='Calibri', size=11, bold=True, color='166534')
font_bold = Font(name='Calibri', size=11, bold=True)
font_regular = Font(name='Calibri', size=11)

# Header Fills
fill_navy = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')      # Primary Navy
fill_soft_blue = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid') # Soft Blue
fill_light_blue = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')# Light Blue
fill_yellow = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')    # KPI / Highlight Yellow
fill_green = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')     # Bonus / Success Green
fill_emerald = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')   # Strong Green
fill_orange = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')    # Target / Hàng điểm Orange
fill_purple = PatternFill(start_color='E8D8F8', end_color='E8D8F8', fill_type='solid')    # Project Purple
fill_gray = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')      # Subtle Gray
fill_dark_gray = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid') # Dark Gray

# Borders
border_thin = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
border_header = Border(
    left=Side(style='thin', color='B0C4DE'),
    right=Side(style='thin', color='B0C4DE'),
    top=Side(style='medium', color='1F4E79'),
    bottom=Side(style='medium', color='1F4E79')
)
border_subtotal = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='808080'),
    bottom=Side(style='double', color='808080')
)

def remove_accents(input_str):
    import unicodedata
    if not input_str:
        return ""
    nfkd_form = unicodedata.normalize('NFKD', str(input_str))
    res = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return res.replace('đ', 'd').replace('Đ', 'D')

def clean_branch_name(cn):
    if not cn:
        return ''
    cn = str(cn).strip()
    mapping = {
        'NT Hàng Bông 247': 'Hàng Bông',
        'NT Đường Láng 247': 'Đường Láng',
        'NT 24H Minh Châu 1': 'Minh Châu',
        'NT 24h Đỗ Quang Đẩu': 'Đỗ Quang Đẩu',
        'NT 24H Lê Bình': 'Lê Bình',
        'NT 24H Nam Hòa': 'Nam Hòa',
        'NT 24H Nguyễn Chí Thanh': 'Nguyễn Chí Thanh',
        'NT 24H Nguyễn Thị Thập': 'Nguyễn Thị Thập',
        'NT 24H Nguyễn Văn Quá': 'Nguyễn Văn Quá',
        'NT 24H Rạch Bùng Binh': 'Rạch Bùng Binh',
        'NT 24H Trường Sa': 'Trường Sa'
    }
    if cn in mapping:
        return mapping[cn]
    cleaned = re.sub(r'^NT\s+(24H\s+|24h\s+)?', '', cn, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s+247$', '', cleaned)
    cleaned = re.sub(r'\s+1$', '', cleaned)
    return cleaned.strip()

def detect_proposal_month(hcm_file_path=None, hn_file_path=None, requested_month=None):
    """
    Tự động phát hiện tháng từ tên file và danh sách sheet bên trong file đề xuất HCM & HN.
    Đảm bảo nếu đưa file Tháng 9 vào thì không bao giờ bị sinh nhầm thành Tháng 8.
    """
    # 1. Kiểm tra từ tên file
    for fpath in [hcm_file_path, hn_file_path]:
        if fpath and os.path.exists(fpath):
            fname = os.path.basename(fpath)
            m = re.search(r'(?:th[aá]ng|t)\s*0?([1-9]|1[0-2])', remove_accents(fname), re.IGNORECASE)
            if m:
                detected = int(m.group(1))
                print(f"--> [Auto Detect] Phát hiện Tháng {detected} từ tên file: '{fname}'")
                return detected

    # 2. Kiểm tra từ sheet names của HCM file
    if hcm_file_path and os.path.exists(hcm_file_path):
        try:
            wb = openpyxl.load_workbook(hcm_file_path, read_only=True)
            for sname in wb.sheetnames:
                s_norm = remove_accents(sname).lower()
                m = re.search(r'kpi\s*(?:thang|t)\s*0?([1-9]|1[0-2])', s_norm)
                if m:
                    detected = int(m.group(1))
                    print(f"--> [Auto Detect] Phát hiện Tháng {detected} từ Sheet HCM: '{sname}'")
                    wb.close()
                    return detected
            wb.close()
        except Exception:
            pass

    # 3. Fallback to requested_month or 9
    return int(requested_month) if requested_month else 9

# ==============================================================================
# 1. PARSE ĐỀ XUẤT HCM (DYNAMIC HEADER PARSER)
# ==============================================================================
def parse_hcm_proposal(hcm_file_path, month_num=9):
    if not os.path.exists(hcm_file_path):
        raise FileNotFoundError(f"Không tìm thấy file đề xuất HCM: {hcm_file_path}")
        
    wb = openpyxl.load_workbook(hcm_file_path, data_only=True)
    
    target_sheet = None
    # 1. Tìm chính xác sheet tháng month_num
    for name in wb.sheetnames:
        n_low = remove_accents(name).lower()
        if 'kpi' in n_low and (f'thang {month_num:02d}' in n_low or f'thang {month_num}' in n_low or f't{month_num}' in n_low):
            target_sheet = name
            break
            
    # 2. Nếu không có, tìm bất kỳ sheet nào có chữ kpi
    if not target_sheet:
        for name in wb.sheetnames:
            n_low = remove_accents(name).lower()
            if 'kpi' in n_low:
                target_sheet = name
                break
    if not target_sheet:
        target_sheet = wb.sheetnames[-1]
        
    ws = wb[target_sheet]
    print(f"--> [HCM Parser] Đọc Sheet: '{target_sheet}' (Tháng {month_num})")

    # Detect header row (row 1 to 4)
    h_row = 1
    for r in range(1, min(5, ws.max_row + 1)):
        row_str = " ".join([str(ws.cell(r, c).value or '') for c in range(1, 15)])
        row_str_norm = remove_accents(row_str).lower()
        if 'nhan vien' in row_str_norm or 'duoc si' in row_str_norm:
            h_row = r
            break

    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(h_row, c).value
        if v is not None and str(v).strip():
            headers[c] = remove_accents(str(v)).lower().strip()

    col_store = next((c for c, h in headers.items() if 'nha thuoc' in h or 'chi nhanh' in h), 1)
    col_name = next((c for c, h in headers.items() if 'nhan vien' in h or 'duoc si' in h), 2)
    col_role = next((c for c, h in headers.items() if 'chuc danh' in h or 'vi tri' in h), 3)
    col_tb = next((c for c, h in headers.items() if 'trung binh bill' in h or 'tb bill' in h or 'tbbill' in h), 4)
    col_gd = next((c for c, h in headers.items() if 'giao dich' in h or 'gd' in h), 5)
    
    # Target month: match f"thang {month_num}" or f"thang {month_num:02d}" or fallback
    col_target = next((c for c, h in headers.items() if f'thang {month_num}' in h or f'thang {month_num:02d}' in h or f't{month_num}' in h), None)
    if not col_target:
        col_target = next((c for c, h in headers.items() if 'doanh thu/thang' in h or 'kpi thang' in h), 7)
        
    col_nt = next((c for c, h in headers.items() if 'nt thang' in h or 'nha thuoc thang' in h or 'target nt' in h), None)

    staff_list = []
    store_targets = {}
    current_store = ''
    
    for r in range(h_row + 1, ws.max_row + 1):
        st_v = ws.cell(r, col_store).value
        name_v = ws.cell(r, col_name).value
        role_v = ws.cell(r, col_role).value
        tb_v = ws.cell(r, col_tb).value
        gd_v = ws.cell(r, col_gd).value if col_gd else 0
        tgt_v = ws.cell(r, col_target).value if col_target else 0
        nt_v = ws.cell(r, col_nt).value if col_nt else 0
        
        if st_v:
            current_store = clean_branch_name(str(st_v))
            
        if nt_v and current_store:
            try:
                store_targets[current_store] = float(nt_v)
            except Exception:
                pass
                
        if name_v and current_store:
            s_name = str(name_v).strip()
            if s_name and s_name.lower() not in ('nhan vien', 'tong', 'none', 'tong cong', 'dược sĩ', 'nhân viên'):
                staff_list.append({
                    'store': current_store,
                    'name': s_name,
                    'role': str(role_v or 'NV').strip(),
                    'tb_bill': float(tb_v or 0) if isinstance(tb_v, (int, float)) else 0.0,
                    'giao_dich_ngay': float(gd_v or 0) if isinstance(gd_v, (int, float)) else 0.0,
                    'kpi_thang': float(tgt_v or 0) if isinstance(tgt_v, (int, float)) else 0.0,
                    'region': 'HCM'
                })
                
    wb.close()
    return staff_list, store_targets

# ==============================================================================
# 2. PARSE ĐỀ XUẤT HN (DYNAMIC HEADER PARSER)
# ==============================================================================
def parse_hn_proposal(hn_file_path, month_num=9):
    if not os.path.exists(hn_file_path):
        raise FileNotFoundError(f"Không tìm thấy file đề xuất HN: {hn_file_path}")
        
    wb = openpyxl.load_workbook(hn_file_path, data_only=True)
    
    target_sheet = None
    for name in wb.sheetnames:
        n_low = remove_accents(name).lower()
        if n_low.startswith(f't{month_num:02d}.') or n_low.startswith(f't{month_num}.') or n_low == f't{month_num}' or f'thang {month_num}' in n_low or f'thang {month_num:02d}' in n_low:
            target_sheet = name
            break
    if not target_sheet:
        for name in wb.sheetnames:
            n_low = remove_accents(name).lower()
            if n_low.startswith('t8') or n_low.startswith('t9') or n_low.startswith('t'):
                target_sheet = name
                break
    if not target_sheet:
        target_sheet = wb.sheetnames[-1]
        
    ws = wb[target_sheet]
    print(f"--> [HN Parser] Đọc Sheet: '{target_sheet}' (Tháng {month_num})")

    # Detect header row (row 1 to 4)
    h_row = 1
    for r in range(1, min(5, ws.max_row + 1)):
        row_str = " ".join([str(ws.cell(r, c).value or '') for c in range(1, 15)])
        row_str_norm = remove_accents(row_str).lower()
        if 'nhan vien' in row_str_norm or 'duoc si' in row_str_norm:
            h_row = r
            break

    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(h_row, c).value
        if v is not None and str(v).strip():
            headers[c] = remove_accents(str(v)).lower().strip()

    col_store = next((c for c, h in headers.items() if 'nha thuoc' in h or 'chi nhanh' in h), 1)
    col_name = next((c for c, h in headers.items() if 'nhan vien' in h or 'duoc si' in h), 2)
    col_role = next((c for c, h in headers.items() if 'chuc danh' in h or 'vi tri' in h), 3)
    col_tb = next((c for c, h in headers.items() if 'trung binh bill' in h or 'tb bill' in h or 'tbbill' in h), 4)
    col_ck = next((c for c, h in headers.items() if 'ck' in h or 'combo' in h or 'ny3' in h), 5)
    col_rev_ngay = next((c for c, h in headers.items() if 'doanh thu/ngay (100%)' in h or ('doanh thu/ngay' in h and 'thuc' not in h) or 'kpi doanh thu/ngay' in h), 7)
    
    col_target = next((c for c, h in headers.items() if f'thang {month_num}' in h or f'thang {month_num:02d}' in h or f't{month_num}' in h or f'kpis doanh thu/thang {month_num}' in h), None)
    if not col_target:
        col_target = next((c for c, h in headers.items() if 'doanh thu/thang' in h or 'kpi thang' in h), None)

    col_nt = next((c for c, h in headers.items() if f'nt thang {month_num}' in h or f'nt thang {month_num:02d}' in h or 'nt thang' in h or 'target nt' in h), None)

    days_in_month = calendar.monthrange(2026, month_num)[1]
    staff_list = []
    store_targets = {}
    current_store = ''
    
    for r in range(h_row + 1, ws.max_row + 1):
        st_v = ws.cell(r, col_store).value
        name_v = ws.cell(r, col_name).value
        role_v = ws.cell(r, col_role).value
        tb_v = ws.cell(r, col_tb).value
        ck_v = ws.cell(r, col_ck).value if col_ck else 0
        rev_ngay_v = ws.cell(r, col_rev_ngay).value if col_rev_ngay else 0
        tgt_v = ws.cell(r, col_target).value if col_target else 0
        nt_v = ws.cell(r, col_nt).value if col_nt else 0
        
        if st_v:
            current_store = clean_branch_name(str(st_v))
            
        if nt_v and current_store:
            try:
                store_targets[current_store] = float(nt_v)
            except Exception:
                pass
                
        if name_v and current_store:
            s_name = str(name_v).strip()
            if s_name and s_name.lower() not in ('nhan vien', 'tong', 'none', 'tong cong', 'dược sĩ', 'nhân viên'):
                rev_ngay = float(rev_ngay_v or 0) if isinstance(rev_ngay_v, (int, float)) else 0.0
                kpi_thang = float(tgt_v or 0) if isinstance(tgt_v, (int, float)) and float(tgt_v) > 0 else (rev_ngay * days_in_month if rev_ngay > 0 else 0.0)
                staff_list.append({
                    'store': current_store,
                    'name': s_name,
                    'role': str(role_v or 'BC').strip(),
                    'tb_bill': float(tb_v or 0) if isinstance(tb_v, (int, float)) else 0.0,
                    'kpi_ck_cb_ny3_ngay': float(ck_v or 0) if isinstance(ck_v, (int, float)) else 0.0,
                    'kpi_rev_ngay': rev_ngay,
                    'kpi_thang': kpi_thang,
                    'region': 'HN'
                })
                
    # Fallback store targets for Hanoi branches if not detected in columns
    if 'Hàng Bông' not in store_targets or store_targets['Hàng Bông'] <= 0:
        store_targets['Hàng Bông'] = 1350000000.0 if month_num == 9 else 1270000000.0
    if 'Đường Láng' not in store_targets or store_targets['Đường Láng'] <= 0:
        store_targets['Đường Láng'] = 1100000000.0 if month_num == 9 else 980000000.0
        
    wb.close()
    return staff_list, store_targets

# ==============================================================================
# 3. QUÉT THƯ MỤC DỰ ÁN
# ==============================================================================
def scan_project_catalog(project_folder):
    extracted_skus = {}
    if not project_folder or not os.path.exists(project_folder):
        return extracted_skus
        
    for root, dirs, files in os.walk(project_folder):
        for f in files:
            full_path = os.path.join(root, f)
            f_lower = f.lower()
            if f_lower.endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
                try:
                    wb = openpyxl.load_workbook(full_path, data_only=True)
                    for sname in wb.sheetnames:
                        ws = wb[sname]
                        grp = 'CK'
                        if 'kat' in f_lower: grp = 'Mini KAT'
                        elif 'ladycare' in f_lower: grp = 'Mini Ladycare'
                        elif 'party smart' in f_lower: grp = 'Mini PartySmart'
                        elif 'avc' in f_lower: grp = 'Mini AVC'
                        elif 'combo' in sname.lower(): grp = 'Combo'
                        elif 'ny3' in sname.lower() or 'hoahong' in sname.lower(): grp = 'NY3'
                        
                        for r in range(1, min(ws.max_row + 1, 300)):
                            row_vals = [ws.cell(r, c).value for c in range(1, min(ws.max_column + 1, 10))]
                            for idx, val in enumerate(row_vals):
                                val_str = str(val or '').strip()
                                if re.match(r'^SP\d{5,8}$', val_str, re.IGNORECASE):
                                    sku = val_str.upper()
                                    name = ''
                                    for nxt in range(idx + 1, min(len(row_vals), idx + 3)):
                                        if isinstance(row_vals[nxt], str) and len(row_vals[nxt]) > 3:
                                            name = row_vals[nxt].strip()
                                            break
                                    if sku not in extracted_skus:
                                        extracted_skus[sku] = {
                                            'sku': sku,
                                            'name': name or f"Sản phẩm {sku}",
                                            'group': grp,
                                            'source': f
                                        }
                    wb.close()
                except Exception:
                    pass
    return extracted_skus

# ==============================================================================
# 4. TRÌNH TẠO 4 SHEET ĐƯỢC TÔ MÀU VÀ ĐỊNH DẠNG ĐẸP MẮT
# ==============================================================================
def generate_kpisheet_package(
    hcm_file_path,
    hn_file_path,
    month_num=None,
    project_folder=None,
    extra_programs=None,
    output_filepath=None
):
    # Auto-detect month from files if not explicitly provided or if file content indicates specific month
    detected_month = detect_proposal_month(hcm_file_path, hn_file_path, month_num)
    if month_num is None:
        month_num = detected_month
    elif detected_month != month_num:
        print(f"--> [Month Alignment] File đầu vào là dữ liệu Tháng {detected_month} (thay vì Tháng {month_num}). Tự động đồng bộ sang Tháng {detected_month}!")
        month_num = detected_month

    period_str = f"2026-{month_num:02d}"
    if not output_filepath:
        month_out_dir = os.path.join(BASE_DIR, '..', f"thang{month_num}", "output")
        os.makedirs(month_out_dir, exist_ok=True)
        output_filepath = os.path.join(month_out_dir, f"NHÀ THUỐC THÁNG {month_num} 2026.xlsx")
        
    days_in_month = calendar.monthrange(2026, month_num)[1]
    start_date = datetime.datetime(2026, month_num, 1, 0, 0)
    end_date = datetime.datetime(2026, month_num, days_in_month, 0, 0)
    
    hcm_staff, hcm_stores = parse_hcm_proposal(hcm_file_path, month_num)
    hn_staff, hn_stores = parse_hn_proposal(hn_file_path, month_num)
    
    all_staff = hcm_staff + hn_staff
    all_stores = {**hcm_stores, **hn_stores}
    scanned_skus = scan_project_catalog(project_folder)
    
    default_store_targets = {
        'Trường Sa': 640000000,
        'Đỗ Quang Đẩu': 1100000000,
        'Minh Châu': 460000000,
        'Nam Hòa': 330000000,
        'Nguyễn Chí Thanh': 550000000,
        'Nguyễn Thị Thập': 545000000,
        'Nguyễn Văn Quá': 400000000,
        'Rạch Bùng Binh': 475000000,
        'Lê Bình': 550000000,
        'Đường Láng': 1100000000,
        'Hàng Bông': 1350000000
    }
    for st, tgt in default_store_targets.items():
        if st not in all_stores or all_stores[st] <= 0:
            all_stores[st] = tgt
            
    cht_map = {}
    for s in all_staff:
        if s['role'] in ('CHT', 'Q.CHT'):
            cht_map[s['store']] = s['name']
    default_chts = {
        'Trường Sa': 'Nguyễn Trần Ngọc Phương',
        'Đỗ Quang Đẩu': 'Ngô Thị Thanh Thắm',
        'Minh Châu': 'Thái Thùy Linh',
        'Nam Hòa': 'Vũ Ngọc Thanh Thảo',
        'Nguyễn Chí Thanh': 'Lê Thị Quỳnh Trâm',
        'Nguyễn Thị Thập': 'Triệu Thị Ngọc Lý',
        'Nguyễn Văn Quá': 'Võ Ngọc Giàu Sang',
        'Rạch Bùng Binh': 'Cù Thị Tường Vy',
        'Lê Bình': 'Hồ Thị Minh Hòa',
        'Đường Láng': 'Hứa Thị Kim Thoa',
        'Hàng Bông': 'Đinh Thị Lan Anh'
    }
    for st, cht in default_chts.items():
        if st not in cht_map:
            cht_map[st] = cht

    wb = openpyxl.Workbook()
    
    # ==============================================================================
    # SHEET 1: data (Tô màu phân biệt Kênh Offline, Kênh Online, Tổng DT)
    # ==============================================================================
    ws_data = wb.active
    ws_data.title = 'data'
    data_headers = [
        'Nhà thuốc', 'Tên nhân viên', 'Số giao dịch off', 'Doanh thu off',
        'Số giao dịch onl', 'Doanh thu onl', 'Tổng giao dịch', 'Doanh thu tổng'
    ]
    ws_data.append(data_headers)
    ws_data.cell(1, 10, value=end_date).number_format = 'dd/mm/yyyy'
    
    # Header colors for data:
    # A-B: Navy, C-D: Soft Blue (Off), E-F: Soft Orange (Onl), G-H: Soft Green (Total)
    for c_i in range(1, 9):
        cell = ws_data.cell(1, c_i)
        if c_i in (1, 2):
            cell.font = font_header_white
            cell.fill = fill_navy
        elif c_i in (3, 4):
            cell.font = font_header_dark
            cell.fill = fill_soft_blue
        elif c_i in (5, 6):
            cell.font = font_header_dark
            cell.fill = fill_orange
        elif c_i in (7, 8):
            cell.font = font_header_green
            cell.fill = fill_green
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_header
        
    seen_staff = set()
    current_data_row = 2
    for s in all_staff:
        key = (s['store'], s['name'])
        if key not in seen_staff:
            seen_staff.add(key)
            r = current_data_row
            ws_data.append([
                s['store'], s['name'], 0, 0, 0, 0,
                f'=C{r}+E{r}', f'=D{r}+F{r}'
            ])
            is_even = (r % 2 == 0)
            row_bg = fill_gray if is_even else None
            for c_i in range(1, 9):
                cell = ws_data.cell(r, c_i)
                cell.font = font_regular
                cell.border = border_thin
                if row_bg:
                    cell.fill = row_bg
                if c_i in (4, 6, 8):
                    cell.number_format = '#,##0'
                if c_i in (1, 2):
                    cell.alignment = Alignment(horizontal='left')
                elif c_i in (3, 5, 7):
                    cell.alignment = Alignment(horizontal='center')
                else:
                    cell.alignment = Alignment(horizontal='right')
            current_data_row += 1

    ws_data.cell(1, 10, value=end_date).number_format = 'dd/mm/yyyy'

    # ==============================================================================
    # SHEET 2: kpi dược sĩ (Tô màu trực quan theo cụm chỉ tiêu, hoàn thành, tiền thưởng)
    # ==============================================================================
    ws_ds = wb.create_sheet('kpi dược sĩ')
    
    # Row 1: Header Note & Date
    c_note = ws_ds.cell(1, 1, value='Dữ liệu cập nhật đến ngày')
    c_note.font = font_bold
    c_date = ws_ds.cell(1, 3, value='=data!J1')
    c_date.font = font_bold
    
    # Row 2: Headers
    ds_headers = [
        'Nhà thuốc', 'Nhân viên', 'Chức danh', 'KPI trung bình bill', 'Trung bình bill',
        'KPI doanh thu CK+Combo+NY3/ngày', 'CK + Combo + NY3/ngày', 'KPI Doanh thu/ngày (100%)',
        'Doanh thu thực/ngày', '% Hoàn thành KPI', 'Thưởng KPI', 'Hệ số * Dự án', 'Thưởng Dự án',
        'Thưởng Dự án + KPI + HH', 0.8, 0.9, 1.0, 0.8, 0.9, 1.0,
        'KPIs Doanh thu/ngày', 'KPIs Doanh thu/tháng', 'Số giao dịch off', 'Doanh thu off',
        'Trung bình bill off', 'Số giao dịch onl', 'Doanh thu onl', 'Trung bình bill onl',
        'Tổng giao dịch', 'Doanh thu tổng', 'Trung bình Bill'
    ]
    ws_ds.append(ds_headers)
    
    # Color palette for kpi dược sĩ headers:
    # 1-3 (A-C): Info -> Navy
    # 4-5 (D-E): TB Bill -> Light Blue
    # 6-7 (F-G): Dự án / CK -> Soft Purple
    # 8-10 (H-J): Doanh thu & % KPI -> Soft Yellow (Highlight)
    # 11-14 (K-N): THƯỞNG KPI & DỰ ÁN -> Emerald Green
    # 15-20 (O-T): Ngưỡng mức 1,2,3 -> Soft Gray
    # 21-22 (U-V): KPIs Target -> Soft Blue
    # 23-31 (W-AE): Thống kê bán hàng -> Soft Slate
    for c_i in range(1, 32):
        cell = ws_ds.cell(2, c_i)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_header
        
        if c_i in (1, 2, 3):
            cell.font = font_header_white
            cell.fill = fill_navy
        elif c_i in (4, 5):
            cell.font = font_header_dark
            cell.fill = fill_soft_blue
        elif c_i in (6, 7):
            cell.font = font_header_dark
            cell.fill = fill_purple
        elif c_i in (8, 9, 10):
            cell.font = font_header_dark
            cell.fill = fill_yellow
        elif c_i in (11, 12, 13, 14):
            cell.font = font_header_green
            cell.fill = fill_emerald
        elif c_i in range(15, 21):
            cell.font = font_header_dark
            cell.fill = fill_dark_gray
            cell.number_format = '0%'
        elif c_i in (21, 22):
            cell.font = font_header_dark
            cell.fill = fill_light_blue
        else:
            cell.font = font_header_dark
            cell.fill = fill_gray
            
    proj_sheet_name = f"Dự án T{month_num}"
    current_ds_row = 3
    for s in all_staff:
        r = current_ds_row
        st_name = s['store']
        emp_name = s['name']
        role = s['role']
        tb_bill = s['tb_bill'] or (245000 if role in ('BC', 'DSXC') else 150000)
        target_month = s['kpi_thang'] or 270000000
        
        f_hangdiem_formula = (
            f"=IFERROR(ROUND(_xludf.LET(nt,$A{r},nv,$B{r},songay,DAY(EOMONTH($C$1,0)),doanhthungay,$I{r},"
            f"duan_nv,SUMIFS('{proj_sheet_name}'!$D:$D,'{proj_sheet_name}'!$A:$A,nt,'{proj_sheet_name}'!$B:$B,nv),"
            f"duan_nt,SUMIFS('{proj_sheet_name}'!$D:$D,'{proj_sheet_name}'!$A:$A,nt),"
            f"tyle,duan_nv/duan_nt,"
            f"muc2_ngay,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$P:$P),"
            f"muc3_ngay,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$T:$T),"
            f"hangdiem1,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$I:$I),"
            f"hangdiem2,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$N:$N),"
            f"hangdiem3,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$R:$R),"
            f"_xludf.SWITCH(TRUE,doanhthungay>=muc3_ngay,hangdiem3,doanhthungay>=muc2_ngay,hangdiem2,TRUE,hangdiem1)/songay*tyle),-3),0)"
        )

        kpi_bonus_formula = (
            f'=_xlfn.XLOOKUP(B{r}, \'kpi nhà thuốc\'!$A:$A, \'kpi nhà thuốc\'!$G:$G, '
            f'IF(AND(OR(C{r}="DSXC", C{r}="DSCD", C{r}="DSBC", C{r}="DSTV", C{r}="Q.CHT", C{r}="BC"), G{r}>=F{r}), '
            f'AD{r} * _xlfn.IFS(AND(I{r}>=T{r}, E{r}>=D{r}), 0.015, I{r}>=T{r}, 0.013, AND(I{r}>=S{r}, E{r}>=D{r}), 0.013, I{r}>=S{r}, 0.011, AND(I{r}>=R{r}, E{r}>=D{r}), 0.012, I{r}>=R{r}, 0.010, TRUE, 0), 0))'
        )

        ws_ds.append([
            st_name,                                              # A: Nhà thuốc
            emp_name,                                             # B: Nhân viên
            role,                                                 # C: Chức danh
            tb_bill,                                              # D: KPI trung bình bill
            f"=AE{r}",                                            # E: Trung bình bill
            f_hangdiem_formula,                                   # F: KPI doanh thu CK+Combo+NY3/ngày
            f"=IFERROR(SUMIF('{proj_sheet_name}'!$B:$B, $B{r}, '{proj_sheet_name}'!$D:$D), 0)", # G: CK + Combo + NY3/ngày
            f"=Q{r}",                                             # H: KPI Doanh thu/ngày (100%)
            f"=(AD{r}/DAY($C$1))",                                # I: Doanh thu thực/ngày
            f"=_xlfn.XLOOKUP(B{r}, 'kpi nhà thuốc'!A:A, 'kpi nhà thuốc'!F:F, (AD{r}/DAY($C$1))/U{r})", # J: % Hoàn thành KPI
            kpi_bonus_formula,                                    # K: Thưởng KPI
            f'=IF(C{r}="DSTV","", IF(J{r}<0.6, 0.8, 1))',         # L: Hệ số * Dự án
            f"=IFERROR(SUMIF('{proj_sheet_name}'!$B:$B, $B{r}, '{proj_sheet_name}'!$M:$M)*$L{r}, 0)", # M: Thưởng Dự án
            f"=N(K{r})+M{r}",                                     # N: Thưởng Dự án + KPI + HH
            f"=CEILING(R{r}, 100000)",                            # O: Mức 1 (80%)
            f"=CEILING(S{r}, 100000)",                            # P: Mức 2 (90%)
            f"=CEILING(T{r}, 100000)",                            # Q: Mức 3 (100%)
            f"=U{r}*$R$2",                                        # R: 0.8 * ngày
            f"=U{r}*$S$2",                                        # S: 0.9 * ngày
            f"=U{r}*$T$2",                                        # T: 1.0 * ngày
            f"=V{r}/DAY($C$1)",                                   # U: KPIs Doanh thu/ngày
            target_month,                                         # V: KPIs Doanh thu/tháng
            f"=IFERROR(SUMIFS(data!$C:$C, data!$A:$A, $A{r}, data!$B:$B, $B{r}), 0)", # W: Số giao dịch off
            f"=IFERROR(SUMIF(data!$B:$B, $B{r}, data!$D:$D), 0)", # X: Doanh thu off
            f"=IF(W{r}>0, X{r}/W{r}, 0)",                         # Y: Trung bình bill off
            f"=IFERROR(SUMIFS(data!$E:$E, data!$A:$A, $A{r}, data!$B:$B, $B{r}), 0)", # Z: Số giao dịch onl
            f"=IFERROR(SUMIF(data!$B:$B, $B{r}, data!$F:$F), 0)", # AA: Doanh thu onl
            f"=IF(Z{r}>0, AA{r}/Z{r}, 0)",                        # AB: Trung bình bill onl
            f"=W{r}+Z{r}",                                        # AC: Tổng giao dịch
            f"=X{r}+AA{r}",                                       # AD: Doanh thu tổng
            f"=IF(AC{r}>0, AD{r}/AC{r}, 0)"                       # AE: Trung bình Bill
        ])
        
        is_even = (r % 2 == 0)
        row_tint = PatternFill(start_color='FAFAFA', end_color='FAFAFA', fill_type='solid') if is_even else None
        
        for c_i in range(1, 32):
            cell = ws_ds.cell(r, c_i)
            cell.font = font_regular
            cell.border = border_thin
            if row_tint:
                cell.fill = row_tint
                
            # Highlight special calculation columns
            if c_i == 10: # % KPI
                cell.fill = PatternFill(start_color='FFFBEA', end_color='FFFBEA', fill_type='solid')
                cell.font = font_bold
                cell.number_format = '0.0%'
            elif c_i == 14: # Thưởng Tổng
                cell.fill = PatternFill(start_color='EBF9F1', end_color='EBF9F1', fill_type='solid')
                cell.font = Font(name='Calibri', size=11, bold=True, color='15803D')
                cell.number_format = '#,##0'
            elif c_i in (4, 5, 8, 9, 11, 13, 15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 27, 28, 30, 31):
                cell.number_format = '#,##0'
                
            if c_i in (1, 2, 3):
                cell.alignment = Alignment(horizontal='left')
            elif c_i in (10, 12, 15, 16, 17, 18, 19, 20, 23, 26, 29):
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.alignment = Alignment(horizontal='right')
                
        current_ds_row += 1

    # ==============================================================================
    # SHEET 3: kpi nhà thuốc (Tô màu phân nhóm KPI CHT, Hàng điểm, Đạt KPI)
    # ==============================================================================
    ws_nt = wb.create_sheet('kpi nhà thuốc')
    
    # Row 1: Dates
    ws_nt.cell(1, 4, value='Dữ liệu cập nhật đến ngày').font = font_bold
    ws_nt.cell(1, 6, value="='kpi dược sĩ'!$C$1")
    ws_nt.cell(1, 20, value="='kpi dược sĩ'!$C$1")
    
    # Row 2: Headers
    nt_headers = [
        'Cửa hàng trưởng', 'Nhà thuốc', 'KPI doanh thu (tháng)', 'Doanh thu thực tế (tháng)',
        'Doanh thu thực tế (ngày)', 'Tỷ lệ đạt KPIs', 'Thưởng CHT', 'KPIs mức 1', 'KPIs hàng điểm mức 1',
        'KPIs mức 1/tháng', 'KPIs mức 1/ngày', 'KPIs mức 1/ngày', 'KPIs mức 2', 'KPIs hàng điểm mức 2',
        'KPIs mức 2/tháng', 'KPIs mức 2/ngày', 'KPIs mức 3', 'KPIs hàng điểm mức 3', 'KPIs mức 3/tháng',
        'KPIs mức 3/ngày', 'Doanh thu hàng điểm', 'Hàng điểm/ Doanh thu tổng', 'Xét Đạt KPI',
        'Số giao dịch off', 'Doanh thu off', 'Tbbill off', 'Số giao dịch onl', 'Doanh thu onl',
        'Tbbill onl', 'Tổng giao dịch', 'Doanh thu tổng', 'Trung bình Bill'
    ]
    ws_nt.append(nt_headers)
    
    # Header Colors for kpi nhà thuốc:
    # 1-2 (A-B): Info -> Navy
    # 3-6 (C-F): Target & Thực tế -> Soft Blue
    # 7 (G): THƯỞNG CHT -> Emerald Green
    # 8-20 (H-T): Mức 1, 2, 3 -> Soft Yellow / Dark Gray
    # 21-23 (U-W): Hàng điểm & XÉT ĐẠT -> Orange / Soft Green
    # 24-32 (X-AF): Thống kê off/onl -> Soft Gray
    for c_i in range(1, 33):
        cell = ws_nt.cell(2, c_i)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_header
        
        if c_i in (1, 2):
            cell.font = font_header_white
            cell.fill = fill_navy
        elif c_i in (3, 4, 5, 6):
            cell.font = font_header_dark
            cell.fill = fill_soft_blue
        elif c_i == 7:
            cell.font = font_header_green
            cell.fill = fill_emerald
        elif c_i in range(8, 21):
            cell.font = font_header_dark
            cell.fill = fill_yellow if c_i in (8, 13, 17) else fill_gray
        elif c_i in (21, 22):
            cell.font = font_header_dark
            cell.fill = fill_orange
        elif c_i == 23:
            cell.font = font_header_green
            cell.fill = fill_emerald
        else:
            cell.font = font_header_dark
            cell.fill = fill_gray
            
    current_nt_row = 3
    for st_name, tgt_thang in all_stores.items():
        r = current_nt_row
        cht_name = cht_map.get(st_name, '')
        
        cht_bonus_formula = (
            f'=IF(W{r}<>"Đạt", 0, IF(E{r}>=T{r}, D{r}*0.5%+1000000, IF(E{r}>=P{r}, D{r}*0.4%+700000, IF(E{r}>=K{r}, D{r}*0.35%+500000, 0))))'
        )
        
        ws_nt.append([
            cht_name,                                                     # A: CHT
            st_name,                                                      # B: Nhà thuốc
            tgt_thang,                                                    # C: KPI doanh thu (tháng)
            f"=AE{r}",                                                    # D: Doanh thu thực tế (tháng)
            f"=D{r}/DAY('kpi dược sĩ'!$C$1)",                             # E: Doanh thu thực tế (ngày)
            f"=E{r}/T{r}",                                                # F: Tỷ lệ đạt KPIs
            cht_bonus_formula,                                            # G: Thưởng CHT
            0.8,                                                          # H: KPIs mức 1
            f"=J{r}*0.25",                                                # I: KPIs hàng điểm mức 1
            f"=C{r}*H{r}",                                                # J: KPIs mức 1/tháng
            f"=T{r}*H{r}",                                                # K: KPIs mức 1/ngày
            f"=CEILING(K{r}, 100000)",                                    # L: KPIs mức 1/ngày (làm tròn)
            0.9,                                                          # M: KPIs mức 2
            f"=O{r}*0.25",                                                # N: KPIs hàng điểm mức 2
            f"=C{r}*M{r}",                                                # O: KPIs mức 2/tháng
            f"=T{r}*M{r}",                                                # P: KPIs mức 2/ngày
            1.0,                                                          # Q: KPIs mức 3
            f"=S{r}*0.25",                                                # R: KPIs hàng điểm mức 3
            f"=C{r}*Q{r}",                                                # S: KPIs mức 3/tháng
            f"=C{r}/DAY($T$1)",                                           # T: KPIs mức 3/ngày
            f"=SUMIF('{proj_sheet_name}'!$A:$A, B{r}, '{proj_sheet_name}'!D:D)*DAY($F$1)", # U: Doanh thu hàng điểm
            f"=U{r}/D{r}",                                                # V: Hàng điểm/ Doanh thu tổng
            f'=IF(OR(F{r}="", V{r}=""), "", IF(AND(F{r}>=80%, V{r}>=25%), "Đạt", "Không đạt"))', # W: Xét Đạt KPI
            f"=SUMIF(data!$A:$A, B{r}, data!$C:$C)",                      # X: Số giao dịch off
            f"=SUMIF(data!$A:$A, B{r}, data!$D:$D)",                      # Y: Doanh thu off
            f"=IF(X{r}>0, Y{r}/X{r}, 0)",                                 # Z: Tbbill off
            f"=SUMIF(data!$A:$A, B{r}, data!$E:$E)",                      # AA: Số giao dịch onl
            f"=SUMIF(data!$A:$A, B{r}, data!$F:$F)",                      # AB: Doanh thu onl
            f"=IF(AA{r}>0, AB{r}/AA{r}, 0)",                              # AC: Tbbill onl
            f"=X{r}+AA{r}",                                               # AD: Tổng giao dịch
            f"=Y{r}+AB{r}",                                               # AE: Doanh thu tổng
            f"=IF(AD{r}>0, AE{r}/AD{r}, 0)"                               # AF: Trung bình Bill
        ])
        
        is_even = (r % 2 == 0)
        row_tint = PatternFill(start_color='FAFAFA', end_color='FAFAFA', fill_type='solid') if is_even else None
        
        for c_i in range(1, 33):
            cell = ws_nt.cell(r, c_i)
            cell.font = font_regular
            cell.border = border_thin
            if row_tint:
                cell.fill = row_tint
                
            if c_i == 7: # Thưởng CHT
                cell.fill = PatternFill(start_color='EBF9F1', end_color='EBF9F1', fill_type='solid')
                cell.font = Font(name='Calibri', size=11, bold=True, color='15803D')
                cell.number_format = '#,##0'
            elif c_i == 23: # Xét Đạt
                cell.font = font_bold
                cell.alignment = Alignment(horizontal='center')
            elif c_i in (3, 4, 5, 9, 10, 11, 12, 14, 15, 16, 18, 19, 20, 21, 25, 26, 28, 29, 31, 32):
                cell.number_format = '#,##0'
            elif c_i in (6, 8, 13, 17, 22):
                cell.number_format = '0.0%'
                
            if c_i in (1, 2):
                cell.alignment = Alignment(horizontal='left')
            elif c_i in (6, 8, 13, 17, 22, 23, 24, 27, 30):
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.alignment = Alignment(horizontal='right')
                
        current_nt_row += 1

    # ==============================================================================
    # SHEET 4: Dự án T{month} (Tô màu các cột Sản Phẩm, Thưởng Tier & Total)
    # ==============================================================================
    ws_proj = wb.create_sheet(proj_sheet_name)
    
    # Row 1: Date
    ws_proj.cell(1, 1, value="='kpi dược sĩ'!$C$1")
    ws_proj.cell(1, 1).font = font_bold
    
    # Row 2: Headers
    proj_headers = [
        'Nhà thuốc', 'Nhân viên', 'Chức danh', 'CK + Combo + NY3/ngày', 'NY3',
        'Chiết khấu', 'Combo Liều', 'Chiết khấu/ngày', 'Combo Liều/ngày',
        'Thưởng dự án', 'Thưởng thêm', 'Thưởng dự án Hot Bill HN', 'Total dự án'
    ]
    ws_proj.append(proj_headers)
    
    # Header Colors for Dự án:
    # 1-3 (A-C): Info -> Navy
    # 4-9 (D-I): Sản phẩm & Doanh số -> Soft Purple / Soft Blue
    # 10-13 (J-M): THƯỞNG DỰ ÁN & TOTAL -> Emerald Green
    for c_i in range(1, 14):
        cell = ws_proj.cell(2, c_i)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_header
        
        if c_i in (1, 2, 3):
            cell.font = font_header_white
            cell.fill = fill_navy
        elif c_i in (4, 5, 6, 7, 8, 9):
            cell.font = font_header_dark
            cell.fill = fill_purple
        elif c_i in (10, 11, 12, 13):
            cell.font = font_header_green
            cell.fill = fill_emerald
            
    current_proj_row = 3
    for s in all_staff:
        r = current_proj_row
        st_name = s['store']
        emp_name = s['name']
        role = s['role']
        
        tier_proj_formula = (
            f'=IF(OR($A{r}="Hàng Bông",$A{r}="Đường Láng"),'
            f'IF(AND($H{r}>=3400000,$I{r}>=800000),4000000,'
            f'IF(AND($H{r}>=3000000,$I{r}>=700000),2800000,'
            f'IF(AND($H{r}>=2600000,$I{r}>=550000),2000000,'
            f'IF(AND($H{r}>=2200000,$I{r}>=450000),1200000,0)))),'
            f'IF(AND($H{r}>=1500000,$I{r}>=700000),4000000,'
            f'IF(AND($H{r}>=1200000,$I{r}>=650000),2800000,'
            f'IF(AND($H{r}>=1000000,$I{r}>=580000),2200000,'
            f'IF(AND($H{r}>=850000,$I{r}>=480000),1600000,'
            f'IF(AND($H{r}>=630000,$I{r}>=350000),1000000,0))))))'
        )
        
        k_bonus_formula = (
            f'=IF(AND($J{r}=0, IF(OR($A{r}="Hàng Bông",$A{r}="Đường Láng"),$D{r}>=2000000,$D{r}>=950000)), 500000, 0)'
        )
        
        hot_bill_formula = f"=IFERROR(SUMIF('Hot Bill HN'!$D:$D, $B{r}, 'Hot Bill HN'!$G:$G), 0)"
        
        ws_proj.append([
            st_name,                                  # A: Nhà thuốc
            emp_name,                                 # B: Nhân viên
            role,                                     # C: Chức danh
            f"=SUM(E{r}:G{r})/DAY($A$1)",             # D: CK + Combo + NY3/ngày
            0,                                        # E: NY3
            0,                                        # F: Chiết khấu
            0,                                        # G: Combo Liều
            f"=F{r}/DAY($A$1)",                       # H: Chiết khấu/ngày
            f"=G{r}/DAY($A$1)",                       # I: Combo Liều/ngày
            tier_proj_formula,                        # J: Thưởng dự án
            k_bonus_formula,                          # K: Thưởng thêm
            hot_bill_formula,                         # L: Thưởng dự án Hot Bill HN
            f"=$J{r}+$K{r}+$L{r}"                     # M: Total dự án
        ])
        
        is_even = (r % 2 == 0)
        row_tint = PatternFill(start_color='FAFAFA', end_color='FAFAFA', fill_type='solid') if is_even else None
        
        for c_i in range(1, 14):
            cell = ws_proj.cell(r, c_i)
            cell.font = font_regular
            cell.border = border_thin
            if row_tint:
                cell.fill = row_tint
                
            if c_i == 13: # Total dự án
                cell.fill = PatternFill(start_color='EBF9F1', end_color='EBF9F1', fill_type='solid')
                cell.font = Font(name='Calibri', size=11, bold=True, color='15803D')
                cell.number_format = '#,##0'
            elif c_i in (4, 5, 6, 7, 8, 9, 10, 11, 12):
                cell.number_format = '#,##0'
                
            if c_i in (1, 2, 3):
                cell.alignment = Alignment(horizontal='left')
            else:
                cell.alignment = Alignment(horizontal='right')
                
        current_proj_row += 1

    # ==============================================================================
    # SHEET 5: Hot Bill HN (Tô màu bảng hóa đơn thưởng nóng)
    # ==============================================================================
    ws_hot = wb.create_sheet('Hot Bill HN')
    hot_headers = [
        'Mã hóa đơn', 'Thời gian', 'Chi nhánh', 'Tên Dược Sĩ',
        'Số lượng SP CK', 'Doanh thu nhóm CK+CKHN', 'Tiền Thưởng (VNĐ)', 'Chi tiết sản phẩm'
    ]
    ws_hot.append(hot_headers)
    for c_i in range(1, 9):
        cell = ws_hot.cell(1, c_i)
        if c_i in (1, 2, 3, 4):
            cell.font = font_header_white
            cell.fill = fill_navy
        elif c_i in (5, 6):
            cell.font = font_header_dark
            cell.fill = fill_soft_blue
        elif c_i == 7:
            cell.font = font_header_green
            cell.fill = fill_emerald
        else:
            cell.font = font_header_dark
            cell.fill = fill_gray
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_header

    # Apply full unified KPI styling (Times New Roman, Pastel Fills, Column Widths, Freeze Panes)
    try:
        apply_full_kpi_styling(wb)
    except Exception as e:
        print(f"⚠️ [Styling Warning] {e}")

    # Save output strictly to target file only
    os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
    try:
        wb.save(output_filepath)
        print(f"--> [Output] Đã lưu file duy nhất vào thư mục tháng: {output_filepath}")
    except PermissionError:
        base, ext = os.path.splitext(output_filepath)
        fallback_path = f"{base}_new{ext}"
        wb.save(fallback_path)
        print(f"⚠️ File đang mở trong Excel, đã lưu thành công vào: {fallback_path}")
        output_filepath = fallback_path

    wb.close()
    
    return {
        'filepath': output_filepath,
        'filename': os.path.basename(output_filepath),
        'period': period_str,
        'sheets': wb.sheetnames,
        'stores_count': len(all_stores),
        'staff_count': len(all_staff),
        'hcm_staff_count': len(hcm_staff),
        'hn_staff_count': len(hn_staff),
        'stores': list(all_stores.keys()),
        'staff_sample': all_staff[:10]
    }

if __name__ == '__main__':
    hcm_p = os.path.join(BASE_DIR, 'KPI CNT HCM Tháng 09.xlsx')
    hn_p = os.path.join(BASE_DIR, 'e_xuat_KPI_Quy_3.26.xlsx')
    
    if os.path.exists(hcm_p) and os.path.exists(hn_p):
        print("Testing generate_kpisheet_package with professional colors...")
        res = generate_kpisheet_package(hcm_p, hn_p, month_num=9)
        print("Result:", json.dumps(res, ensure_ascii=False, indent=2))
