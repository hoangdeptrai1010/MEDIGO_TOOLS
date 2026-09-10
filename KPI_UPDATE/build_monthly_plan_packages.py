import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import json
import sys
import io
import re
import glob
import calendar

try:
    import docx
except ImportError:
    docx = None

try:
    import pptx
except ImportError:
    pptx = None

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PLANS_DIR = os.path.join(BASE_DIR, 'plans')
os.makedirs(PLANS_DIR, exist_ok=True)

CONFIG_PATH = os.path.join(BASE_DIR, 'monthly_config.json')

def load_default_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "branches": [],
        "monthly_policies": {}
    }

# Styling
header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

def style_sheet(ws):
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            if isinstance(cell.value, (int, float)):
                if abs(cell.value) > 1000:
                    cell.number_format = '#,##0'

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

def scan_folder_for_project_data(folder_path):
    """
    Quét đệ quy thư mục dự án tháng để tìm tất cả các file .xlsx, .docx, .pptx
    và trích xuất danh sách SKUs, mini projects, chính sách thưởng.
    """
    extracted = {
        'skus': {},           # sku -> {'name': ..., 'group': ..., 'price': ..., 'discount': ..., 'note': ...}
        'mini_projects': [],  # list of mini project metadata
        'docx_notes': [],     # extracted text from docx
        'pptx_rules': [],     # extracted text/tables from pptx
        'found_files': []
    }
    
    if not os.path.exists(folder_path):
        return extracted

    for root, dirs, files in os.walk(folder_path):
        for f in files:
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, folder_path)
            f_lower = f.lower()
            
            # 1. File Excel (.xlsx, .xls)
            if f_lower.endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
                extracted['found_files'].append({'type': 'excel', 'path': rel_path})
                try:
                    wb = openpyxl.load_workbook(full_path, data_only=True)
                    for sname in wb.sheetnames:
                        ws = wb[sname]
                        s_lower = sname.lower()
                        
                        # Xác định nhóm mặc định dựa trên tên file hoặc tên sheet
                        default_group = 'CK'
                        if 'kat' in f_lower:
                            default_group = 'Mini KAT'
                        elif 'ladycare' in f_lower:
                            default_group = 'Mini Ladycare'
                        elif 'party smart' in f_lower or 'partysmart' in f_lower:
                            default_group = 'Mini PartySmart'
                        elif 'avc' in f_lower:
                            default_group = 'Mini AVC'
                        elif 'combo' in s_lower or 'combo' in f_lower:
                            default_group = 'Combo'
                        elif 'hoahong' in s_lower or 'ny3' in s_lower:
                            default_group = 'NY3'
                        elif 'hn' in f_lower:
                            default_group = 'CK-HN'
                        elif 'hcm' in f_lower:
                            default_group = 'CK-HCM'

                        # Quét các dòng
                        for r in range(1, min(ws.max_row + 1, 500)):
                            row_vals = [ws.cell(r, c).value for c in range(1, min(ws.max_column + 1, 15))]
                            
                            # Tìm SKU
                            for idx, val in enumerate(row_vals):
                                val_str = str(val or '').strip()
                                if re.match(r'^SP\d{5,8}$', val_str, re.IGNORECASE):
                                    sku = val_str.upper()
                                    name = ''
                                    price = 0
                                    discount = 0
                                    for next_idx in range(idx + 1, min(len(row_vals), idx + 4)):
                                        nxt_val = row_vals[next_idx]
                                        if isinstance(nxt_val, str) and len(nxt_val) > 3 and not name:
                                            name = nxt_val.strip()
                                        elif isinstance(nxt_val, (int, float)) and nxt_val > 0:
                                            if nxt_val > 1000 and price == 0:
                                                price = float(nxt_val)
                                            elif discount == 0:
                                                discount = float(nxt_val)
                                                
                                    if sku not in extracted['skus']:
                                        extracted['skus'][sku] = {
                                            'sku': sku,
                                            'name': name or f"Sản phẩm {sku}",
                                            'group': default_group,
                                            'price': price,
                                            'discount': discount,
                                            'source': f
                                        }
                    wb.close()
                except Exception as e:
                    print(f"Lỗi đọc file Excel {f}: {e}")

            # 2. File Word (.docx)
            elif f_lower.endswith('.docx') and not f.startswith('~$') and docx:
                extracted['found_files'].append({'type': 'docx', 'path': rel_path})
                try:
                    doc = docx.Document(full_path)
                    for p in doc.paragraphs:
                        text = p.text.strip()
                        if text and len(text) > 10:
                            extracted['docx_notes'].append(text)
                except Exception as e:
                    print(f"Lỗi đọc file Word {f}: {e}")

            # 3. File PowerPoint (.pptx)
            elif f_lower.endswith('.pptx') and not f.startswith('~$') and pptx:
                extracted['found_files'].append({'type': 'pptx', 'path': rel_path})
                try:
                    prs = pptx.Presentation(full_path)
                    for slide_idx, slide in enumerate(prs.slides, 1):
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for p in shape.text_frame.paragraphs:
                                    t = p.text.strip()
                                    if t and len(t) > 5:
                                        extracted['pptx_rules'].append(f"[Slide {slide_idx}] {t}")
                            elif shape.has_table:
                                table_rows = []
                                for row in shape.table.rows:
                                    r_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                                    if r_text:
                                        table_rows.append(" | ".join(r_text))
                                if table_rows:
                                    extracted['pptx_rules'].append(f"[Slide {slide_idx} Table] " + " // ".join(table_rows))
                except Exception as e:
                    print(f"Lỗi đọc file PPTX {f}: {e}")

    return extracted

# ==============================================================================
# BỘ TRÍCH XUẤT ĐỀ XUẤT KPI TỪ 2 QUẢN LÝ KHU VỰC (HCM & HN)
# ==============================================================================
def parse_hcm_kpi_proposal(hcm_file_path, month_num=9):
    """
    Trích xuất danh sách nhân sự, chỉ tiêu doanh thu, bill, số đơn và KPI nhà thuốc từ file QL HCM.
    """
    if not os.path.exists(hcm_file_path):
        raise FileNotFoundError(f"Không tìm thấy file KPI HCM tại: {hcm_file_path}")
        
    wb = openpyxl.load_workbook(hcm_file_path, data_only=True)
    
    # Tìm sheet chỉ tiêu tháng
    target_sheet = None
    for name in wb.sheetnames:
        n_lower = name.lower()
        if 'kpi' in n_lower and (f'tháng {month_num:02d}' in n_lower or f'tháng {month_num}' in n_lower or f't{month_num}' in n_lower):
            target_sheet = name
            break
    if not target_sheet:
        for name in wb.sheetnames:
            if 'kpi' in name.lower():
                target_sheet = name
                break
    if not target_sheet:
        target_sheet = wb.sheetnames[-1]
        
    ws = wb[target_sheet]
    print(f"--> [HCM Proposal] Đang nạp từ Sheet: [{target_sheet}]")
    
    staff_list = []
    store_targets = {}
    current_store = ''
    
    for r in range(2, ws.max_row + 1):
        store_val = ws.cell(r, 1).value
        staff_val = ws.cell(r, 2).value
        role_val = ws.cell(r, 3).value
        tb_bill = ws.cell(r, 4).value
        giao_dich_ngay = ws.cell(r, 5).value
        kpi_thang = ws.cell(r, 7).value or ws.cell(r, 6).value
        nt_thang = ws.cell(r, 8).value
        
        if store_val:
            current_store = clean_branch_name(str(store_val))
            
        if nt_thang and current_store:
            try:
                store_targets[current_store] = float(nt_thang)
            except Exception:
                pass
                
        if staff_val and current_store:
            s_name = str(staff_val).strip()
            if s_name and s_name.lower() not in ('nhân viên', 'tổng', 'none'):
                staff_list.append({
                    'store': current_store,
                    'name': s_name,
                    'role': str(role_val or 'NV').strip(),
                    'tb_bill': float(tb_bill or 0),
                    'giao_dich_ngay': float(giao_dich_ngay or 0),
                    'kpi_thang': float(kpi_thang or 0),
                    'region': 'HCM'
                })
                
    wb.close()
    return staff_list, store_targets

def parse_hn_kpi_proposal(hn_file_path, month_num=9):
    """
    Trích xuất danh sách nhân sự, chỉ tiêu doanh thu, bill, CK/ngày và KPI nhà thuốc từ file QL HN.
    """
    if not os.path.exists(hn_file_path):
        raise FileNotFoundError(f"Không tìm thấy file KPI HN tại: {hn_file_path}")
        
    wb = openpyxl.load_workbook(hn_file_path, data_only=True)
    
    # Tìm sheet chỉ tiêu tháng
    target_sheet = None
    for name in wb.sheetnames:
        n_lower = name.lower()
        if n_lower.startswith(f't{month_num:02d}.') or n_lower.startswith(f't{month_num}.') or n_lower == f't{month_num}':
            target_sheet = name
            break
    if not target_sheet:
        for name in wb.sheetnames:
            if name.lower().startswith('t8') or name.lower().startswith('t9'):
                target_sheet = name
                break
    if not target_sheet:
        target_sheet = wb.sheetnames[-1]
        
    ws = wb[target_sheet]
    print(f"--> [HN Proposal] Đang nạp từ Sheet: [{target_sheet}]")
    
    days_in_month = calendar.monthrange(2026, month_num)[1]
    staff_list = []
    store_targets = {}
    
    for r in range(2, ws.max_row + 1):
        store_val = ws.cell(r, 1).value
        staff_val = ws.cell(r, 2).value
        role_val = ws.cell(r, 3).value
        tb_bill = ws.cell(r, 4).value
        kpi_ck_cb_ny3_ngay = ws.cell(r, 5).value
        kpi_rev_ngay = ws.cell(r, 7).value
        
        if store_val and staff_val:
            s_name = str(staff_val).strip()
            s_store = clean_branch_name(str(store_val))
            if s_name and s_name.lower() not in ('nhân viên', 'tổng', 'none'):
                rev_ngay = float(kpi_rev_ngay or 0)
                kpi_thang = rev_ngay * days_in_month if rev_ngay > 0 else 0
                staff_list.append({
                    'store': s_store,
                    'name': s_name,
                    'role': str(role_val or 'BC').strip(),
                    'tb_bill': float(tb_bill or 0),
                    'kpi_ck_cb_ny3_ngay': float(kpi_ck_cb_ny3_ngay or 0),
                    'kpi_rev_ngay': rev_ngay,
                    'kpi_thang': kpi_thang,
                    'region': 'HN'
                })
                
    # Lấy thêm target nhà thuốc từ sheet KPI nếu có
    if 'KPI' in wb.sheetnames:
        ws_kpi = wb['KPI']
        # Dò dòng tháng month_num
        for r in range(3, 16):
            if ws_kpi.cell(r, 1).value == month_num:
                hb_target = ws_kpi.cell(r, 3).value or ws_kpi.cell(r, 2).value
                dl_target = ws_kpi.cell(r, 8).value or ws_kpi.cell(r, 7).value
                if hb_target and float(hb_target) > 0:
                    store_targets['Hàng Bông'] = float(hb_target)
                if dl_target and float(dl_target) > 0:
                    store_targets['Đường Láng'] = float(dl_target)
                break
                
    if 'Hàng Bông' not in store_targets:
        store_targets['Hàng Bông'] = 1350000000.0
    if 'Đường Láng' not in store_targets:
        store_targets['Đường Láng'] = 1100000000.0
        
    wb.close()
    return staff_list, store_targets

def build_plan_from_manager_proposals(
    hcm_proposal_path,
    hn_proposal_path,
    month_num=9,
    project_folder=None,
    extra_programs=None,
    output_filepath=None
):
    """
    Xây dựng gói kế hoạch hoàn chỉnh KeHoachKPI_YYYY-MM.xlsx từ 2 file đề xuất của 2 QLKV.
    """
    period_str = f"2026-{month_num:02d}"
    days_in_month = calendar.monthrange(2026, month_num)[1]
    
    if not output_filepath:
        output_filepath = os.path.join(PLANS_DIR, f'KeHoachKPI_{period_str}.xlsx')
        
    print(f"\n=======================================================")
    print(f"🚀 [PlanBuilder] Khởi tạo Gói Kế Hoạch Tháng {month_num}/2026 từ 2 file QLKV")
    print(f"=======================================================")
    
    hcm_staff, hcm_stores = parse_hcm_kpi_proposal(hcm_proposal_path, month_num)
    hn_staff, hn_stores = parse_hn_kpi_proposal(hn_proposal_path, month_num)
    
    all_staff = hcm_staff + hn_staff
    all_stores = {**hcm_stores, **hn_stores}
    
    # Quét dữ liệu dự án nếu có
    scanned_data = {'skus': {}, 'mini_projects': [], 'docx_notes': [], 'pptx_rules': []}
    if project_folder and os.path.exists(project_folder):
        print(f"--> [PlanBuilder] Đang quét thư mục dự án: {project_folder}")
        scanned_data = scan_folder_for_project_data(project_folder)
        
    wb = openpyxl.Workbook()
    
    # 1. SHEET: NHÂN SỰ
    ws_staff = wb.active
    ws_staff.title = 'Nhân sự'
    ws_staff.append(['Mã NV', 'Tên Nhân Viên', 'Chức Danh', 'Mã Nhà Thuốc', 'Tên Nhà Thuốc', 'Khu Vực', 'Trạng Thái'])
    
    seen_staff = set()
    for idx, s in enumerate(all_staff, 1):
        s_name = s['name']
        s_cn = s['store']
        key = (s_cn, s_name)
        if key not in seen_staff:
            seen_staff.add(key)
            s_id = f"NV{idx:03d}"
            st_id = f"NT_{s_cn.replace(' ', '')}"
            ws_staff.append([s_id, s_name, s['role'], st_id, s_cn, s['region'], 'Đang làm việc'])
    style_sheet(ws_staff)
    
    # 2. SHEET: KPI NHÀ THUỐC
    ws_stores = wb.create_sheet('KPI nhà thuốc')
    ws_stores.append(['Mã Nhà Thuốc', 'Tên Nhà Thuốc', 'Khu Vực', 'KPI Doanh Thu Tháng', 'Tên CHT', 'Mốc Hàng Điểm Mức 1', 'Mốc Hàng Điểm Mức 2', 'Mốc Hàng Điểm Mức 3'])
    
    # Xác định CHT cho từng nhà thuốc
    cht_map = {}
    for s in all_staff:
        if s['role'] in ('CHT', 'Q.CHT'):
            cht_map[s['store']] = s['name']
            
    for store_name, tgt in all_stores.items():
        kv = 'HN' if store_name in ['Hàng Bông', 'Đường Láng'] else 'HCM'
        st_id = f"NT_{store_name.replace(' ', '')}"
        cht_name = cht_map.get(store_name, '')
        m1 = tgt * 0.8
        m2 = tgt * 0.9
        m3 = tgt * 1.0
        ws_stores.append([st_id, store_name, kv, tgt, cht_name, m1, m2, m3])
    style_sheet(ws_stores)
    
    # 3. SHEET: KPI NHÂN VIÊN
    ws_emp = wb.create_sheet('KPI nhân viên')
    ws_emp.append(['Mã Chức Danh', 'Chức Danh', 'KPI Doanh Thu Tháng', 'KPI Trung Bình Bill', 'Thưởng Đạt Bậc 1 (%)', 'Thưởng Đạt Bậc 2 (%)', 'Thưởng Đạt Bậc 3 (%)'])
    roles_data = [
        ('CD_CHT', 'CHT', 270000000, 150000, 1.2, 1.3, 1.5),
        ('CD_DSXC', 'DSXC', 270000000, 245000, 1.2, 1.3, 1.5),
        ('CD_BC', 'BC', 270000000, 245000, 1.2, 1.3, 1.5),
        ('CD_NV', 'NV', 270000000, 200000, 1.0, 1.1, 1.3),
        ('CD_DSTV', 'DSTV', 270000000, 150000, 0.8, 1.0, 1.2)
    ]
    for r in roles_data:
        ws_emp.append(list(r))
    style_sheet(ws_emp)
    
    # 4. SHEET: DANH MỤC DỰ ÁN
    ws_proj = wb.create_sheet('Danh mục dự án')
    ws_proj.append(['Mã Hàng (SKU)', 'Tên Sản Phẩm', 'Nhóm Dự Án', 'Trọng Số / Đơn Giá Thưởng', 'Ghi Chú & Nguồn'])
    
    seen_skus = set()
    for sku, item in scanned_data['skus'].items():
        if sku not in seen_skus:
            seen_skus.add(sku)
            ws_proj.append([sku, item.get('name', ''), item.get('group', 'CK'), item.get('discount', 1.0), f"Nguồn: {item.get('source', '')}"])
            
    # Bổ sung danh mục CK chuẩn
    ck_files_fallback = [
        os.path.join(BASE_DIR, 'thang8', 'Pharmacy_retail_Store_KPIs_August 2026', 'Hà Nội', 'Dự án', 'Chương trình', 'CK-HN.xlsx'),
        os.path.join(BASE_DIR, 'thang8', 'Pharmacy_retail_Store_KPIs_August 2026', 'Hồ Chí Minh', 'Dự án', 'Chương trình', 'CK-HCM.xlsx')
    ]
    for ck_f in ck_files_fallback:
        if os.path.exists(ck_f):
            try:
                wb_ck = openpyxl.load_workbook(ck_f, data_only=True)
                for sname in wb_ck.sheetnames:
                    ws_ck = wb_ck[sname]
                    s_lower = sname.lower()
                    grp = 'CK' if 'ck' in s_lower or 'chietkhau' in s_lower else ('Combo' if 'combo' in s_lower else ('NY3' if 'hoahong' in s_lower else None))
                    if not grp: continue
                    for r in range(2, ws_ck.max_row+1):
                        sku = str(ws_ck.cell(r, 2).value or '').strip()
                        name = str(ws_ck.cell(r, 3).value or '').strip()
                        if sku.startswith('SP') and sku not in seen_skus:
                            seen_skus.add(sku)
                            ws_proj.append([sku, name, grp, 1.0, f"Hàng điểm {grp}"])
                wb_ck.close()
            except Exception:
                pass
    style_sheet(ws_proj)
    
    # 5. SHEET: QUY TẮC THƯỞNG
    ws_rules = wb.create_sheet('Quy tắc thưởng')
    ws_rules.append(['Kỳ', 'Khu Vực', 'Tier', 'Mức CK/Ngày (Min)', 'Mức Combo/Ngày (Min)', 'Tiền Thưởng (VNĐ)', 'Khuyến Khích Min/Ngày', 'Thưởng Khuyến Khích'])
    
    # Hanoi Tiers
    hn_tiers = [
        {"min_ck": 2200000, "min_combo": 450000, "bonus": 1000000},
        {"min_ck": 2500000, "min_combo": 520000, "bonus": 1600000},
        {"min_ck": 2800000, "min_combo": 650000, "bonus": 2200000},
        {"min_ck": 3200000, "min_combo": 750000, "bonus": 2800000},
        {"min_ck": 3400000, "min_combo": 900000, "bonus": 4000000}
    ]
    for idx, t in enumerate(hn_tiers, 1):
        ws_rules.append([period_str, 'HN', f'Tier {idx}', t['min_ck'], t['min_combo'], t['bonus'], 2000000, 500000])
        
    # HCM Tiers
    hcm_tiers = [
        {"min_ck": 630000, "min_combo": 350000, "bonus": 1000000},
        {"min_ck": 850000, "min_combo": 480000, "bonus": 1600000},
        {"min_ck": 1000000, "min_combo": 580000, "bonus": 2200000},
        {"min_ck": 1200000, "min_combo": 650000, "bonus": 2800000},
        {"min_ck": 1500000, "min_combo": 700000, "bonus": 4000000}
    ]
    for idx, t in enumerate(hcm_tiers, 1):
        ws_rules.append([period_str, 'HCM', f'Tier {idx}', t['min_ck'], t['min_combo'], t['bonus'], 950000, 500000])
    style_sheet(ws_rules)
    
    # 6. SHEET: CHƯƠNG TRÌNH ĐẶC BIỆT
    ws_special = wb.create_sheet('Chương trình đặc biệt')
    ws_special.append(['Tên Chương Trình', 'Khu Vực Áp Dụng', 'Điều Kiện Áp Dụng', 'Mức Thưởng', 'Ghi Chú & Cách Tính'])
    ws_special.append(['Hot Bill HN', 'HN (Hàng Bông, Đường Láng)', 'Hóa đơn chứa hàng điểm CK >= 1,000,000đ', 50000, 'Thưởng trực tiếp theo hóa đơn'])
    ws_special.append(['Mini KAT', 'Toàn hệ thống', 'Doanh số dự án Mini KAT đạt target', 0, 'Cộng vào bảng lương'])
    ws_special.append(['Mini Ladycare', 'Toàn hệ thống', 'Doanh số dự án Mini Ladycare', 0, 'Cộng vào bảng lương'])
    ws_special.append(['Mini PartySmart', 'Toàn hệ thống', 'Doanh số dự án Mini Party Smart', 0, 'Cộng vào bảng lương'])
    ws_special.append(['Mini AVC', 'HN (Hàng Bông, Đường Láng)', 'Doanh số dự án Mini AVC', 0, 'Cộng vào bảng lương'])
    ws_special.append(['WhatsApp', 'Toàn hệ thống', 'Tỷ lệ khách kết nối WhatsApp', 0, 'Cộng vào bảng lương'])
    
    if extra_programs:
        for prog in extra_programs:
            ws_special.append([
                prog.get('name', 'Chương trình phát sinh'),
                prog.get('region', 'Toàn hệ thống'),
                prog.get('condition', ''),
                prog.get('bonus', 0),
                prog.get('note', 'Chương trình bổ sung')
            ])
    style_sheet(ws_special)
    
    wb.save(output_filepath)
    print(f"✅ Đã tạo thành công Gói Kế Hoạch KPI: {output_filepath}")
    
    return {
        'filepath': output_filepath,
        'filename': os.path.basename(output_filepath),
        'period': period_str,
        'stores_count': len(all_stores),
        'staff_count': len(seen_staff),
        'skus_count': len(seen_skus)
    }

def build_monthly_plan_file(
    month_num=8,
    input_folder=None,
    extra_programs=None,
    custom_branches=None,
    custom_staff=None,
    output_filepath=None
):
    # Fallback / backward-compatibility wrapper
    hcm_prop = os.path.join(BASE_DIR, '..', 'TOOL_KPISHEET', 'KPI CNT HCM Tháng 09.xlsx')
    hn_prop = os.path.join(BASE_DIR, '..', 'TOOL_KPISHEET', 'e_xuat_KPI_Quy_3.26.xlsx')
    if os.path.exists(hcm_prop) and os.path.exists(hn_prop):
        return build_plan_from_manager_proposals(
            hcm_prop, hn_prop, month_num=month_num,
            project_folder=input_folder,
            extra_programs=extra_programs,
            output_filepath=output_filepath
        )
    return {}

if __name__ == '__main__':
    hcm_file = os.path.join(BASE_DIR, '..', 'TOOL_KPISHEET', 'KPI CNT HCM Tháng 09.xlsx')
    hn_file = os.path.join(BASE_DIR, '..', 'TOOL_KPISHEET', 'e_xuat_KPI_Quy_3.26.xlsx')
    res = build_plan_from_manager_proposals(hcm_file, hn_file, month_num=9)
    print("Result:", res)
