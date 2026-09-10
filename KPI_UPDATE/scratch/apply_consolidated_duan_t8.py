import sys, os, openpyxl, subprocess
sys.stdout.reconfigure(encoding='utf-8')

def process_kpi_file(kpi_path):
    print(f"\n=======================================================")
    print(f"--> Đang cập nhật gộp doanh số xoay ca Dự án T8 cho: {kpi_path}")
    print(f"=======================================================")
    
    wb = openpyxl.load_workbook(kpi_path, data_only=False)
    if 'Dự án T8' not in wb.sheetnames:
        print("Không tìm thấy sheet 'Dự án T8'!")
        return
        
    ws_da = wb['Dự án T8']
    
    # 1. Đọc mapping chi nhánh chính từ 'kpi dược sĩ'
    primary_branch = {}
    if 'kpi dược sĩ' in wb.sheetnames:
        ws_ds = wb['kpi dược sĩ']
        for r in range(3, ws_ds.max_row + 1):
            cn = ws_ds.cell(r, 1).value
            name = ws_ds.cell(r, 2).value
            if name and str(name).strip() and str(name).strip() != 'Medigo':
                primary_branch[str(name).strip()] = str(cn).strip() if cn else ''

    # 2. Quét và gom toàn bộ doanh số CK & Combo theo từng Dược sĩ
    staff_totals = {}
    row_headers = []
    
    for r in range(3, ws_da.max_row + 1):
        cn = ws_da.cell(r, 1).value
        name = ws_da.cell(r, 2).value
        chuc_vu = ws_da.cell(r, 3).value
        
        if not name or str(name).strip() == 'Medigo' or not str(name).strip():
            continue
            
        n_str = str(name).strip()
        
        # Doanh số CK & Combo
        try:
            val_f = ws_da.cell(r, 6).value
            ck_val = float(val_f) if val_f and isinstance(val_f, (int, float)) else 0.0
        except:
            ck_val = 0.0
            
        try:
            val_g = ws_da.cell(r, 7).value
            combo_val = float(val_g) if val_g and isinstance(val_g, (int, float)) else 0.0
        except:
            combo_val = 0.0
            
        try:
            val_l = ws_da.cell(r, 12).value
            hb_val = float(val_l) if val_l and isinstance(val_l, (int, float)) else 0.0
        except:
            hb_val = 0.0
            
        main_cn = primary_branch.get(n_str, str(cn).strip() if cn else '')
        
        if n_str not in staff_totals:
            staff_totals[n_str] = {
                'main_cn': main_cn,
                'chuc_vu': chuc_vu or 'DSBC',
                'ck': 0.0,
                'combo': 0.0,
                'hotbill': 0.0,
                'branches': []
            }
            row_headers.append(n_str)
            
        staff_totals[n_str]['ck'] += ck_val
        staff_totals[n_str]['combo'] += combo_val
        staff_totals[n_str]['hotbill'] += hb_val
        staff_totals[n_str]['branches'].append(str(cn).strip() if cn else '')

    print(f"--> Tổng hợp được {len(row_headers)} Dược sĩ chính thức trong sheet 'Dự án T8'.")

    # 3. Xóa sạch dữ liệu cũ từ dòng 3 trở đi
    for r in range(3, 100):
        for c in range(1, 16):
            ws_da.cell(r, c).value = None

    # 4. Ghi lại dữ liệu đã gộp 100% doanh số cho từng Dược sĩ
    # Sắp xếp theo Chi nhánh chính để bảng biểu đẹp và chuẩn
    sorted_staff = sorted(row_headers, key=lambda n: (staff_totals[n]['main_cn'], n))

    for idx, n_str in enumerate(sorted_staff, start=3):
        info = staff_totals[n_str]
        ws_da.cell(idx, 1, info['main_cn'])
        ws_da.cell(idx, 2, n_str)
        ws_da.cell(idx, 3, info['chuc_vu'])
        ws_da.cell(idx, 4, f"=SUM(F{idx}:G{idx})/DAY($A$1)")
        ws_da.cell(idx, 5, 0) # Target or other
        ws_da.cell(idx, 6, round(info['ck'], 0)).number_format = '#,##0'
        ws_da.cell(idx, 7, round(info['combo'], 0)).number_format = '#,##0'
        ws_da.cell(idx, 8, f"=F{idx}/DAY($A$1)").number_format = '#,##0'
        ws_da.cell(idx, 9, f"=G{idx}/DAY($A$1)").number_format = '#,##0'
        
        # Công thức thưởng mốc ngày (Cột J)
        f_j = f'=IFERROR(IF(OR($A{idx}="Hàng Bông",$A{idx}="Đường Láng"),_xlfn.IFS(AND($H{idx}>=3400000,$I{idx}>=900000),4000000,AND($H{idx}>=3200000,$I{idx}>=750000),2800000,AND($H{idx}>=2800000,$I{idx}>=650000),2200000,AND($H{idx}>=2500000,$I{idx}>=520000),1600000,AND($H{idx}>=2200000,$I{idx}>=450000),1000000,TRUE,0),_xlfn.IFS(AND($H{idx}>=1500000,$I{idx}>=700000),4000000,AND($H{idx}>=1200000,$I{idx}>=650000),2800000,AND($H{idx}>=1000000,$I{idx}>=580000),2200000,AND($H{idx}>=850000,$I{idx}>=480000),1600000,AND($H{idx}>=630000,$I{idx}>=350000),1000000,TRUE,0)),0)'
        ws_da.cell(idx, 10, f_j).number_format = '#,##0'
        
        # Thưởng sàn 500k nếu chưa đạt mốc (Cột K)
        f_k = f'=IFERROR(IF(AND($J{idx}=0, IF(OR($A{idx}="Hàng Bông",$A{idx}="Đường Láng"),$D{idx}>=2000000,$D{idx}>=950000)), 500000, 0), 0)'
        ws_da.cell(idx, 11, f_k).number_format = '#,##0'
        
        # Hot Bill HN (Cột L)
        f_l = f"=IFERROR(SUMIFS('Hot Bill HN'!$G:$G, 'Hot Bill HN'!$C:$C, $A{idx}, 'Hot Bill HN'!$D:$D, $B{idx}), 0)"
        ws_da.cell(idx, 12, f_l).number_format = '#,##0'
        
        # Total thưởng (Cột M)
        ws_da.cell(idx, 13, f"=IFERROR($J{idx}+$K{idx}+$L{idx}, 0)").number_format = '#,##0'

    wb.save(kpi_path)
    print(f"✅ Đã cập nhật xong file KPI: {kpi_path}")

# Run on both KPI files
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
kpi_root = os.path.join(base_dir, 'baocaokpi_thang8_hoanthien.xlsx')
kpi_thang8 = os.path.join(base_dir, 'thang8', 'baocaokpi_thang8_hoanthien.xlsx')

if os.path.exists(kpi_root): process_kpi_file(kpi_root)
if os.path.exists(kpi_thang8): process_kpi_file(kpi_thang8)
