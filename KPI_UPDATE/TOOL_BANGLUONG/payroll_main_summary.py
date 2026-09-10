"""
MODULE 10: HOÀN THIỆN SHEET TRUNG TÂM 'BẢNG LƯƠNG'
- Tự động đồng bộ 100% tất cả các cặp (Chi nhánh, Nhân viên) phát sinh chấm công từ sheet 'Giờ công' / 'Ngày công'
- Tự động bổ sung các dòng phân bổ chi nhánh còn thiếu cho nhân sự xoay ca (Hoàng Lâm Gia Bảo, Phan Công Vũ Tài, Võ Ngọc Giàu Sang, Nguyễn Trí Nghĩa...)
- Gắn công thức tính toán tự động chuẩn 100%:
  + Cột E (Tổng giờ ca ngày): =SUM(G, I, L, O, S, T)
  + Cột F (Tổng giờ ca đêm): =SUM(H, P)
  + Cột G (Số giờ ca ngày): =SUMIFS('Giờ công'!I:I, 'Giờ công'!G:G, B, 'Giờ công'!H:H, C)
  + Cột H (Số giờ ca đêm): =SUMIFS('Giờ công'!J:J, 'Giờ công'!G:G, B, 'Giờ công'!H:H, C)
  + Cột I (Tăng ca OT do làm >= 29 hay >= 30 công): =IF(ROW()=MATCH(C, $C$1:$C${last_r}, 0), IF(SUMIF($C$3:$C${last_r}, C, $R$3:$R${last_r})>=30, 16, IF(SUMIF($C$3:$C${last_r}, C, $R$3:$R${last_r})>=29, 8, "")), "")
  + Cột J (Số ngày làm ngày): =ROUND(G/24, 1)
  + Cột K (Số ca làm đêm): =SUMIFS('Ngày công'!I:I, 'Ngày công'!A:A, B, 'Ngày công'!B:B, C)
  + Cột R (Số ngày công thực tế): =SUMIFS('Ngày công'!M:M, 'Ngày công'!K:K, B, 'Ngày công'!L:L, C)
  + Cột V (Phụ cấp CHT): =IF(AND(ROW()=MATCH(C, $C$1:$C${last_r}, 0), OR(TRIM(D)="CHT", TRIM(D)="Q.CHT")), 500000/31*SUMIF($C$3:$C${last_r}, C, $R$3:$R${last_r}), 0)
  + Cột X (Phụ cấp sức khỏe ca đêm): =IF(ISNUMBER(SEARCH("DSCD", D)), MIN(1500000, ROUND(1500000/28*K, 0)), IF(K>20, MIN(1500000, ROUND(1500000/28*K, 0)), 0))
  + Cột AA (Thưởng MiniKat Dược sĩ)
  + Cột AB (Thưởng MiniKat CHT)
  + Cột AC (Thưởng WhatsApp)
  + Cột AD (Thưởng Dự án)
  + Cột AE (Thưởng KPI)
  + Cột AF (Thưởng Hàng điểm CK)
  + Cột AG (Thưởng Cận date)
  + Cột AH (Thưởng Maps)
  + Cột AJ-AP (Giảm trừ KPI)
- Dòng Tổng cộng & Bảng tổng hợp theo chi nhánh
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
try:
    from payroll_candate import normalize_branch
except ImportError:
    from TOOL_BANGLUONG.payroll_candate import normalize_branch

def process_main_payroll_sheet(wb_out):
    if 'BẢNG LƯƠNG' not in wb_out.sheetnames:
        return

    print(f"--> [Module Bảng Lương] Đang gắn toàn bộ công thức chuẩn vào sheet 'BẢNG LƯƠNG'...")
    ws_bl = wb_out['BẢNG LƯƠNG']

    # 1. Thu thập tất cả các cặp (Chi nhánh, Nhân sự) có phát sinh từ Giờ công
    all_timecard_pairs = []
    seen_pairs = set()
    if 'Giờ công' in wb_out.sheetnames:
        ws_gc = wb_out['Giờ công']
        for r in range(2, ws_gc.max_row + 1):
            br_val = ws_gc.cell(r, 1).value
            nm_val = ws_gc.cell(r, 2).value
            if br_val and nm_val and str(nm_val).strip() and str(nm_val).strip() != 'Tổng':
                b_norm = normalize_branch(br_val)
                n_clean = str(nm_val).strip()
                k = (b_norm, n_clean)
                if k not in seen_pairs:
                    seen_pairs.add(k)
                    all_timecard_pairs.append(k)

    # 2. Quét các dòng nhân sự hiện có trong BẢNG LƯƠNG
    existing_bl_map = {} # (branch, name) -> row_num
    last_staff_row = 2
    for r in range(3, ws_bl.max_row + 1):
        b_val = ws_bl.cell(r, 2).value
        n_val = ws_bl.cell(r, 3).value
        stt_val = ws_bl.cell(r, 1).value
        if b_val and n_val and str(n_val).strip() != 'Tổng':
            b_norm = normalize_branch(b_val)
            n_clean = str(n_val).strip()
            existing_bl_map[(b_norm, n_clean)] = r
            last_staff_row = max(last_staff_row, r)
        elif str(ws_bl.cell(r, 4).value or '').strip() == 'Tổng':
            break

    # 3. Tra cứu chức danh của nhân sự từ các dòng khác hoặc danh sách
    role_lookup = {}
    for (b, n), r in existing_bl_map.items():
        role = ws_bl.cell(r, 4).value
        if role:
            role_lookup[n] = str(role).strip()

    # 4. Bổ sung các cặp (Chi nhánh, Nhân viên) còn thiếu vào BẢNG LƯƠNG (ví dụ: Hoàng Lâm Gia Bảo ở Trường Sa)
    missing_pairs = [k for k in all_timecard_pairs if k not in existing_bl_map]
    if missing_pairs:
        print(f"--> [Bảng Lương] Phát hiện {len(missing_pairs)} dòng công tác xoay ca còn thiếu, tự động bổ sung:")
        for (b, n) in missing_pairs:
            last_staff_row += 1
            stt = last_staff_row - 2
            role = role_lookup.get(n, 'DSXC')
            existing_bl_map[(b, n)] = last_staff_row
            ws_bl.cell(last_staff_row, 1, stt)
            ws_bl.cell(last_staff_row, 2, b)
            ws_bl.cell(last_staff_row, 3, n)
            ws_bl.cell(last_staff_row, 4, role)
            print(f"     ➕ Dòng {last_staff_row}: [{b}] {n} ({role})")

    # 5. Gắn công thức tính toán hoàn chỉnh cho toàn bộ các dòng nhân sự
    last_r_str = str(last_staff_row)
    for r in range(3, last_staff_row + 1):
        name_val = ws_bl.cell(r, 3).value
        if not name_val or not str(name_val).strip() or str(name_val).strip() == 'Tổng':
            continue

        # Cột E & F: Tổng giờ
        ws_bl.cell(r, 5, f"=SUM(G{r}, I{r}, L{r}, O{r}, S{r}, T{r})")
        ws_bl.cell(r, 6, f"=SUM(H{r}, P{r})")

        # Cột G & H: Giờ ca ngày, ca đêm
        ws_bl.cell(r, 7, f"=SUMIFS('Giờ công'!I:I, 'Giờ công'!G:G, B{r}, 'Giờ công'!H:H, C{r})")
        ws_bl.cell(r, 8, f"=SUMIFS('Giờ công'!J:J, 'Giờ công'!G:G, B{r}, 'Giờ công'!H:H, C{r})")

        # Cột I: Chuyên cần / Tăng ca OT vượt định mức
        ws_bl.cell(r, 9, f'=IF(ROW()=MATCH(C{r}, $C$1:$C${last_r_str}, 0), IF(SUMIF($C$3:$C${last_r_str}, C{r}, $R$3:$R${last_r_str})>=30, 16, IF(SUMIF($C$3:$C${last_r_str}, C{r}, $R$3:$R${last_r_str})>=29, 8, "")), "")')

        # Cột J: Quy đổi ngày làm ngày
        ws_bl.cell(r, 10, f"=ROUND(G{r}/24, 1)")

        # Cột K: Số ca đêm
        ws_bl.cell(r, 11, f"=SUMIFS('Ngày công'!I:I, 'Ngày công'!A:A, B{r}, 'Ngày công'!B:B, C{r})")

        # Cột N, O, P: Lễ
        ws_bl.cell(r, 14, f"=SUMIFS('Tăng ca lễ'!M:M, 'Tăng ca lễ'!K:K, B{r}, 'Tăng ca lễ'!L:L, C{r})")
        ws_bl.cell(r, 15, f'=IF(OR(D{r}="PT", D{r}="DSHV"), SUMIFS(\'Tăng ca lễ\'!O:O, \'Tăng ca lễ\'!K:K, B{r}, \'Tăng ca lễ\'!L:L, C{r})*2, SUMIFS(\'Tăng ca lễ\'!O:O, \'Tăng ca lễ\'!K:K, B{r}, \'Tăng ca lễ\'!L:L, C{r})*3)')
        ws_bl.cell(r, 16, f'=IF(OR(D{r}="PT", D{r}="DSHV"), SUMIFS(\'Tăng ca lễ\'!P:P, \'Tăng ca lễ\'!K:K, B{r}, \'Tăng ca lễ\'!L:L, C{r})*2, SUMIFS(\'Tăng ca lễ\'!P:P, \'Tăng ca lễ\'!K:K, B{r}, \'Tăng ca lễ\'!L:L, C{r})*3)')

        # Cột R: Ngày công thực tế
        ws_bl.cell(r, 18, f"=SUMIFS('Ngày công'!M:M, 'Ngày công'!K:K, B{r}, 'Ngày công'!L:L, C{r})")

        # Cột S: Đào tạo
        ws_bl.cell(r, 19, f"=SUMIFS('Công Đào tạo'!G:G, 'Công Đào tạo'!E:E, B{r}, 'Công Đào tạo'!F:F, C{r})")

        # Cột V: Phụ cấp trách nhiệm CHT
        ws_bl.cell(r, 22, f'=IF(AND(ROW()=MATCH(C{r}, $C$1:$C${last_r_str}, 0), OR(TRIM(D{r})="CHT", TRIM(D{r})="Q.CHT")), 500000/31*SUMIF($C$3:$C${last_r_str}, C{r}, $R$3:$R${last_r_str}), 0)')

        # Cột X: Phụ cấp sức khỏe ca đêm
        ws_bl.cell(r, 24, f'=IF(ISNUMBER(SEARCH("DSCD", D{r})), MIN(1500000, ROUND(1500000/28*K{r}, 0)), IF(K{r}>20, MIN(1500000, ROUND(1500000/28*K{r}, 0)), 0))')

        # Cột AA & AB: Thưởng MiniKat
        ws_bl.cell(r, 27, f"=IF(ROW()=MATCH(C{r}, $C$1:$C${last_r_str}, 0), IFERROR(VLOOKUP(C{r}, 'MiniKat - HN'!$A:$L, 12, FALSE), 0) + IFERROR(VLOOKUP(C{r}, 'MiniKat - HCM'!$A:$J, 10, FALSE), 0), 0)")
        ws_bl.cell(r, 28, f"=IF(OR(TRIM(D{r})=\"CHT\", TRIM(D{r})=\"Q.CHT\"), SUMIFS('MiniKat - HN'!$Y:$Y, 'MiniKat - HN'!$N:$N, B{r}) + SUMIFS('MiniKat - HCM'!$U:$U, 'MiniKat - HCM'!$M:$M, B{r}), 0)")

        # Cột AC: WhatsApp
        ws_bl.cell(r, 29, f"=IF(ROW()=MATCH(C{r}, $C$1:$C${last_r_str}, 0), SUMIFS(whatsapp!$D:$D, whatsapp!$A:$A, C{r}), 0)")

        # Cột AD: Dự án
        ws_bl.cell(r, 30, f"=SUMIFS('Dự án'!$C:$C, 'Dự án'!$A:$A, B{r}, 'Dự án'!$B:$B, C{r})")

        # Cột AE: KPI
        ws_bl.cell(r, 31, f"=SUMIFS(KPI!$C:$C, KPI!$A:$A, B{r}, KPI!$B:$B, C{r})")

        # Cột AF: Hàng điểm CK
        ws_bl.cell(r, 32, f"=SUMIFS('Thưởng CK'!$C:$C, 'Thưởng CK'!$A:$A, B{r}, 'Thưởng CK'!$B:$B, C{r})")

        # Cột AG: Cận date
        ws_bl.cell(r, 33, f"=SUMIFS('Cận date'!$C:$C, 'Cận date'!$A:$A, B{r}, 'Cận date'!$B:$B, C{r})")

        # Cột AH: Maps
        ws_bl.cell(r, 34, f"=SUMIFS(Map!$D:$D, Map!$B:$B, B{r}, Map!$A:$A, C{r})")

        # Cột AJ - AP: Giảm trừ
        ws_bl.cell(r, 36, f"=SUM(AK{r}:AO{r})")
        ws_bl.cell(r, 37, f"=SUMIFS('KPI trừ'!$C:$C, 'KPI trừ'!$A:$A, B{r}, 'KPI trừ'!$B:$B, C{r})")
        ws_bl.cell(r, 38, f"=SUMIFS('KPI trừ'!$D:$D, 'KPI trừ'!$A:$A, B{r}, 'KPI trừ'!$B:$B, C{r})")
        ws_bl.cell(r, 39, f"=SUMIFS('KPI trừ'!$E:$E, 'KPI trừ'!$A:$A, B{r}, 'KPI trừ'!$B:$B, C{r})")
        ws_bl.cell(r, 40, f"=SUMIFS('KPI trừ'!$F:$F, 'KPI trừ'!$A:$A, B{r}, 'KPI trừ'!$B:$B, C{r})")
        ws_bl.cell(r, 41, f"=SUMIFS('KPI trừ'!$G:$G, 'KPI trừ'!$A:$A, B{r}, 'KPI trừ'!$B:$B, C{r})")
        ws_bl.cell(r, 42, f"=IFERROR(INDEX('KPI trừ'!$H:$H, MATCH(1, ('KPI trừ'!$A:$A=B{r})*('KPI trừ'!$B:$B=C{r}), 0)), \"\")")

        # Định dạng hiển thị số
        ws_bl.cell(r, 5).number_format = '#,##0.00'
        ws_bl.cell(r, 6).number_format = '#,##0.00'
        ws_bl.cell(r, 7).number_format = '#,##0.00'
        ws_bl.cell(r, 8).number_format = '#,##0.00'
        ws_bl.cell(r, 9).number_format = '#,##0'
        ws_bl.cell(r, 10).number_format = '#,##0.0'
        ws_bl.cell(r, 11).number_format = '#,##0'
        ws_bl.cell(r, 14).number_format = '#,##0'
        ws_bl.cell(r, 18).number_format = '#,##0'
        for col_idx in range(21, 42):
            ws_bl.cell(r, col_idx).number_format = '#,##0'

    # Dòng Tổng cộng sau hàng nhân sự cuối cùng
    r_bl_total = last_staff_row + 2
    ws_bl.cell(r_bl_total, 4, 'Tổng')
    total_cols = [5, 6, 7, 8, 9, 10, 11, 14, 15, 16, 18, 19, 20] + list(range(22, 42))
    for c_idx in total_cols:
        col_letter = openpyxl.utils.get_column_letter(c_idx)
        ws_bl.cell(r_bl_total, c_idx, f"=SUM({col_letter}3:{col_letter}{last_staff_row})")
        if c_idx in [5, 6, 7, 8]: ws_bl.cell(r_bl_total, c_idx).number_format = '#,##0.00'
        elif c_idx == 10: ws_bl.cell(r_bl_total, c_idx).number_format = '#,##0.0'
        else: ws_bl.cell(r_bl_total, c_idx).number_format = '#,##0'

    # Bảng tổng hợp chi nhánh (Dưới dòng tổng)
    all_branches = [
        'Trường Sa', 'Đỗ Quang Đẩu', 'Nam Hòa', 'Minh Châu', 'Lê Bình',
        'Nguyễn Chí Thanh', 'Nguyễn Thị Thập', 'Nguyễn Văn Quá', 'Rạch Bùng Binh',
        'Đường Láng', 'Hàng Bông'
    ]
    for idx, b_name in enumerate(all_branches):
        r_b = r_bl_total + 3 + idx
        ws_bl.cell(r_b, 5, b_name)
        ws_bl.cell(r_b, 7, f"=SUMIF(B:B, E{r_b}, G:G)").number_format = '#,##0.00'
        ws_bl.cell(r_b, 8, f"=SUMIF(B:B, E{r_b}, H:H)").number_format = '#,##0.00'

    print(f"✅ Đã hoàn thiện toàn bộ sheet 'BẢNG LƯƠNG' với {last_staff_row - 2} dòng nhân sự!")
