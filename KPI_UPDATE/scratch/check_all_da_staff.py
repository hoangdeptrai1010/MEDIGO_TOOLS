import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb_kpi = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=True)
ws_da = wb_kpi['Dự án T8']
ws_kpi_ds = wb_kpi['kpi dược sĩ']

# Read KPI %
kpi_pct = {}
for r in range(3, ws_kpi_ds.max_row + 1):
    name = ws_kpi_ds.cell(r, 2).value
    pct = ws_kpi_ds.cell(r, 11).value # % KPI
    if name:
        kpi_pct[str(name).strip()] = pct

print("=== CHI TIẾT TỪNG NHÂN SỰ ĐẠT THƯỞNG DỰ ÁN T8 & HỆ SỐ KPI ===")
print(f"{'STT':<4} | {'Chi nhánh':<15} | {'Họ và tên':<25} | {'Chức danh':<8} | {'% KPI':<8} | {'Hệ số':<6} | {'Thưởng DA':<10} | {'Thêm 500k':<10} | {'Hot Bill':<10} | {'Total DA':<12}")
print("-" * 125)

stt = 0
for r in range(3, ws_da.max_row + 1):
    cn = ws_da.cell(r, 1).value
    name = ws_da.cell(r, 2).value
    cd = ws_da.cell(r, 3).value
    t_da = ws_da.cell(r, 10).value or 0
    t_them = ws_da.cell(r, 11).value or 0
    t_hb = ws_da.cell(r, 12).value or 0
    tot_da = ws_da.cell(r, 13).value or 0
    
    if name and (tot_da > 0 or t_da > 0 or t_them > 0 or t_hb > 0):
        stt += 1
        n_str = str(name).strip()
        pct = kpi_pct.get(n_str, 0)
        pct_str = f"{pct*100:.1f}%" if isinstance(pct, (int, float)) else str(pct)
        he_so = 0.8 if (isinstance(pct, (int, float)) and pct < 0.6) else 1.0
        print(f"{stt:<4} | {cn:<15} | {n_str:<25} | {str(cd):<8} | {pct_str:<8} | {he_so:<6} | {t_da:>10,f} | {t_them:>10,f} | {t_hb:>10,f} | {tot_da:>12,f}")

