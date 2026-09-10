import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb_kpi = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=True)
ws_da = wb_kpi['Dự án T8']

staff_rows = {}
for r in range(3, ws_da.max_row + 1):
    cn = ws_da.cell(r, 1).value
    name = ws_da.cell(r, 2).value
    tot = ws_da.cell(r, 13).value or 0
    t_da = ws_da.cell(r, 10).value or 0
    t_add = ws_da.cell(r, 11).value or 0
    t_hb = ws_da.cell(r, 12).value or 0
    if name and str(name).strip():
        n_str = str(name).strip()
        staff_rows.setdefault(n_str, []).append((r, cn, tot, t_da, t_add, t_hb))

print("=== NHÂN SỰ XUẤT HIỆN NHIỀU HƠN 1 DÒNG TRONG DỰ ÁN T8 ===")
for name, rows in staff_rows.items():
    if len(rows) > 1:
        print(f"\n* {name} ({len(rows)} dòng):")
        for r, cn, tot, t_da, t_add, t_hb in rows:
            print(f"   - Dòng {r:2d} ({cn}): Total={tot:,.0f} (DA={t_da:,.0f}, Thêm={t_add:,.0f}, HB={t_hb:,.0f})")

