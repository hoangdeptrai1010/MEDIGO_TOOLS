import os, openpyxl

base_dir = r"d:\MEDIGO\KPI_UPDATE\thang8\Pharmacy_retail_Store_KPIs_August 2026"

def dump_excel(fpath):
    wb = openpyxl.load_workbook(fpath, data_only=True)
    out = [f"=== FILE: {os.path.basename(fpath)} ==="]
    for sname in wb.sheetnames:
        ws = wb[sname]
        out.append(f"\n--- Sheet: {sname} ({ws.max_row} rows x {ws.max_column} cols) ---")
        for r in range(1, ws.max_row+1):
            vals = [str(ws.cell(r, c).value or '').strip() for c in range(1, ws.max_column+1)]
            val_str = ' | '.join([v for v in vals if v])
            if val_str:
                out.append(f"Row {r:2d}: {val_str}")
    return "\n".join(out)

# 1. Hanoi Mini Projects
for f in ["HN_AVC project_August_2026.xlsx", "HN_KAT Project August_2026.xlsx", "HN_Ladycare Project_August _2026.xlsx", "HN_Party Smart_Project _August _2026.xlsx"]:
    p = os.path.join(base_dir, r"Hà Nội\Dự án Mini\Chương trình", f)
    print(dump_excel(p))
    print("\n" + "="*80 + "\n")

# 2. HCM Mini Projects
for f in ["HCM_KAT Project August_2026.xlsx", "HCM_Ladycare Project_August_2026.xlsx", "HCM_Party Smart_Project _August _2026.xlsx"]:
    p = os.path.join(base_dir, r"Hồ Chí Minh\Dự án mini\Chương trình", f)
    print(dump_excel(p))
    print("\n" + "="*80 + "\n")
