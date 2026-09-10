import os, openpyxl, zipfile
import xml.etree.ElementTree as ET

base_dir = r"d:\MEDIGO\KPI_UPDATE\thang8\Pharmacy_retail_Store_KPIs_August 2026"

def read_excel_detail(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    res = [f"Workbook: {os.path.basename(path)} | Sheets: {wb.sheetnames}"]
    for sname in wb.sheetnames:
        ws = wb[sname]
        res.append(f"\n--- Sheet '{sname}' ({ws.max_row} rows x {ws.max_column} cols) ---")
        # Print first 5 rows
        for r in range(1, min(ws.max_row+1, 8)):
            row_vals = [str(ws.cell(r, c).value or '') for c in range(1, min(ws.max_column+1, 10))]
            res.append(f"  Row {r}: " + " | ".join(row_vals))
    return "\n".join(res)

print("=== 1. NỘI DUNG CK-HN.XLSX ===")
print(read_excel_detail(os.path.join(base_dir, r"Hà Nội\Dự án\Chương trình\CK-HN.xlsx")))

print("\n=== 2. NỘI DUNG MINI PROJECTS HÀ NỘI ===")
for f in ["HN_AVC project_August_2026.xlsx", "HN_KAT Project August_2026.xlsx", "HN_Ladycare Project_August _2026.xlsx", "HN_Party Smart_Project _August _2026.xlsx"]:
    fpath = os.path.join(base_dir, r"Hà Nội\Dự án Mini\Chương trình", f)
    print(read_excel_detail(fpath))

print("\n=== 3. NỘI DUNG MINI PROJECTS HCM ===")
for f in ["HCM_KAT Project August_2026.xlsx", "HCM_Ladycare Project_August_2026.xlsx", "HCM_Party Smart_Project _August _2026.xlsx"]:
    fpath = os.path.join(base_dir, r"Hồ Chí Minh\Dự án mini\Chương trình", f)
    print(read_excel_detail(fpath))
