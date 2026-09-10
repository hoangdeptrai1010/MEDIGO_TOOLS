import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

print("=== CHECKING JULY THƯỞNG CK & DỰ ÁN ===")
wb_july = openpyxl.load_workbook('thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=False)
for s in ['Thưởng CK', 'Dự án']:
    if s in wb_july.sheetnames:
        ws = wb_july[s]
        print(f"\n--- July Sheet: {s} ---")
        for r in range(1, 15):
            print(f"Row {r:2d}: {[ws.cell(r, c).value for c in range(1, 15)]}")

print("\n=== CHECKING JULY KPI REPORT ===")
for kpi_f in ['thang7/BaoCao_CanDate_Thang7_HoanThien.xlsx', 'baocaokpi_thang7_hoanthien.xlsx']:
    try:
        wb = openpyxl.load_workbook(kpi_f, read_only=True)
        print(f"{kpi_f} sheets: {wb.sheetnames}")
    except Exception as e:
        print(f"{kpi_f} error: {e}")

