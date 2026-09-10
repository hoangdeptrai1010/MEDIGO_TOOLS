import openpyxl, sys
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)

print("==================== MiniKat - HN ====================")
ws = wb['MiniKat - HN']
for r in range(1, 15):
    seller_row = [ws.cell(r, c).value for c in range(1, 11)]
    branch_row = [ws.cell(r, c).value for c in range(13, 21)]
    if any(seller_row) or any(branch_row):
        print(f"Row {r:2d} | Seller: {seller_row} | Branch: {branch_row}")

print("\n==================== MiniKat - HCM (Sample 20 rows) ====================")
ws = wb['MiniKat - HCM']
for r in range(1, 65):
    seller_row = [ws.cell(r, c).value for c in range(1, 11)]
    branch_row = [ws.cell(r, c).value for c in range(13, 21)]
    if any(seller_row) or any(branch_row):
        print(f"Row {r:2d} | Seller: {seller_row} | Branch: {branch_row}")
