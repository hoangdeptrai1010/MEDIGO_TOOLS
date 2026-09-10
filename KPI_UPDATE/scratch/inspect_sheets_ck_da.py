import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

print("================================================================================")
print("CHI TIẾT SHEET 'Thưởng CK' & 'Dự án' - THÁNG 7 (Target vs Gốc)")
print("================================================================================")

wb7_t = openpyxl.load_workbook(r'thang7\target\BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=False)
wb7_tv = openpyxl.load_workbook(r'thang7\target\BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=True)

ws7_ck = wb7_t['Thưởng CK']
ws7_ckv = wb7_tv['Thưởng CK']
print("\n--- Sheet Thưởng CK Tháng 7 (10 dòng đầu) ---")
for r in range(1, 12):
    f_row = [ws7_ck.cell(r, c).value for c in range(1, 13)]
    v_row = [ws7_ckv.cell(r, c).value for c in range(1, 13)]
    print(f"R{r} FORMULA: {f_row}")
    print(f"R{r} VALUE:   {v_row}")

ws7_da = wb7_t['Dự án']
ws7_dav = wb7_tv['Dự án']
print("\n--- Sheet Dự án Tháng 7 (10 dòng đầu) ---")
for r in range(1, 12):
    f_row = [ws7_da.cell(r, c).value for c in range(1, 13)]
    v_row = [ws7_dav.cell(r, c).value for c in range(1, 13)]
    print(f"R{r} FORMULA: {f_row}")
    print(f"R{r} VALUE:   {v_row}")

print("\n================================================================================")
print("CHI TIẾT SHEET 'Thưởng CK' & 'Dự án' - THÁNG 8 (Hoàn thiện)")
print("================================================================================")

wb8_t = openpyxl.load_workbook(r'thang8\bangluong_thang8_hoanthien.xlsx', data_only=False)
wb8_tv = openpyxl.load_workbook(r'thang8\bangluong_thang8_hoanthien.xlsx', data_only=True)

ws8_ck = wb8_t['Thưởng CK']
ws8_ckv = wb8_tv['Thưởng CK']
print("\n--- Sheet Thưởng CK Tháng 8 (10 dòng đầu) ---")
for r in range(1, 12):
    f_row = [ws8_ck.cell(r, c).value for c in range(1, 13)]
    v_row = [ws8_ckv.cell(r, c).value for c in range(1, 13)]
    print(f"R{r} FORMULA: {f_row}")
    print(f"R{r} VALUE:   {v_row}")

ws8_da = wb8_t['Dự án']
ws8_dav = wb8_tv['Dự án']
print("\n--- Sheet Dự án Tháng 8 (10 dòng đầu) ---")
for r in range(1, 12):
    f_row = [ws8_da.cell(r, c).value for c in range(1, 13)]
    v_row = [ws8_dav.cell(r, c).value for c in range(1, 13)]
    print(f"R{r} FORMULA: {f_row}")
    print(f"R{r} VALUE:   {v_row}")

