import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

wb_raw = openpyxl.load_workbook('thang8/DATA/BangChiTietChamCong_thang8.xlsx', data_only=True)
ws_th = wb_raw['Bảng tổng hợp chấm công']
ws_ct = wb_raw['Bảng chi tiết chấm công']
ws_h = wb_raw['Bảng chấm công theo giờ']

print("--- Check Bảng chấm công theo giờ row 1 to 10 ---")
for r in range(1, 10):
    print(r, [ws_h.cell(r, c).value for c in range(1, 15)])

print("\n--- Check Bảng chi tiết chấm công row 1 to 10 ---")
for r in range(1, 10):
    print(r, [ws_ct.cell(r, c).value for c in range(1, 12)])
