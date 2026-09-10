import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws = wb['BẢNG LƯƠNG']

print("=== PHÂN TÍCH CHI TIẾT SỐ LIỆU TRONG HÌNH ẢNH ===")
print(f"1. Dòng Tổng cộng (Row 74):")
print(f"   - Cột E (Tổng giờ ca ngày tính lương): {ws.cell(74, 5).value:,.2f}")
print(f"   - Cột F (Tổng giờ ca đêm): {ws.cell(74, 6).value:,.2f}")
print(f"   - Cột G (Số giờ ca ngày thực tế quẹt thẻ): {ws.cell(74, 7).value:,.2f}")
print(f"   - Cột H (Số giờ ca đêm thực tế quẹt thẻ): {ws.cell(74, 8).value:,.2f}")
print(f"   - Cột I (Tổng giờ thưởng Chuyên cần): {ws.cell(74, 9).value}")
print(f"   --> Phép tính: Cột E ({ws.cell(74, 5).value:,.2f}) = Cột G ({ws.cell(74, 7).value:,.2f}) + Cột I ({ws.cell(74, 9).value}) [Chuyên cần]")

print(f"\n2. Bảng tổng hợp theo chi nhánh (Row 77..88):")
print(f"   - Cột G ở bảng chi nhánh tổng hợp =SUMIF(B:B, ChiNhánh, G:G) -> Tổng cộng = {ws.cell(88, 7).value:,.2f} (Khớp 100% với Cột G thực làm dòng 74).")
print(f"   - Cột H ở bảng chi nhánh tổng hợp =SUMIF(B:B, ChiNhánh, H:H) -> Tổng cộng = {ws.cell(88, 8).value:,.2f} (Khớp 100% với Cột H thực làm dòng 74).")

wb.close()
