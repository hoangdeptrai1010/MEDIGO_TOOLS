import openpyxl

wb = openpyxl.load_workbook('d:/MEDIGO/KPI_UPDATE/thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws_bl = wb['BẢNG LƯƠNG']

print('='*105)
print('KIỂM TRA TOÀN DIỆN NGÀY CÔNG CỦA 70 NHÂN VIÊN TRONG BẢNG LƯƠNG HOÀN THIỆN THÁNG 8:')
print('='*105)
print(f'{"STT":4} | {"Chi nhánh":15} | {"Tên nhân viên":25} | {"Chức danh":8} | {"J(Ngày làm)":11} | {"K(Làm đêm)":10} | {"R(Ngày công)":12} | {"PC CHT(Col V)":12}')
print('-'*105)

over_31_bl = []
last_stt = 0
for r in range(3, ws_bl.max_row + 1):
    stt = ws_bl.cell(r, 1).value
    if isinstance(stt, (int, float)):
        last_stt = int(stt)
        cn = str(ws_bl.cell(r, 2).value or '').strip()
        name = str(ws_bl.cell(r, 3).value or '').strip()
        role = str(ws_bl.cell(r, 4).value or '').strip()
        cj = ws_bl.cell(r, 10).value
        ck = ws_bl.cell(r, 11).value
        cr = ws_bl.cell(r, 18).value
        cv = ws_bl.cell(r, 22).value
        
        cv_str = f'{cv:,.0f}' if isinstance(cv, (int, float)) else str(cv)
        print(f'{int(stt):3d}  | {cn:15} | {name:25} | {role:8} | {str(cj):11} | {str(ck):10} | {str(cr):12} | {cv_str:12}')
        
        if (isinstance(cj, (int, float)) and cj > 31) or (isinstance(cr, (int, float)) and cr > 31):
            over_31_bl.append((int(stt), cn, name, cj, cr))

print('='*105)
print(f'TỔNG SỐ NHÂN SỰ ĐƯỢC KIỂM TRA: {last_stt}')
print(f'SỐ NHÂN SỰ BỊ NGÀY CÔNG > 31: {len(over_31_bl)}')
if len(over_31_bl) == 0:
    print('==> XÁC NHẬN: 100% NHÂN SỰ ĐỀU CÓ SỐ NGÀY CÔNG <= 31 NGÀY, HOÀN TOÀN CHÍNH XÁC!')
else:
    print('LỖI CÒN LẠI:', over_31_bl)
