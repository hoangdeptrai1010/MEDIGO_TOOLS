# -*- coding: utf-8 -*-
import openpyxl, os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

for p in ['thang8/bangluong_thang8_hoanthien.xlsx', 'thang8/BANGLUONGTHANG8.xlsx']:
    if not os.path.exists(p):
        continue
    print(f"--> Updating Col J formula in {p}...")
    try:
        wb = openpyxl.load_workbook(p)
        ws = wb['BẢNG LƯƠNG']
        ws.cell(2, 10, "Số ngày làm ngày")
        ws.cell(2, 11, "Số ca làm đêm")
        for r in range(3, 59):
            # Cột J: Quy đổi số giờ ca ngày (G) sang ngày theo quy ước 24 giờ = 1 ngày (3 ca 8h = 24h = 1 ngày)
            ws.cell(r, 10, f"=ROUND(G{r}/24, 0)")
            # Cột K: Ca đêm >= 4h
            ws.cell(r, 11, f"=SUMIFS('Ngày công'!I:I, 'Ngày công'!B:B, C{r})")
            # Cột X: Phụ cấp ca đêm
            ws.cell(r, 24, f'=IF(ISNUMBER(SEARCH("DSCD", D{r})), ROUND(1500000/28*K{r}, 0), IF(K{r}>20, ROUND(1500000/28*K{r}, 0), 0))')
        wb.save(p)
        print(f"✅ Saved {p}")
    except PermissionError:
        print(f"⚠️ File {p} đang được mở trong Excel bởi người dùng, tạm thời bỏ qua lưu trực tiếp (sẽ đồng bộ sau).")
    except Exception as e:
        print(f"Error {p}: {e}")
