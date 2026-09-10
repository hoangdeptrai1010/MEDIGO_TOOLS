import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/test_rebuilt_timecards.xlsx', data_only=False)
ws_nc = wb['Ngày công']
ws_gc = wb['Giờ công']
ws_bl = wb['BẢNG LƯƠNG']

# Read gc detail (rows 2..ws_gc.max_row)
gc_rows = []
for r in range(2, ws_nc.max_row + 1):
    br = ws_gc.cell(r, 1).value
    name = ws_gc.cell(r, 2).value
    shift = ws_gc.cell(r, 3).value
    hrs = ws_gc.cell(r, 5).value
    if br and name and hrs is not None:
        try:
            gc_rows.append((str(br).strip(), str(name).strip(), str(shift).strip(), float(hrs)))
        except:
            pass

# Read nc detail (rows 2..ws_nc.max_row)
nc_rows = []
for r in range(2, ws_nc.max_row + 1):
    br = ws_nc.cell(r, 1).value
    name = ws_nc.cell(r, 2).value
    dt = ws_nc.cell(r, 3).value
    shift = ws_nc.cell(r, 4).value
    hrs = ws_nc.cell(r, 6).value
    if br and name and hrs is not None:
        try:
            nc_rows.append({
                'br': str(br).strip(),
                'name': str(name).strip(),
                'dt': dt,
                'shift': str(shift).strip(),
                'hrs': float(hrs),
                'is_night': ('đêm' in str(shift).lower() or 'qua đêm' in str(shift).lower()),
                'gt4': float(hrs) > 4.0
            })
        except:
            pass

print(f"Total GC rows: {len(gc_rows)}, Total NC rows: {len(nc_rows)}")

# Simulate SUMIFS for all BL rows
print("\n" + "="*90)
print(f"{'STT':<4} | {'Chi nhánh':<15} | {'Tên NV':<24} | {'Giờ ngày':<9} | {'Giờ đêm':<9} | {'Ngày(J)':<8} | {'Ca đêm(K)':<10} | {'Công TT(R)':<10}")
print("="*90)

for r in range(3, 73):
    stt = ws_bl.cell(r, 1).value
    cn = str(ws_bl.cell(r, 2).value or '').strip()
    name = str(ws_bl.cell(r, 3).value or '').strip()
    role = str(ws_bl.cell(r, 4).value or '').strip()
    
    if not name:
        continue
        
    # Col G: Giờ ca ngày = SUMIFS(gc E:E, gc B:B, name, gc A:A, cn) - Col H
    # Col H: Giờ ca đêm = SUMIFS(gc E:E, gc B:B, name, gc C:C, "*đêm*", gc A:A, cn)
    tot_gc_night = sum(h for (b, n, s, h) in gc_rows if b == cn and n == name and 'đêm' in s.lower())
    tot_gc_all = sum(h for (b, n, s, h) in gc_rows if b == cn and n == name)
    tot_gc_day = tot_gc_all - tot_gc_night
    
    col_j = round(tot_gc_day / 24.0, 1)
    
    # Col K: Ca đêm chuẩn (> 4h) = SUMIFS(nc I:I, nc A:A, cn, nc B:B, name)
    ca_dem_gt4 = sum(1 for item in nc_rows if item['br'] == cn and item['name'] == name and item['is_night'] and item['gt4'])
    
    # Col R: Công thực tế = unique days worked with shift > 4h
    days_gt4 = set(item['dt'] for item in nc_rows if item['br'] == cn and item['name'] == name and item['gt4'])
    cong_tt = len(days_gt4)
    
    print(f"{stt:<4} | {cn:<15} | {name:<24} | {tot_gc_day:<9.2f} | {tot_gc_night:<9.2f} | {col_j:<8.1f} | {ca_dem_gt4:<10} | {cong_tt:<10}")

wb.close()
