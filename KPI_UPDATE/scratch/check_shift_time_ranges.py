# -*- coding: utf-8 -*-
import openpyxl, sys, re
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/BANGLUONGTHANG8.xlsx', data_only=True)
ws_gc = wb['Giờ công']

print("=== CHECKING ALL SHIFT DEFINITIONS IN GIỜ CÔNG ===")
for r in range(2, ws_gc.max_row+1):
    c_name = ws_gc.cell(r, 2).value
    c_ca = str(ws_gc.cell(r, 3).value or '').strip()
    c_raw = ws_gc.cell(r, 4).value
    c_hours = ws_gc.cell(r, 5).value
    
    if not c_ca or c_ca == '-':
        continue
        
    # Check if shift time is in 06:00-22:00 vs 22:00-06:00
    # Search for (HH:MM - HH:MM)
    m = re.search(r'(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})', c_ca)
    if m:
        start_h = int(m.group(1))
        end_h = int(m.group(3))
        
        # If start_h >= 22 or start_h < 6: night shift
        # If start_h >= 6 and end_h <= 22: day shift
        is_night_time = (start_h >= 22 or start_h < 6 or 'qua đêm' in c_ca.lower())
        is_day_time = (start_h >= 6 and (end_h <= 22 or end_h > start_h) and not is_night_time)
        
        # Check current classification in Giờ công (current logic uses '*qua đêm*')
        classified_as_night = ('qua đêm' in c_ca.lower())
        
        if is_night_time != classified_as_night:
            print(f"MISMATCH Row {r:3d}: Name={c_name} | Shift={c_ca} | Time={start_h}h->{end_h}h | is_night={is_night_time} vs classified_night={classified_as_night}")

print("\n=== ALL MISMATCH CHECKS COMPLETED ===")
