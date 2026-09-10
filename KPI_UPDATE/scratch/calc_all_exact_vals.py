import zipfile, shutil, re, openpyxl

# Let's restore from bak first to have clean base
shutil.copy2('baocaokpi_thang8_hoanthien.xlsx.bak', 'baocaokpi_thang8_hoanthien.xlsx')

wb = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=True)
ws_da = wb['Dự án T8']
ws_ds = wb['kpi dược sĩ']

# Calculate exact new values for Du An T8
new_da_vals = {} # row -> {d, j, k, l, m}
for r in range(3, 54):
    b = ws_da.cell(r, 1).value
    name = ws_da.cell(r, 2).value
    if not name: continue
    
    ck = ws_da.cell(r, 6).value or 0
    cb = ws_da.cell(r, 7).value or 0
    hb = ws_da.cell(r, 12).value or 0
    
    ck_d = ck / 31.0
    cb_d = cb / 31.0
    ck_cb_d = (ck + cb) / 31.0
    
    is_hn = b in ['Hàng Bông', 'Đường Láng']
    new_j = 0
    if is_hn:
        if ck_d >= 3400000 and cb_d >= 900000: new_j = 4000000
        elif ck_d >= 3200000 and cb_d >= 750000: new_j = 2800000
        elif ck_d >= 2800000 and cb_d >= 650000: new_j = 2200000
        elif ck_d >= 2500000 and cb_d >= 520000: new_j = 1600000
        elif ck_d >= 2200000 and cb_d >= 450000: new_j = 1000000
    else:
        if ck_d >= 1500000 and cb_d >= 700000: new_j = 4000000
        elif ck_d >= 1200000 and cb_d >= 650000: new_j = 2800000
        elif ck_d >= 1000000 and cb_d >= 580000: new_j = 2200000
        elif ck_d >= 850000 and cb_d >= 480000: new_j = 1600000
        elif ck_d >= 630000 and cb_d >= 350000: new_j = 1000000
        
    new_k = 0
    if new_j == 0:
        min_extra = 2000000 if is_hn else 950000
        if ck_cb_d >= min_extra:
            new_k = 500000
            
    new_m = new_j + new_k + hb
    new_da_vals[r] = {
        'name': name,
        'd': ck_cb_d,
        'j': new_j,
        'k': new_k,
        'l': hb,
        'm': new_m
    }

print("=== CÁC DÒNG THAY ĐỔI TRONG DỰ ÁN T8 ===")
for r, d in new_da_vals.items():
    old_m = ws_da.cell(r, 13).value
    if d['m'] != old_m:
        print(f"Row {r:2d} ({d['name']:25s}): Old Total={old_m:,.0f} -> New Total={d['m']:,.0f} (Bậc={d['j']:,.0f}, Thêm={d['k']:,.0f}, HotBill={d['l']:,.0f})")

# Calculate new values for kpi dược sĩ
new_ds_vals = {} # row -> {m, n}
da_bonus_by_name = {d['name']: d['m'] for d in new_da_vals.values()}

for r in range(3, ws_ds.max_row+1):
    name = ws_ds.cell(r, 2).value
    if not name: continue
    
    kpi_bonus = ws_ds.cell(r, 11).value or 0
    he_so = ws_ds.cell(r, 12).value or 1
    if isinstance(kpi_bonus, str):
        try: kpi_bonus = float(kpi_bonus)
        except: kpi_bonus = 0.0
    if isinstance(he_so, str):
        try: he_so = float(he_so)
        except: he_so = 1.0
        
    da_tot = da_bonus_by_name.get(name, 0)
    new_m = da_tot * he_so
    new_n = kpi_bonus + new_m
    new_ds_vals[r] = {
        'name': name,
        'm': new_m,
        'n': new_n
    }
