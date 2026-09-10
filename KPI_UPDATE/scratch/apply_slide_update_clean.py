import zipfile, shutil, re, openpyxl

# Start from clean bak
shutil.copy2('baocaokpi_thang8_hoanthien.xlsx.bak', 'baocaokpi_thang8_hoanthien.xlsx')

wb = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=True)
ws_da = wb['Dự án T8']
ws_ds = wb['kpi dược sĩ']

# Precompute Du An T8 values
da_vals = {}
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
    j_val = 0
    if is_hn:
        if ck_d >= 3400000 and cb_d >= 900000: j_val = 4000000
        elif ck_d >= 3200000 and cb_d >= 750000: j_val = 2800000
        elif ck_d >= 2800000 and cb_d >= 650000: j_val = 2200000
        elif ck_d >= 2500000 and cb_d >= 520000: j_val = 1600000
        elif ck_d >= 2200000 and cb_d >= 450000: j_val = 1000000
    else:
        if ck_d >= 1500000 and cb_d >= 700000: j_val = 4000000
        elif ck_d >= 1200000 and cb_d >= 650000: j_val = 2800000
        elif ck_d >= 1000000 and cb_d >= 580000: j_val = 2200000
        elif ck_d >= 850000 and cb_d >= 480000: j_val = 1600000
        elif ck_d >= 630000 and cb_d >= 350000: j_val = 1000000
        
    k_val = 0
    if j_val == 0:
        min_extra = 2000000 if is_hn else 950000
        if ck_cb_d >= min_extra:
            k_val = 500000
            
    m_val = j_val + k_val + hb
    da_vals[r] = {
        'name': name,
        'd': ck_cb_d,
        'j': j_val,
        'k': k_val,
        'l': hb,
        'm': m_val
    }

# Precompute KPI Duoc Si values
da_by_name = {d['name']: d['m'] for d in da_vals.values()}
ds_vals = {}
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
    
    da_tot = da_by_name.get(name, 0)
    m_val = da_tot * he_so
    n_val = kpi_bonus + m_val
    ds_vals[r] = {
        'name': name,
        'm': m_val,
        'n': n_val
    }

wb.close()

# Process XML zip
src_file = 'baocaokpi_thang8_hoanthien.xlsx'
z_in = zipfile.ZipFile(src_file, 'r')
temp_zip = 'temp_clean.xlsx'
z_out = zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED)

new_j3_formula = (
    'IFERROR('
    'IF(OR($A3=&quot;Hàng Bông&quot;,$A3=&quot;Đường Láng&quot;),'
    '_xlfn.IFS('
    'AND($H3&gt;=3400000,$I3&gt;=900000),4000000,'
    'AND($H3&gt;=3200000,$I3&gt;=750000),2800000,'
    'AND($H3&gt;=2800000,$I3&gt;=650000),2200000,'
    'AND($H3&gt;=2500000,$I3&gt;=520000),1600000,'
    'AND($H3&gt;=2200000,$I3&gt;=450000),1000000,'
    'TRUE,0),'
    '_xlfn.IFS('
    'AND($H3&gt;=1500000,$I3&gt;=700000),4000000,'
    'AND($H3&gt;=1200000,$I3&gt;=650000),2800000,'
    'AND($H3&gt;=1000000,$I3&gt;=580000),2200000,'
    'AND($H3&gt;=850000,$I3&gt;=480000),1600000,'
    'AND($H3&gt;=630000,$I3&gt;=350000),1000000,'
    'TRUE,0)),'
    '0)'
)

for item in z_in.infolist():
    content = z_in.read(item.filename)
    
    if item.filename == 'xl/worksheets/sheet4.xml':
        s4 = content.decode('utf-8')
        
        # 1. Col D formula: SUM(F3:G3)/DAY($A$1)
        s4 = s4.replace('SUM(E3:G3)/DAY($A$1)', 'SUM(F3:G3)/DAY($A$1)')
        
        # 2. Col J formula: new_j3_formula
        s4 = re.sub(r'<f t="shared" ref="J3:J53" si="4">.*?</f>', f'<f t="shared" ref="J3:J53" si="4">{new_j3_formula}</f>', s4, flags=re.DOTALL)
        
        # 3. Update cached values for sheet 4
        for r, vals in da_vals.items():
            # Update D
            d_val_str = f"{vals['d']:.3f}"
            s4 = re.sub(rf'(<c r="D{r}"[^>]*><f[^>]*/>)<v>[^<]*</v>', rf'\g<1><v>{d_val_str}</v>', s4)
            s4 = re.sub(rf'(<c r="D{r}"[^>]*><f t="shared" ref="D3:D53" si="2">[^<]*</f>)<v>[^<]*</v>', rf'\g<1><v>{d_val_str}</v>', s4)
            
            # Update J
            j_val_str = f"{vals['j']}"
            s4 = re.sub(rf'(<c r="J{r}"[^>]*><f[^>]*/>)<v>[^<]*</v>', rf'\g<1><v>{j_val_str}</v>', s4)
            s4 = re.sub(rf'(<c r="J{r}"[^>]*><f t="shared" ref="J3:J53" si="4">[^<]*</f>)<v>[^<]*</v>', rf'\g<1><v>{j_val_str}</v>', s4)
            
            # Update K
            k_val_str = f"{vals['k']}"
            s4 = re.sub(rf'(<c r="K{r}"[^>]*><f[^>]*/>)<v>[^<]*</v>', rf'\g<1><v>{k_val_str}</v>', s4)
            s4 = re.sub(rf'(<c r="K{r}"[^>]*><f t="shared" ref="K3:K53" si="5">[^<]*</f>)<v>[^<]*</v>', rf'\g<1><v>{k_val_str}</v>', s4)
            
            # Update M
            m_val_str = f"{vals['m']}"
            s4 = re.sub(rf'(<c r="M{r}"[^>]*><f[^>]*/>)<v>[^<]*</v>', rf'\g<1><v>{m_val_str}</v>', s4)
            s4 = re.sub(rf'(<c r="M{r}"[^>]*><f t="shared" ref="M3:M53" si="6">[^<]*</f>)<v>[^<]*</v>', rf'\g<1><v>{m_val_str}</v>', s4)
            
        content = s4.encode('utf-8')
        
    elif item.filename == 'xl/worksheets/sheet2.xml':
        s2 = content.decode('utf-8')
        
        # 1. Fix IFS in Col K: prefix with _xlfn.
        s2 = s2.replace('IFS(\n            AND(I', '_xlfn.IFS(\n            AND(I')
        s2 = s2.replace('IFS(AND(I', '_xlfn.IFS(AND(I')
        
        # 2. Update cached values for M and N
        for r, vals in ds_vals.items():
            m_val_str = f"{vals['m']}"
            n_val_str = f"{vals['n']:.3f}" if isinstance(vals['n'], float) else f"{vals['n']}"
            
            # Update M
            s2 = re.sub(rf'(<c r="M{r}"[^>]*><f[^>]*>[^<]*</f>)<v>[^<]*</v>', rf'\g<1><v>{m_val_str}</v>', s2)
            
            # Update N
            s2 = re.sub(rf'(<c r="N{r}"[^>]*><f[^>]*/>)<v>[^<]*</v>', rf'\g<1><v>{n_val_str}</v>', s2)
            s2 = re.sub(rf'(<c r="N{r}"[^>]*><f t="shared" ref="N3:N55" si="7">[^<]*</f>)<v>[^<]*</v>', rf'\g<1><v>{n_val_str}</v>', s2)
            
        content = s2.encode('utf-8')

    z_out.writestr(item, content)

z_in.close()
z_out.close()

shutil.move(temp_zip, src_file)
shutil.copy2(src_file, 'thang8/baocaokpi_thang8_hoanthien.xlsx')
print("✅ Hoàn thành cập nhật sạch cả 2 file báo cáo!")
