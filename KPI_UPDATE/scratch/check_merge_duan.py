import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=True)
ws = wb['Dự án T8']

all_staff = {}
for r in range(3, ws.max_row+1):
    cn = ws.cell(r, 1).value
    name = ws.cell(r, 2).value
    chuc_vu = ws.cell(r, 3).value
    try:
        ck = float(ws.cell(r, 6).value or 0)
    except:
        continue
    combo = float(ws.cell(r, 7).value or 0)
    hotbill = float(ws.cell(r, 12).value or 0)
    m_val = float(ws.cell(r, 13).value or 0)
    if not name or str(name).strip() == 'Medigo':
        continue
    n_str = str(name).strip()
    if n_str not in all_staff:
        all_staff[n_str] = {'main_cn': cn, 'branches': [], 'chuc_vu': chuc_vu, 'ck': 0.0, 'combo': 0.0, 'hotbill': 0.0, 'orig_total': 0.0}
    all_staff[n_str]['branches'].append((cn, ck, combo, m_val))
    all_staff[n_str]['ck'] += ck
    all_staff[n_str]['combo'] += combo
    all_staff[n_str]['hotbill'] += hotbill
    all_staff[n_str]['orig_total'] += m_val

print(f"{'STT':<4} | {'Họ và tên':<24} | {'Chi nhánh':<18} | {'Số CN':<5} | {'Thưởng cũ':<12} | {'Thưởng GỘP':<12} | {'Chênh lệch':<12}")
print("-" * 105)
stt = 0
for name, d in all_staff.items():
    avg_ck = d['ck'] / 31.0
    avg_combo = d['combo'] / 31.0
    avg_tot = (d['ck'] + d['combo']) / 31.0
    is_hn = any(b in str(d['main_cn']).lower() for b in ['hàng bông', 'đường láng'])
    
    bonus_moc = 0
    if is_hn:
        if avg_ck >= 3400000 and avg_combo >= 900000: bonus_moc = 4000000
        elif avg_ck >= 3200000 and avg_combo >= 750000: bonus_moc = 2800000
        elif avg_ck >= 2800000 and avg_combo >= 650000: bonus_moc = 2200000
        elif avg_ck >= 2500000 and avg_combo >= 520000: bonus_moc = 1600000
        elif avg_ck >= 2200000 and avg_combo >= 450000: bonus_moc = 1000000
    else:
        if avg_ck >= 1500000 and avg_combo >= 700000: bonus_moc = 4000000
        elif avg_ck >= 1200000 and avg_combo >= 650000: bonus_moc = 2800000
        elif avg_ck >= 1000000 and avg_combo >= 580000: bonus_moc = 2200000
        elif avg_ck >= 850000 and avg_combo >= 480000: bonus_moc = 1600000
        elif avg_ck >= 630000 and avg_combo >= 350000: bonus_moc = 1000000
        
    bonus_san = 0
    if bonus_moc == 0:
        threshold = 2000000 if is_hn else 950000
        if avg_tot >= threshold:
            bonus_san = 500000
            
    new_tot = bonus_moc + bonus_san + d['hotbill']
    diff = new_tot - d['orig_total']
    stt += 1
    flag = '*** TĂNG THƯỞNG ***' if diff > 0 else ('' if len(d['branches'])==1 else '(Đã gộp)')
    if len(d['branches']) > 1 or diff != 0:
        print(f"{stt:<4} | {name:<24} | {str(d['main_cn']):<18} | {len(d['branches']):<5} | {d['orig_total']:>12,.0f} | {new_tot:>12,.0f} | {diff:>12,.0f} {flag}")
