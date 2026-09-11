from python_calamine import CalamineWorkbook
import sys, io, glob, os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

for f in [
    'd:/MEDIGO/KPI_UPDATE/thang9/data/phieunhap_tra/hoadon1092026.xlsx',
    'C:/Users/10102/Downloads/DanhSachChiTietHoaDon_KV10092026-105951-828.xlsx',
    'd:/MEDIGO/KPI_UPDATE/thang7/DATAKIOT/DanhSachChiTietHoaDon_KV03092026-134834-071.xlsx',
    'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
]:
    if not os.path.exists(f): continue
    wb = CalamineWorkbook.from_path(f)
    sheet = wb.get_sheet_by_name(wb.sheet_names[0])
    rows = sheet.to_python()
    hdr = rows[0]
    note_col = -1
    for idx, h in enumerate(hdr):
        if h and str(h).strip().lower() == 'ghi chú':
            note_col = idx
            break
    
    wa_bills = {}
    for r in rows[1:]:
        if note_col != -1 and note_col < len(r) and r[note_col]:
            note_val = str(r[note_col]).strip()
            if 'whatsapp' in note_val.lower() or 'whats app' in note_val.lower() or 'watsapp' in note_val.lower():
                code = r[1]
                if code not in wa_bills:
                    wa_bills[code] = {
                        'branch': r[0],
                        'code': code,
                        'time': str(r[6])[:19] if len(r)>6 else '',
                        'seller': r[20] if len(r)>20 else '',
                        'cust': r[12] if len(r)>12 else '',
                        'note': note_val,
                        'need_pay': float(r[43]) if len(r)>43 and r[43] is not None else 0.0
                    }
    print(f'File: {os.path.basename(f)} -> Found {len(wa_bills)} invoices with whatsapp in Ghi chú:')
    tot = sum(b['need_pay'] for b in wa_bills.values())
    print(f'  Total revenue: {tot:,.0f} đ')
    for code, b in list(wa_bills.items())[:10]:
        print(f"   {code} | {b['branch']} | {b['seller']} | {b['time']} | Note: {repr(b['note'])} | {b['need_pay']:,.0f} đ")
