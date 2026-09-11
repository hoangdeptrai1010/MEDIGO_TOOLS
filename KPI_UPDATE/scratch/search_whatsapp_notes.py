from python_calamine import CalamineWorkbook
import sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

files = [
    'C:/Users/10102/Downloads/DanhSachChiTietHoaDon_KV10092026-105951-828.xlsx',
    'd:/MEDIGO/KPI_UPDATE/thang9/data/phieunhap_tra/hoadon1092026.xlsx',
    'd:/MEDIGO/KPI_UPDATE/thang7/DATAKIOT/DanhSachChiTietHoaDon_KV03092026-134834-071.xlsx',
    'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
]

for f in files:
    try:
        wb = CalamineWorkbook.from_path(f)
        sheet = wb.get_sheet_by_name(wb.sheet_names[0])
        rows = sheet.to_python()
        hdr = rows[0]
        note_cols = [i for i, h in enumerate(hdr) if h and 'ghi chú' in str(h).strip().lower()]
        
        wa_invoices = {}
        for r in rows[1:]:
            inv_code = r[1] if len(r) > 1 else ''
            for nc in note_cols:
                if nc < len(r) and r[nc]:
                    val = str(r[nc]).strip()
                    v_low = val.lower()
                    if 'whatsapp' in v_low or 'whats app' in v_low or 'watsapp' in v_low or v_low == 'wa' or 'wa ' in v_low or ' wa' in v_low or 'ws' in v_low:
                        if inv_code not in wa_invoices:
                            wa_invoices[inv_code] = {
                                'branch': r[0],
                                'code': inv_code,
                                'time': r[6] if len(r) > 6 else (r[2] if len(r) > 2 else ''),
                                'seller': r[20] if len(r) > 20 else (r[7] if len(r) > 7 else ''),
                                'note': val,
                                'need_pay': float(r[43] if len(r) > 43 and r[43] is not None else (r[12] if len(r) > 12 and r[12] is not None else 0.0))
                            }
        print(f'File: {f}')
        print(f'  Total invoices with whatsapp note: {len(wa_invoices)}')
        total_rev = sum(inv['need_pay'] for inv in wa_invoices.values())
        print(f'  Total revenue: {total_rev:,.0f} đ')
        for k, v in list(wa_invoices.items())[:5]:
            print('    Sample:', v)
    except Exception as e:
        print(f'File: {f} Error: {e}')
