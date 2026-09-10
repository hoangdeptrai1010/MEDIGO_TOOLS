import openpyxl, sys, datetime
from collections import defaultdict

p = r'C:\Users\10102\Downloads\DanhSachChiTietHoaDon_KV03092026-140832-762.xlsx'
wb = openpyxl.load_workbook(p, read_only=True)
ws = wb.active

print('Scanning rows in', p)
dates = set()
wa_bills = {} # bill_code -> {branch, time, seller, note, total_amount, items: []}

r_idx = 0
for row in ws.iter_rows(values_only=True):
    r_idx += 1
    if r_idx == 1:
        continue
    
    time_val = row[6]
    if time_val:
        dates.add(str(time_val)[:10])
        
    note_val = row[37]
    item_note_val = row[62]
    
    # Check both invoice note and item note
    has_wa = False
    match_txt = ''
    for note in [note_val, item_note_val]:
        if note:
            s = str(note).lower()
            if 'whatsapp' in s or 'khách q' in s or 'khach q' in s or 'quốc tế' in s or 'quoc te' in s:
                has_wa = True
                match_txt = str(note)
                break
                
    if has_wa:
        bill_code = row[1]
        if bill_code not in wa_bills:
            wa_bills[bill_code] = {
                'branch': row[0],
                'bill': bill_code,
                'time': row[6],
                'seller': row[20],
                'note': match_txt,
                'total_need_pay': row[43],
                'items_rev': 0.0,
                'item_count': 0
            }
        item_rev = row[68] # Thanh tien
        if item_rev is not None:
            wa_bills[bill_code]['items_rev'] += float(item_rev)
        wa_bills[bill_code]['item_count'] += 1

print(f'Total rows scanned: {r_idx}')
sorted_dates = sorted(list(dates))
if sorted_dates:
    print(f'Date range: {sorted_dates[0]} to {sorted_dates[-1]} (Total {len(sorted_dates)} unique dates)')

print(f'\nTotal WhatsApp bills found: {len(wa_bills)}')

# Aggregate by Branch and Seller
by_branch_seller = defaultdict(lambda: {'bills': 0, 'rev': 0.0, 'notes': []})
for b_code, b_data in wa_bills.items():
    key = (b_data['branch'], b_data['seller'])
    by_branch_seller[key]['bills'] += 1
    rev = b_data['total_need_pay'] if b_data['total_need_pay'] is not None else b_data['items_rev']
    by_branch_seller[key]['rev'] += float(rev)
    by_branch_seller[key]['notes'].append(b_data['note'])

print('\n=== Breakdown by Branch & Seller ===')
for (branch, seller), data in sorted(by_branch_seller.items()):
    b_count = data['bills']
    b_rev = data['rev']
    print(f'Branch: {branch} | Seller: {seller} | Bills: {b_count} | Total Rev: {b_rev:,.0f}đ')

print('\n=== Sample Notes ===')
for b_code, b_data in list(wa_bills.items())[:20]:
    print(f"[{b_data['branch']}] [{b_data['seller']}] Bill {b_code} ({b_data['time']}): Note: {repr(b_data['note'])}")
