import openpyxl
import os
import sys
import unicodedata
import re

sys.stdout.reconfigure(encoding='utf-8')

def remove_accents(input_str):
    if not input_str:
        return ""
    nfkd_form = unicodedata.normalize('NFKD', str(input_str))
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def normalize_name(s):
    if not s:
        return ""
    s = unicodedata.normalize('NFC', str(s))
    return " ".join(s.lower().strip().split())

def clean_stk(stk):
    if stk is None:
        return ""
    stk_str = str(stk).strip()
    if stk_str in ["#N/A", "#REF!", "None", "nan", ""]:
        return ""
    stk_clean = re.sub(r'[\s\.\-_]', '', stk_str)
    if stk_clean.endswith(".0"):
        stk_clean = stk_clean[:-2]
    return stk_clean

def normalize_bank(b):
    if not b:
        return ""
    b_str = str(b).strip()
    if b_str in ["#N/A", "#REF!", "None", "nan", ""]:
        return ""
    # Simplify common bank name variations
    b_norm = remove_accents(b_str).lower()
    b_norm = re.sub(r'[^a-z0-9]', '', b_norm)
    return b_norm

# Let's inspect all rows of Sheet1 and Med in BANGLUONGT8FIXV31 final.xlsx
latvat_dir = r"d:\MEDIGO\KPI_UPDATE\thang8\latvat"
f_bl = os.path.join(latvat_dir, "BANGLUONGT8FIXV31 final.xlsx")
wb_bl = openpyxl.load_workbook(f_bl, data_only=True)

# Build a comprehensive reference dictionary from multiple sources:
# 1. Gởi Ms. Ngoc - 07.09. Bang Luong T8.2026.xlsx
# 2. 2chiem- Gởi Ms. Ngoc - 07.09. Bang Luong T8.2026.xlsx
# 3. Bang luong tong T8-2026 - CNT.xlsx
# 4. danhsachnhanvien.xlsx

ref_files = [
    os.path.join(latvat_dir, "Gởi Ms. Ngoc - 07.09. Bang Luong T8.2026.xlsx"),
    r"d:\MEDIGO\KPI_UPDATE\thang8\FIXFORMAT\2chiem- Gởi Ms. Ngoc - 07.09. Bang Luong T8.2026.xlsx",
    r"d:\MEDIGO\KPI_UPDATE\thang8\FIXFORMAT\Bang luong tong T8-2026 - CNT.xlsx",
    r"d:\MEDIGO\KPI_UPDATE\thang8\bangluong_thang8_hoanthien.xlsx",
    r"d:\MEDIGO\KPI_UPDATE\thang8\BANGLUONGTHANG8.xlsx",
]

master_dict = {} # key: norm_name -> list of records

for fpath in ref_files:
    if not os.path.exists(fpath):
        continue
    fname = os.path.basename(fpath)
    print(f"Loading reference: {fname}")
    wb_ref = openpyxl.load_workbook(fpath, data_only=True)
    for sname in wb_ref.sheetnames:
        ws = wb_ref[sname]
        # scan for columns
        for r in range(1, min(15, ws.max_row + 1)):
            for c_name_idx in range(1, min(10, ws.max_column + 1)):
                val = str(ws.cell(r, c_name_idx).value or '').strip().lower()
                if any(k == val or val.startswith(k) for k in ['họ tên', 'họ và tên', 'tên chủ tk', 'tên chủ tài khoản', 'tên ctv', 'tên nhân viên']):
                    # look for STK and Bank in same row or header
                    for c_stk_idx in range(1, ws.max_column + 1):
                        val_stk = str(ws.cell(r, c_stk_idx).value or '').strip().lower()
                        if any(k in val_stk for k in ['số tk', 'số tài khoản']):
                            for c_bank_idx in range(1, ws.max_column + 1):
                                val_bank = str(ws.cell(r, c_bank_idx).value or '').strip().lower()
                                if any(k in val_bank for k in ['ngân hàng', 'cn ngân hàng']):
                                    # iterate data rows
                                    for data_r in range(r + 1, ws.max_row + 1):
                                        n = ws.cell(data_r, c_name_idx).value
                                        s = ws.cell(data_r, c_stk_idx).value
                                        b = ws.cell(data_r, c_bank_idx).value
                                        if n and str(n).strip() and not str(n).strip().startswith("Tổng"):
                                            norm_n = normalize_name(n)
                                            no_acc_n = normalize_name(remove_accents(n))
                                            c_s = clean_stk(s)
                                            c_b = str(b).strip() if b else ""
                                            if c_s or c_b:
                                                entry = {
                                                    'file': fname,
                                                    'sheet': sname,
                                                    'raw_name': str(n).strip(),
                                                    'stk': c_s,
                                                    'bank': c_b,
                                                    'bank_norm': normalize_bank(c_b)
                                                }
                                                for key in [norm_n, no_acc_n]:
                                                    if key not in master_dict:
                                                        master_dict[key] = []
                                                    master_dict[key].append(entry)

print(f"\nBuilt master dictionary with {len(master_dict)} name variants.")

# Now check Sheet1 in BANGLUONGT8FIXV31 final.xlsx
print("\n=======================================================")
print("AUDITING Sheet1 in BANGLUONGT8FIXV31 final.xlsx")
print("=======================================================")
ws_s1 = wb_bl["Sheet1"]
for r in range(2, ws_s1.max_row + 1):
    name = ws_s1.cell(r, 3).value # Col C: Họ và tên
    acc_holder = ws_s1.cell(r, 4).value # Col D: Tên chủ tài khoản
    stk = ws_s1.cell(r, 5).value # Col E: Số TK
    bank = ws_s1.cell(r, 6).value # Col F: Ngân hàng
    
    if not name and not stk and not bank:
        continue
        
    c_stk = clean_stk(stk)
    c_bank = str(bank).strip() if bank else ""
    norm_n = normalize_name(name)
    norm_acc = normalize_name(acc_holder)
    
    ref_entries = master_dict.get(norm_n, []) or master_dict.get(norm_acc, [])
    
    status_stk = "OK"
    status_bank = "OK"
    expected_stk = ""
    expected_bank = ""
    
    if not c_stk:
        status_stk = "MISSING/INVALID"
    if not c_bank or c_bank in ["#N/A", "#REF!"]:
        status_bank = "MISSING/INVALID"
        
    if ref_entries:
        # find matching
        valid_refs = [e for e in ref_entries if e['stk']]
        if valid_refs:
            expected_stk = valid_refs[0]['stk']
            expected_bank = valid_refs[0]['bank']
            if c_stk and c_stk != expected_stk:
                status_stk = f"MISMATCH (Got: {c_stk}, Ref: {expected_stk})"
            if c_bank and normalize_bank(c_bank) != normalize_bank(expected_bank):
                # check if substring or reasonable match
                if normalize_bank(c_bank) not in normalize_bank(expected_bank) and normalize_bank(expected_bank) not in normalize_bank(c_bank):
                    status_bank = f"MISMATCH (Got: {c_bank}, Ref: {expected_bank})"
    else:
        if status_stk == "OK":
            status_stk = "NO_REF_FOUND"
        if status_bank == "OK":
            status_bank = "NO_REF_FOUND"

    print(f"Row {r:2d} | Name: {str(name):<28} | STK: {str(stk):<20} | Bank: {str(bank):<22} | STK_STAT: {status_stk} | BANK_STAT: {status_bank} | REF: {expected_stk} / {expected_bank}")

# Now check Med in BANGLUONGT8FIXV31 final.xlsx
print("\n=======================================================")
print("AUDITING Med in BANGLUONGT8FIXV31 final.xlsx")
print("=======================================================")
ws_med = wb_bl["Med"]
for r in range(10, ws_med.max_row + 1):
    code = ws_med.cell(r, 3).value
    name = ws_med.cell(r, 4).value
    stk = ws_med.cell(r, 9).value
    bank = ws_med.cell(r, 10).value
    
    if not name or str(name).strip() == "" or str(name).strip().startswith("CHUỖI") or str(name).strip().startswith("Tổng"):
        continue
        
    c_stk = clean_stk(stk)
    c_bank = str(bank).strip() if bank else ""
    norm_n = normalize_name(name)
    norm_code = normalize_name(code)
    
    ref_entries = master_dict.get(norm_n, []) or master_dict.get(norm_code, [])
    
    status_stk = "OK"
    status_bank = "OK"
    expected_stk = ""
    expected_bank = ""
    
    if not c_stk:
        status_stk = "MISSING/INVALID"
    if not c_bank or c_bank in ["#N/A", "#REF!"]:
        status_bank = "MISSING/INVALID"
        
    if ref_entries:
        valid_refs = [e for e in ref_entries if e['stk']]
        if valid_refs:
            expected_stk = valid_refs[0]['stk']
            expected_bank = valid_refs[0]['bank']
            if c_stk and c_stk != expected_stk:
                status_stk = f"MISMATCH (Got: {c_stk}, Ref: {expected_stk})"
            if c_bank and normalize_bank(c_bank) != normalize_bank(expected_bank):
                if normalize_bank(c_bank) not in normalize_bank(expected_bank) and normalize_bank(expected_bank) not in normalize_bank(c_bank):
                    status_bank = f"MISMATCH (Got: {c_bank}, Ref: {expected_bank})"
    else:
        if status_stk == "OK":
            status_stk = "NO_REF_FOUND"
        if status_bank == "OK":
            status_bank = "NO_REF_FOUND"

    if status_stk != "OK" or status_bank != "OK":
        print(f"Row {r:3d} | Name: {str(name):<28} | STK: {str(stk):<20} | Bank: {str(bank):<22} | STK_STAT: {status_stk} | BANK_STAT: {status_bank} | REF: {expected_stk} / {expected_bank}")
