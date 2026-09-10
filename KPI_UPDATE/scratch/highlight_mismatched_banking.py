import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
import shutil
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

latvat_dir = r"d:\MEDIGO\KPI_UPDATE\thang8\latvat"
f_bl = os.path.join(latvat_dir, "BANGLUONGT8FIXV31 final.xlsx")
f_backup = os.path.join(latvat_dir, "BANGLUONGT8FIXV31 final_backup.xlsx")
f_ngoc = os.path.join(latvat_dir, "Gởi Ms. Ngoc - 07.09. Bang Luong T8.2026.xlsx")

# Create backup if not exists
if not os.path.exists(f_backup):
    shutil.copy2(f_bl, f_backup)
    print(f"Created backup: {f_backup}")

# Load master dictionary from Gởi Ms. Ngoc and reference sheets
wb_ngoc = openpyxl.load_workbook(f_ngoc, data_only=True, read_only=True)
master_records = {}

for sname in ["TTCK CNT", "TT CK CNT", "TTCK VP ", "Med", "Medigo ", "DSDB"]:
    if sname not in wb_ngoc.sheetnames:
        continue
    ws = wb_ngoc[sname]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = -1
    c_name, c_stk, c_bank = -1, -1, -1
    for r_i, row in enumerate(rows[:15]):
        for c_i, val in enumerate(row):
            v_str = str(val or '').strip().lower()
            if any(k == v_str or v_str.startswith(k) for k in ['họ tên', 'họ và tên', 'tên chủ tk', 'tên chủ tài khoản', 'tên ctv', 'tên nhân viên', 'họ và tên nhân viên']):
                if c_name == -1: c_name = c_i
            if any(k in v_str for k in ['số tk', 'số tài khoản']):
                if c_stk == -1: c_stk = c_i
            if any(k in v_str for k in ['ngân hàng', 'cn ngân hàng']):
                if c_bank == -1: c_bank = c_i
        if c_name != -1 and (c_stk != -1 or c_bank != -1):
            header_idx = r_i
            break
            
    if header_idx != -1:
        for row in rows[header_idx + 1:]:
            if c_name < len(row):
                name = row[c_name]
                stk = row[c_stk] if c_stk != -1 and c_stk < len(row) else None
                bank = row[c_bank] if c_bank != -1 and c_bank < len(row) else None
                if name and str(name).strip() and not str(name).strip().startswith("Tổng"):
                    c_s = clean_stk(stk)
                    c_b = str(bank).strip() if bank and str(bank).strip() not in ["#N/A", "#REF!"] else ""
                    if c_s or c_b:
                        norm = normalize_name(name)
                        norm_noacc = normalize_name(remove_accents(name))
                        for k in [norm, norm_noacc]:
                            if k not in master_records:
                                master_records[k] = {'stk': c_s, 'bank': c_b, 'sources': []}
                            if c_s and not master_records[k]['stk']:
                                master_records[k]['stk'] = c_s
                            if c_b and not master_records[k]['bank']:
                                master_records[k]['bank'] = c_b
                            master_records[k]['sources'].append(f"{sname}: {name} | {c_s} | {c_b}")

print(f"Loaded master dictionary with {len(master_records)} keys.")

# Open BANGLUONGT8FIXV31 final.xlsx in read/write mode (preserving formatting/formulas)
wb_bl = openpyxl.load_workbook(f_bl)

# Define Red highlight styling
red_fill = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid") # Soft Excel warning red
red_font = Font(name="Calibri", size=11, bold=True, color="FF9C0006")
solid_red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid") # Strong red
solid_red_font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")

# We will use solid bright red highlight so it is clearly visible to the user
highlight_fill = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")
highlight_font = Font(name="Calibri", size=11, bold=True, color="FFCC0000")

highlight_count_s1 = 0
highlight_count_med = 0

print("\n=======================================================")
print("PROCESSING Sheet1 in BANGLUONGT8FIXV31 final.xlsx")
print("=======================================================")
ws_s1 = wb_bl["Sheet1"]
for r in range(2, ws_s1.max_row + 1):
    c_name = ws_s1.cell(r, 3) # Col C: Họ và tên
    c_acc = ws_s1.cell(r, 4)  # Col D: Tên chủ tài khoản
    c_stk = ws_s1.cell(r, 5)  # Col E: Số TK
    c_bank = ws_s1.cell(r, 6) # Col F: Ngân hàng
    
    name_val = c_name.value
    acc_val = c_acc.value
    stk_val = c_stk.value
    bank_val = c_bank.value
    
    if not name_val and not stk_val and not bank_val:
        continue
        
    c_s = clean_stk(stk_val)
    c_b = str(bank_val).strip() if bank_val else ""
    norm_n = normalize_name(name_val)
    norm_acc = normalize_name(acc_val)
    
    ref = master_records.get(norm_n) or master_records.get(norm_acc)
    ref_stk = ref['stk'] if ref else ""
    ref_bank = ref['bank'] if ref else ""
    
    stk_ok = (c_s == ref_stk and c_s != "") if ref and ref_stk else (c_s != "")
    bank_ok = False
    if ref and ref_bank:
        bank_ok = (remove_accents(c_b).lower() in remove_accents(ref_bank).lower() or remove_accents(ref_bank).lower() in remove_accents(c_b).lower()) and c_b not in ["#N/A", "#REF!", ""]
    else:
        bank_ok = (c_b not in ["#N/A", "#REF!", ""])
        
    if not stk_ok:
        c_stk.fill = highlight_fill
        c_stk.font = highlight_font
        highlight_count_s1 += 1
        print(f"Sheet1 [Row {r:2d}] STK ERROR -> Name: {name_val!r:<25} | Cur STK: {str(stk_val):<15} | Master STK: {ref_stk:<15}")
        
    if not bank_ok:
        c_bank.fill = highlight_fill
        c_bank.font = highlight_font
        highlight_count_s1 += 1
        print(f"Sheet1 [Row {r:2d}] BANK ERROR -> Name: {name_val!r:<25} | Cur Bank: {str(bank_val):<15} | Master Bank: {ref_bank:<15}")

print(f"\nTotal highlighted cells in Sheet1: {highlight_count_s1}")

print("\n=======================================================")
print("PROCESSING Med in BANGLUONGT8FIXV31 final.xlsx")
print("=======================================================")
ws_med = wb_bl["Med"]
for r in range(10, ws_med.max_row + 1):
    c_code = ws_med.cell(r, 3) # Col C: Mã NV
    c_name = ws_med.cell(r, 4) # Col D: Họ tên
    c_stk = ws_med.cell(r, 9)  # Col I: Số tài khoản
    c_bank = ws_med.cell(r, 10) # Col J: CN Ngân hàng
    
    name_val = c_name.value
    code_val = c_code.value
    stk_val = c_stk.value
    bank_val = c_bank.value
    
    if not name_val or str(name_val).strip() == "" or str(name_val).strip().startswith("CHUỖI") or str(name_val).strip().startswith("Tổng"):
        continue
        
    c_s = clean_stk(stk_val)
    c_b = str(bank_val).strip() if bank_val else ""
    norm_n = normalize_name(name_val)
    norm_c = normalize_name(code_val)
    
    ref = master_records.get(norm_n) or master_records.get(norm_c)
    ref_stk = ref['stk'] if ref else ""
    ref_bank = ref['bank'] if ref else ""
    
    stk_ok = (c_s == ref_stk and c_s != "") if ref and ref_stk else (c_s != "")
    bank_ok = False
    if ref and ref_bank:
        bank_ok = (remove_accents(c_b).lower() in remove_accents(ref_bank).lower() or remove_accents(ref_bank).lower() in remove_accents(c_b).lower()) and c_b not in ["#N/A", "#REF!", ""]
    else:
        bank_ok = (c_b not in ["#N/A", "#REF!", ""])
        
    if not stk_ok:
        c_stk.fill = highlight_fill
        c_stk.font = highlight_font
        highlight_count_med += 1
        print(f"Med [Row {r:3d}] STK ERROR -> Name: {name_val!r:<25} | Cur STK: {str(stk_val):<15} | Master STK: {ref_stk:<15}")
        
    if not bank_ok:
        c_bank.fill = highlight_fill
        c_bank.font = highlight_font
        highlight_count_med += 1
        print(f"Med [Row {r:3d}] BANK ERROR -> Name: {name_val!r:<25} | Cur Bank: {str(bank_val):<15} | Master Bank: {ref_bank:<15}")

print(f"\nTotal highlighted cells in Med: {highlight_count_med}")

# Save the updated workbook
wb_bl.save(f_bl)
print(f"\nSaved highlighted workbook successfully to: {f_bl}")
