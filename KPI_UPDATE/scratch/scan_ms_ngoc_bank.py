import openpyxl
import os
import sys
import unicodedata

sys.stdout.reconfigure(encoding='utf-8')

def normalize_name(s):
    if not s:
        return ""
    # normalize unicode NFC
    s = unicodedata.normalize('NFC', str(s))
    # lower and trim extra whitespace
    return " ".join(s.lower().strip().split())

def clean_stk(stk):
    if stk is None:
        return ""
    stk_str = str(stk).strip()
    if stk_str in ["#N/A", "#REF!", "None", "nan", ""]:
        return ""
    # Remove dots, spaces, dashes
    stk_clean = re.sub(r'[\s\.\-_]', '', stk_str)
    # If float ending with .0
    if stk_clean.endswith(".0"):
        stk_clean = stk_clean[:-2]
    return stk_clean

import re

latvat_dir = r"d:\MEDIGO\KPI_UPDATE\thang8\latvat"
f_bangluong = os.path.join(latvat_dir, "BANGLUONGT8FIXV31 final.xlsx")
f_ngoc = os.path.join(latvat_dir, "Gởi Ms. Ngoc - 07.09. Bang Luong T8.2026.xlsx")

# 1. Inspect all sheets in Gởi Ms. Ngoc
wb_ngoc = openpyxl.load_workbook(f_ngoc, data_only=True)
master_bank = {}

for sname in wb_ngoc.sheetnames:
    ws = wb_ngoc[sname]
    print(f"\n--- Scanning Gởi Ms. Ngoc sheet '{sname}' ---")
    # find headers
    header_row = -1
    col_name = -1
    col_stk = -1
    col_bank = -1
    for r in range(1, min(15, ws.max_row + 1)):
        row_vals = [str(ws.cell(r, c).value or '').strip().lower() for c in range(1, ws.max_column + 1)]
        for idx, val in enumerate(row_vals, 1):
            if any(k in val for k in ['họ tên', 'họ và tên', 'tên chủ tk', 'tên chủ tài khoản', 'tên ctv', 'tên nhân viên']):
                if col_name == -1: col_name = idx
            if any(k in val for k in ['số tk', 'số tài khoản']):
                if col_stk == -1: col_stk = idx
            if any(k in val for k in ['ngân hàng', 'cn ngân hàng']):
                if col_bank == -1: col_bank = idx
        if col_name != -1 and (col_stk != -1 or col_bank != -1):
            header_row = r
            break
    
    print(f"Header row: {header_row}, col_name: {col_name}, col_stk: {col_stk}, col_bank: {col_bank}")
    if header_row != -1:
        count = 0
        for r in range(header_row + 1, ws.max_row + 1):
            name = ws.cell(r, col_name).value
            stk = ws.cell(r, col_stk).value if col_stk != -1 else None
            bank = ws.cell(r, col_bank).value if col_bank != -1 else None
            if name and str(name).strip() and not str(name).strip().startswith("Tổng"):
                norm = normalize_name(name)
                c_stk = clean_stk(stk)
                c_bank = str(bank).strip() if bank else ""
                if norm and (c_stk or c_bank):
                    count += 1
                    if norm not in master_bank:
                        master_bank[norm] = []
                    master_bank[norm].append({
                        'source': f"MsNgoc_{sname}_r{r}",
                        'raw_name': name,
                        'stk': c_stk,
                        'bank': c_bank
                    })
        print(f"Extracted {count} rows from {sname}")

print(f"\nTotal unique names in master_bank: {len(master_bank)}")
