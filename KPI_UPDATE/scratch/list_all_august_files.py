import os

root_dir = 'thang8/Pharmacy_retail_Store_KPIs_August 2026'
for r, d, files in os.walk(root_dir):
    for f in files:
        full = os.path.join(r, f)
        print(f"{os.path.splitext(f)[1]} | {f} | {full}")
