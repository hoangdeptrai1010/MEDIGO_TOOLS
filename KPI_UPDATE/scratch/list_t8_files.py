import sys, os, glob
sys.stdout.reconfigure(encoding='utf-8')

print("=== CHECKING ALL FILES IN thang8/Pharmacy_retail_Store_KPIs_August 2026 ===")
for root, dirs, files in os.walk('thang8'):
    for f in files:
        p = os.path.join(root, f)
        print(f"File: {p} ({os.path.getsize(p)} bytes)")

