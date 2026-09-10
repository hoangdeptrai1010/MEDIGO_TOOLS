import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

print("=== CHECKING HOW BUILDPAYROLL HANDLES THƯỞNG CK & DỰ ÁN ===")
with open('build_payroll_report.py', 'r', encoding='utf-8') as f:
    code = f.read()

for kw in ['Thưởng CK', 'Dự án', 'Dự án T8', 'thuong_ck', 'du_an']:
    count = code.count(kw)
    print(f"Keyword '{kw}' appears {count} times in build_payroll_report.py")

