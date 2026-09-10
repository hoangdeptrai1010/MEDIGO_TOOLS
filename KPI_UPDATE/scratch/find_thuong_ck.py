import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('build_payroll_report.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for idx, line in enumerate(lines, 1):
    if 'THUONG_CK' in line or 'thuong_ck' in line.lower():
        print(f"Line {idx:4d}: {line.strip()}")
