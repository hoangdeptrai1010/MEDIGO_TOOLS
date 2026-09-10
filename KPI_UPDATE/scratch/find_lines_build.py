import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('build_payroll_report.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for idx, line in enumerate(lines, 1):
    if any(k in line for k in ['Thưởng CK', 'Dự án', 'Dự án T8', 'Hot Bill']):
        print(f"Line {idx:4d}: {line.strip()}")
