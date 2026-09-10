import sys, os, openpyxl, subprocess

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Run the build_payroll_report with month 8
cmd = [
    r"C:\Users\10102\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe",
    "build_payroll_report.py",
    "--month", "8"
]
print("Running command:", " ".join(cmd))
res = subprocess.run(cmd, capture_output=True, text=True)
print("STDOUT:\n", res.stdout)
if res.stderr:
    print("STDERR:\n", res.stderr)
