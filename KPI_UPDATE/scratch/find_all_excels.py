import openpyxl
import os
import sys
import glob
import re

sys.stdout.reconfigure(encoding='utf-8')

workspace = r"d:\MEDIGO\KPI_UPDATE"

# Search for all excel files that might contain bank info
excel_files = glob.glob(os.path.join(workspace, "**/*.xlsx"), recursive=True)
print(f"Found {len(excel_files)} excel files:")
for f in excel_files:
    if "~$" not in f:
        print(" -", os.path.relpath(f, workspace))
