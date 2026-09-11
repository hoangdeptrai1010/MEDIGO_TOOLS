import openpyxl, sys, io, glob, os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("=== SEARCHING ALL EXCEL FILES FOR 'WHATSAPP' ===")

for p in glob.glob('d:/MEDIGO/**/*.xlsx', recursive=True):
    if '~$' in p: continue
    try:
        wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
        found_in_sheets = []
        for sname in wb.sheetnames:
            if 'whatsapp' in sname.lower():
                found_in_sheets.append(f"SheetName: '{sname}'")
        if found_in_sheets:
            print(f"File: {p} -> {', '.join(found_in_sheets)}")
        wb.close()
    except Exception:
        pass
