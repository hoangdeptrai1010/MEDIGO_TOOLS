import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

for f in ['thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/Dự án/Chương trình/CK-HN.xlsx',
          'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hồ Chí Minh/Dự án/Chương trình/CK-HCM.xlsx']:
    print(f"\n==================== {f} ====================")
    wb = openpyxl.load_workbook(f, data_only=True)
    for s in wb.sheetnames:
        ws = wb[s]
        print(f"--- Sheet: {s} ---")
        for r in range(1, 10):
            vals = [ws.cell(r, c).value for c in range(1, 15)]
            if any(v is not None for v in vals):
                print(f"Row {r:2d}: {vals}")

# Read PPTX files if python-pptx is available or search existing scripts
try:
    from pptx import Presentation
    for p in ['thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/Dự án/Chương trình/Hanoi_Projects_August_2026_v1.pptx',
              'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hồ Chí Minh/Dự án/Chương trình/HCM_Projects_August_2026_v1.pptx']:
        print(f"\n==================== PPTX: {p} ====================")
        prs = Presentation(p)
        for idx, slide in enumerate(prs.slides, 1):
            print(f"--- Slide {idx} ---")
            for shape in slide.shapes:
                if shape.has_text_frame:
                    print(shape.text_frame.text)
except Exception as e:
    print(f"PPTX error: {e}")
