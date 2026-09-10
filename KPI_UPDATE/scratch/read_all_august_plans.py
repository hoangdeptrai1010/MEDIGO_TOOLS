import os, zipfile, openpyxl, urllib.parse
import xml.etree.ElementTree as ET

base_dir = r"d:\MEDIGO\KPI_UPDATE\thang8\Pharmacy_retail_Store_KPIs_August 2026"

def read_docx(path):
    try:
        z = zipfile.ZipFile(path)
        xml_content = z.read('word/document.xml')
        tree = ET.fromstring(xml_content)
        text = []
        for p in tree.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
            p_text = ''.join(p.itertext()).strip()
            if p_text:
                text.append(p_text)
        return '\n'.join(text)
    except Exception as e:
        return f"Error reading {path}: {e}"

def read_pptx(path):
    try:
        z = zipfile.ZipFile(path)
        slides_text = []
        slide_files = sorted([f for f in z.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')])
        for sf in slide_files:
            xml_content = z.read(sf)
            tree = ET.fromstring(xml_content)
            texts = [t.text for t in tree.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t') if t.text]
            if texts:
                slides_text.append(f'[{sf}]: ' + ' | '.join(texts))
        return '\n'.join(slides_text)
    except Exception as e:
        return f"Error reading {path}: {e}"

def read_excel_summary(path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        out = [f"Sheets: {wb.sheetnames}"]
        for s in wb.sheetnames[:3]:
            ws = wb[s]
            out.append(f"  Sheet '{s}': {ws.max_row} rows, {ws.max_column} cols")
            headers = [str(ws.cell(1, c).value) for c in range(1, min(ws.max_column+1, 10))]
            out.append(f"    Headers: {headers}")
            if ws.max_row >= 2:
                row2 = [str(ws.cell(2, c).value) for c in range(1, min(ws.max_column+1, 10))]
                out.append(f"    Sample row 2: {row2}")
        return '\n'.join(out)
    except Exception as e:
        return f"Error reading {path}: {e}"

print("="*80)
print("KHẢO SÁT TOÀN BỘ CÁC FILE KẾ HOẠCH & DỰ ÁN THÁNG 8.2026")
print("="*80)

for root, dirs, files in os.walk(base_dir):
    for f in files:
        fpath = os.path.join(root, f)
        rel_path = os.path.relpath(fpath, base_dir)
        ext = os.path.splitext(f)[1].lower()
        print(f"\n>>> FILE: {rel_path} ({ext})")
        if ext == '.docx':
            content = read_docx(fpath)
            print(content[:1500])
        elif ext == '.pptx':
            content = read_pptx(fpath)
            print(content[:1500])
        elif ext == '.xlsx':
            content = read_excel_summary(fpath)
            print(content)
        elif ext == '.url':
            try:
                with open(fpath, 'r', encoding='utf-8', errors='ignore') as ufile:
                    print(ufile.read())
            except Exception as e:
                print(e)
