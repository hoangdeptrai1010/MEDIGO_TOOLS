import os, zipfile, xml.etree.ElementTree as ET

def read_docx(path):
    try:
        z = zipfile.ZipFile(path)
        xml_content = z.read('word/document.xml')
        tree = ET.fromstring(xml_content)
        texts = [t.text for t in tree.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text]
        return ' '.join(texts)
    except Exception as e:
        return f"Error reading {path}: {e}"

def read_pptx(path):
    try:
        z = zipfile.ZipFile(path)
        texts = []
        slide_files = sorted([f for f in z.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')])
        for sf in slide_files:
            xml_content = z.read(sf)
            tree = ET.fromstring(xml_content)
            st = [t.text for t in tree.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t') if t.text]
            if st:
                texts.append(f"[{sf}]: " + ' | '.join(st))
        return '\n'.join(texts)
    except Exception as e:
        return f"Error reading {path}: {e}"

base = 'thang8/Pharmacy_retail_Store_KPIs_August 2026'

print("=== WHATSAPP HN ===")
print(read_docx(os.path.join(base, 'Hà Nội/KPIs/KPIs Whatsapp tháng 8.2026.docx')))

print("\n=== HN KAT ===")
for f in os.listdir(os.path.join(base, 'Hà Nội/Dự án Mini/Chương trình')):
    if f.endswith('.pptx'):
        print(f"\n--- {f} ---")
        print(read_pptx(os.path.join(base, 'Hà Nội/Dự án Mini/Chương trình', f)))

print("\n=== HCM KAT ===")
for f in os.listdir(os.path.join(base, 'Hồ Chí Minh/Dự án mini/Chương trình')):
    if f.endswith('.pptx'):
        print(f"\n--- {f} ---")
        print(read_pptx(os.path.join(base, 'Hồ Chí Minh/Dự án mini/Chương trình', f)))
