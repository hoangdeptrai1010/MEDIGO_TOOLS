import os, zipfile, xml.etree.ElementTree as ET

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

root_dir = 'thang8/Pharmacy_retail_Store_KPIs_August 2026'
for r, d, files in os.walk(root_dir):
    for f in files:
        if f.endswith('.pptx'):
            full_path = os.path.join(r, f)
            print("="*60)
            print(f"FILE: {f}")
            print("="*60)
            print(read_pptx(full_path))
