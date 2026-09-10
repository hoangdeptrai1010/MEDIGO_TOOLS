import sys, zipfile, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

for p in ['thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/Dự án/Chương trình/Hanoi_Projects_August_2026_v1.pptx',
          'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hồ Chí Minh/Dự án/Chương trình/HCM_Projects_August_2026_v1.pptx']:
    print(f"\n==================== {p} ====================")
    z = zipfile.ZipFile(p)
    for f in z.namelist():
        if f.startswith('ppt/slides/slide'):
            xml = z.read(f)
            tree = ET.fromstring(xml)
            texts = [t.text for t in tree.iter() if t.tag.endswith('}t') and t.text]
            print(f"[{f}]:\n" + "\n".join(texts))
