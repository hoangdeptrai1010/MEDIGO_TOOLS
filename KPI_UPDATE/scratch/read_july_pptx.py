import sys, zipfile, xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

for p in ['thang7/Pharmacy_Retail_store_KPIs_July_2026/Hà Nội/Dự án/Chương trình/Hanoi_Projects_July_2026_v1.pptx',
          'thang7/Pharmacy_Retail_store_KPIs_July_2026/Hồ Chí Minh/Dự án/Chương trình/HCM_Projects_July_2026_v1.pptx']:
    print(f"\n==================== {p} ====================")
    z = zipfile.ZipFile(p)
    for f in z.namelist():
        if f.startswith('ppt/slides/slide'):
            xml = z.read(f)
            tree = ET.fromstring(xml)
            texts = [t.text for t in tree.iter() if t.tag.endswith('}t') and t.text]
            print(f"[{f}]:\n" + "\n".join(texts))
