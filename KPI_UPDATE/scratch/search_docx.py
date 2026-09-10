import zipfile, xml.etree.ElementTree as ET

def search_word(path, word):
    z = zipfile.ZipFile(path)
    xml_content = z.read('word/document.xml')
    tree = ET.fromstring(xml_content)
    full_text = ''.join(tree.itertext())
    return word.lower() in full_text.lower()

for p in [
    'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/KPIs/Chương trình/Chương trình KPI.docx',
    'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hồ Chí Minh/KPIs/Chương trình/Chương trình KPI.docx'
]:
    print(p)
    print('  Chua tu "trung binh bill":', search_word(p, 'trung bình bill'))
    print('  Chua tu "bill":', search_word(p, 'bill'))
    print('  Chua tu "hàng điểm":', search_word(p, 'hàng điểm'))
    print('  Chua tu "25%":', search_word(p, '25%'))
