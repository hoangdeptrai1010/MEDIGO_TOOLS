import zipfile, xml.etree.ElementTree as ET

z = zipfile.ZipFile('thang8/baocaokpi_thang8_hoanthien.xlsx')
sst_xml = z.read('xl/sharedStrings.xml')
root = ET.fromstring(sst_xml)

ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
for idx, si in enumerate(root.findall('ns:si', ns)):
    t = ''.join(si.itertext())
    if t in ['Trường Sa', 'Minh Châu']:
        print(f'Index {idx}: "{t}"')
