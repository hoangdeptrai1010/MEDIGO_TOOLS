import zipfile

z = zipfile.ZipFile('thang8/baocaokpi_thang8_hoanthien.xlsx')
data = z.read('xl/worksheets/sheet2.xml').decode('utf-8')

pos26 = data.find('<row r="26"')
pos27 = data.find('<row r="27"')
if pos26 != -1 and pos27 != -1:
    row26_xml = data[pos26:pos27]
    print('Row 26 XML in sheet2:')
    print(row26_xml)
