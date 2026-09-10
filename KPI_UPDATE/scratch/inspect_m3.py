import zipfile

z = zipfile.ZipFile('thang8/baocaokpi_thang8_hoanthien.xlsx')
data = z.read('xl/worksheets/sheet2.xml').decode('utf-8')

pos = data.find('r="M3"')
print(data[pos-5:pos+250])
