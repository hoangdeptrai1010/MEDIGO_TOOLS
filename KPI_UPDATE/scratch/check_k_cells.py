import zipfile, re

z = zipfile.ZipFile('baocaokpi_thang8_hoanthien.xlsx.bak')
s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')

m = re.findall(r'<c r="K\d+"[^>]*>.*?</c>', s2, re.DOTALL)
for item in m[:5]:
    print(item[:150])
