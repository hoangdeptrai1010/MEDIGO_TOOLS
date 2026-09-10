import zipfile, re

z = zipfile.ZipFile('baocaokpi_thang8_hoanthien.xlsx')
s4 = z.read('xl/worksheets/sheet4.xml').decode('utf-8')
m = re.findall(r'<c r="J\d+"[^>]*>.*?</c>', s4, re.DOTALL)
for item in m[:5]:
    print(item[:200])
