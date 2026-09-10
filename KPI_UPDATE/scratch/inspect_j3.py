import zipfile, re

z = zipfile.ZipFile('baocaokpi_thang8_hoanthien.xlsx')
s4 = z.read('xl/worksheets/sheet4.xml').decode('utf-8')
m = re.search(r'<c r="J3"[^>]*><f[^>]*>(.*?)</f>', s4, re.DOTALL)
if m:
    print("Formula in J3:\n" + m.group(1))
