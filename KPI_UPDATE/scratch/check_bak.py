import zipfile, re

z = zipfile.ZipFile('baocaokpi_thang8_hoanthien.xlsx.bak')
s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
m = re.search(r'<c r="K3"[^>]*><f[^>]*>(.*?)</f>', s2, re.DOTALL)
if m:
    print("Formula in K3 sheet 2 (bak):\n" + m.group(1))

s4 = z.read('xl/worksheets/sheet4.xml').decode('utf-8')
m4 = re.search(r'<c r="J3"[^>]*><f[^>]*>(.*?)</f>', s4, re.DOTALL)
if m4:
    print("Formula in J3 sheet 4 (bak):\n" + m4.group(1))
