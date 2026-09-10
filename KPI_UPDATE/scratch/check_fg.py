import zipfile, re

z = zipfile.ZipFile('baocaokpi_thang8_hoanthien.xlsx')
s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')

for col in ['F', 'G']:
    m = re.search(rf'<c r="{col}3"[^>]*><f[^>]*>(.*?)</f>', s2, re.DOTALL)
    if m:
        print(f"Col {col}3:\n{m.group(1)}")
