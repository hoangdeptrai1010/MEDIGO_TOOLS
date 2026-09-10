import zipfile, re

z = zipfile.ZipFile('baocaokpi_thang8_hoanthien.xlsx')
s4 = z.read('xl/worksheets/sheet4.xml').decode('utf-8')

for c in ['D5', 'J5', 'K5', 'M5']:
    pattern = rf'<c r="{c}"[^>]*>(.*?)</c>'
    m = re.search(pattern, s4)
    if m:
        print(f'{c}: {m.group(1)}')
