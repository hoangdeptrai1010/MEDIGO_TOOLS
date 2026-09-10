import zipfile, re

z = zipfile.ZipFile('baocaokpi_thang8_hoanthien.xlsx.bak')
s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')

for r in [3, 4, 5, 24, 26]:
    for col in ['K', 'M', 'N']:
        cell = f"{col}{r}"
        m = re.search(rf'<c r="{cell}"[^>]*>(.*?)</c>', s2)
        if m:
            print(f"{cell}: {m.group(1)}")
