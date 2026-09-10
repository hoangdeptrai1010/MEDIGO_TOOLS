import zipfile

z = zipfile.ZipFile('goc/NHÀ THUỐC THÁNG 8 2026.xlsx')
for name in z.namelist():
    if 'sheet2' in name:
        data = z.read(name).decode('utf-8')
        pos = data.find('r="F3"')
        if pos != -1:
            print("Found F3 in goc", name)
            print(data[pos-10:pos+800])
