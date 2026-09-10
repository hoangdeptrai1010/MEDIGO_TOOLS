import zipfile, os, shutil, re

for tfile in ['NHÀ THUỐC THÁNG 8 2026.xlsx', 'goc/NHÀ THUỐC THÁNG 8 2026.xlsx']:
    if not os.path.exists(tfile): continue
    print(f"Updating template: {tfile}")
    z_in = zipfile.ZipFile(tfile, 'r')
    temp_zip = tfile + '.tmp'
    z_out = zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED)
    
    for item in z_in.infolist():
        content = z_in.read(item.filename)
        if item.filename == 'xl/worksheets/sheet4.xml':
            s4 = content.decode('utf-8')
            s4 = s4.replace('SUM(E3:G3)/DAY($A$1)', 'SUM(F3:G3)/DAY($A$1)')
            content = s4.encode('utf-8')
        z_out.writestr(item, content)
        
    z_in.close()
    z_out.close()
    shutil.move(temp_zip, tfile)
    print(f"Done updating {tfile}")
