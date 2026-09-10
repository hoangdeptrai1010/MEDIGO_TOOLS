import zipfile, os, shutil, re

tfile = 'goc/NHÀ THUỐC THÁNG 8 2026.xlsx'
bakfile = 'goc/NHÀ THUỐC THÁNG 8 2026.xlsx.bak'
shutil.copy2(tfile, bakfile)

z_in = zipfile.ZipFile(tfile, 'r')
temp_zip = tfile + '.tmp'
z_out = zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED)

for item in z_in.infolist():
    content = z_in.read(item.filename)
    if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
        xml = content.decode('utf-8')
        # Replace IFS( with _xlfn.IFS( where not already preceded by _xlfn.
        xml = re.sub(r'(?<!_xlfn\.)IFS\(', '_xlfn.IFS(', xml)
        xml = re.sub(r'(?<!_xlfn\.)XLOOKUP\(', '_xlfn.XLOOKUP(', xml)
        content = xml.encode('utf-8')
    z_out.writestr(item, content)

z_in.close()
z_out.close()
shutil.move(temp_zip, tfile)
print("✅ Đã chuẩn hóa toàn bộ hàm OpenXML (_xlfn.) trong template gốc!")
