import zipfile, os, re, openpyxl

files_to_update = [
    'baocaokpi_thang8_hoanthien.xlsx',
    'thang8/baocaokpi_thang8_hoanthien.xlsx'
]

for fpath in files_to_update:
    print(f"\n=== CHUẨN HÓA CỘT M CHO TOÀN BỘ 53 DÒNG TRONG {fpath} ===")
    
    tmp_zip = fpath + '.tmp'
    with zipfile.ZipFile(fpath, 'r') as zin, zipfile.ZipFile(tmp_zip, 'w') as zout:
        for item in zin.infolist():
            buffer = zin.read(item.filename)
            if item.filename == 'xl/worksheets/sheet2.xml':
                xml_str = buffer.decode('utf-8')
                
                # Regex matching &amp;
                pattern = re.compile(r'(<c r="M(\d+)"[^>]*><f[^>]*>)IFERROR\(\s*_xlfn\.XLOOKUP\(\s*\$A\2&amp;\$B\2,\s*\'Dự án T8\'!\$A:\$A&amp;\'Dự án T8\'!\$B:\$B,\s*\'Dự án T8\'!\$M:\$M,\s*0\s*\)\*\$L\2,\s*0\)(</f>)', re.DOTALL)
                
                def repl(match):
                    row_num = match.group(2)
                    open_tag = match.group(1)
                    close_tag = match.group(3)
                    new_f = f"IFERROR(_xlfn.XLOOKUP($B{row_num}, 'Dự án T8'!$B:$B, 'Dự án T8'!$M:$M, 0)*$L{row_num}, 0)"
                    return f"{open_tag}{new_f}{close_tag}"
                
                new_xml, count = pattern.subn(repl, xml_str)
                print(f"-> Đã chuyển đổi thành công {count} ô trong cột M sang tra cứu theo tên nhân viên $B!")
                buffer = new_xml.encode('utf-8')
            
            zout.writestr(item, buffer)
            
    os.replace(tmp_zip, fpath)
    print(f"Đã lưu thành công: {fpath}")

# Verification
print("\n=== KIỂM TRA LẠI SAU KHI CHUẨN HÓA ===")
for fpath in files_to_update:
    wb = openpyxl.load_workbook(fpath, data_only=False)
    ws = wb['kpi dược sĩ']
    m_with_a = 0
    for r in range(3, ws.max_row+1):
        f = ws.cell(r, 13).value
        f_txt = f.text if hasattr(f, 'text') else str(f)
        if '$A' in f_txt:
            m_with_a += 1
    print(f"{fpath}: Số dòng cột M còn kẹp chi nhánh ($A) = {m_with_a}")
    assert m_with_a == 0, "Vẫn còn dòng dùng $A!"
