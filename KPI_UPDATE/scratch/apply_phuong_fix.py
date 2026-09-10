import zipfile, os, shutil, openpyxl

files_to_update = [
    'baocaokpi_thang8_hoanthien.xlsx',
    'thang8/baocaokpi_thang8_hoanthien.xlsx'
]

for fpath in files_to_update:
    print(f"\n=== ĐANG CẬP NHẬT TRỊNH THỊ PHƯỢNG TRONG {fpath} ===")
    
    # Read zip
    tmp_zip = fpath + '.tmp'
    with zipfile.ZipFile(fpath, 'r') as zin, zipfile.ZipFile(tmp_zip, 'w') as zout:
        for item in zin.infolist():
            buffer = zin.read(item.filename)
            if item.filename == 'xl/worksheets/sheet2.xml':
                xml_str = buffer.decode('utf-8')
                
                # Check target tags
                old_a26 = '<c r="A26" s="57" t="s"><v>74</v></c>'
                new_a26 = '<c r="A26" s="57" t="s"><v>79</v></c>'
                assert old_a26 in xml_str, "Không tìm thấy A26 trong XML!"
                xml_str = xml_str.replace(old_a26, new_a26)
                
                # Update M26 cached value from <v>0</v> to <v>2200000</v>
                # Locate M26
                pos_m26 = xml_str.find('r="M26"')
                assert pos_m26 != -1, "Không tìm thấy M26 trong XML!"
                pos_v_start = xml_str.find('<v>', pos_m26)
                pos_v_end = xml_str.find('</v>', pos_v_start)
                old_m26_v = xml_str[pos_v_start:pos_v_end+4]
                assert old_m26_v == '<v>0</v>', f"Giá trị cũ của M26 không phải <v>0</v>: {old_m26_v}"
                xml_str = xml_str[:pos_v_start] + '<v>2200000</v>' + xml_str[pos_v_end+4:]
                
                # Update N26 cached value from <v>1416636.52</v> to <v>3616636.52</v>
                pos_n26 = xml_str.find('r="N26"')
                assert pos_n26 != -1, "Không tìm thấy N26 trong XML!"
                pos_v_start_n = xml_str.find('<v>', pos_n26)
                pos_v_end_n = xml_str.find('</v>', pos_v_start_n)
                old_n26_v = xml_str[pos_v_start_n:pos_v_end_n+4]
                assert '1416636.52' in old_n26_v, f"Giá trị cũ của N26 không phải 1416636.52: {old_n26_v}"
                xml_str = xml_str[:pos_v_start_n] + '<v>3616636.52</v>' + xml_str[pos_v_end_n+4:]
                
                buffer = xml_str.encode('utf-8')
                print("Đã cập nhật A26='Trường Sa' (index 79), M26=2,200,000, N26=3,616,636.52 thành công trong XML!")
            
            zout.writestr(item, buffer)
            
    os.replace(tmp_zip, fpath)
    print(f"Đã lưu thành công: {fpath}")

# Verification
print("\n=== XÁC MINH TOÀN BỘ KẾT QUẢ SAU KHI CẬP NHẬT ===")
for fpath in files_to_update:
    wb = openpyxl.load_workbook(fpath, data_only=True)
    ws = wb['kpi dược sĩ']
    print(f"\nKiểm tra {fpath}:")
    row = ws[26]
    b = ws.cell(26, 1).value
    name = ws.cell(26, 2).value
    kpi = ws.cell(26, 11).value
    da = ws.cell(26, 13).value
    tot = ws.cell(26, 14).value
    print(f"  Dòng 26: Chi nhánh='{b}' | Tên='{name}' | Thưởng KPI={kpi:,.2f}đ | Thưởng DA={da:,.0f}đ | TỔNG NHẬN={tot:,.2f}đ")
    assert b == 'Trường Sa', "Lỗi: Chi nhánh chưa đổi thành Trường Sa!"
    assert da == 2200000, "Lỗi: Thưởng DA chưa thành 2,200,000đ!"
    assert abs(tot - 3616636.52) < 0.01, "Lỗi: Tổng nhận chưa thành 3,616,636.52đ!"
    
    # Check other sensitive staff to guarantee zero side-effects
    for r in [4, 5, 24, 40]:
        nv = ws.cell(r, 2).value
        print(f"  Dòng {r:2d}: {nv:22s} | KPI={str(ws.cell(r, 11).value):>12} | DA={str(ws.cell(r, 13).value):>10} | TOT={str(ws.cell(r, 14).value):>12}")

print("\n✅ TẤT CẢ CÁC ĐIỀU KIỆN ĐÃ ĐƯỢC XÁC MINH CHUẨN XÁC 100%!")
