import zipfile, os, shutil, re, subprocess

src_file = 'baocaokpi_thang8_hoanthien.xlsx'
backup_file = 'baocaokpi_thang8_hoanthien.xlsx.bak'
shutil.copy2(src_file, backup_file)

z_in = zipfile.ZipFile(src_file, 'r')
new_zip_name = 'temp_mod.xlsx'
z_out = zipfile.ZipFile(new_zip_name, 'w', zipfile.ZIP_DEFLATED)

for item in z_in.infolist():
    content = z_in.read(item.filename)
    if item.filename == 'xl/worksheets/sheet4.xml':
        s4 = content.decode('utf-8')
        
        # 1. Update Col D formula: SUM(E3:G3)/DAY($A$1) -> SUM(F3:G3)/DAY($A$1)
        s4 = s4.replace('SUM(E3:G3)/DAY($A$1)', 'SUM(F3:G3)/DAY($A$1)')
        
        # 2. Update Col J formula for Hanoi:
        old_hanoi_ifs = (
            'IFS(\n'
            '      AND($H3&gt;=3400000,$I3&gt;=800000),4000000,\n'
            '      AND($H3&gt;=3000000,$I3&gt;=700000),2800000,\n'
            '      AND($H3&gt;=2600000,$I3&gt;=550000),2000000,\n'
            '      AND($H3&gt;=2200000,$I3&gt;=450000),1200000,\n'
            '      TRUE,0\n'
            '    )'
        )
        
        new_hanoi_ifs = (
            'IFS(\n'
            '      AND($H3&gt;=3400000,$I3&gt;=900000),4000000,\n'
            '      AND($H3&gt;=3200000,$I3&gt;=750000),2800000,\n'
            '      AND($H3&gt;=2800000,$I3&gt;=650000),2200000,\n'
            '      AND($H3&gt;=2500000,$I3&gt;=520000),1600000,\n'
            '      AND($H3&gt;=2200000,$I3&gt;=450000),1000000,\n'
            '      TRUE,0\n'
            '    )'
        )
        
        if old_hanoi_ifs in s4:
            s4 = s4.replace(old_hanoi_ifs, new_hanoi_ifs)
            print("✅ Đã thay thế thành công công thức bậc Hà Nội sang khung bậc Slide Tháng 8!")
        else:
            # Try regex replacement if whitespace differs
            print("⚠️ Cần kiểm tra regex cho old_hanoi_ifs...")
            pattern = r'IFS\s*\(\s*AND\(\$H3&gt;=3400000,\$I3&gt;=800000\),4000000.*?TRUE,0\s*\)'
            s4 = re.sub(pattern, new_hanoi_ifs, s4, count=1, flags=re.DOTALL)
            print("✅ Đã áp dụng regex thay thế công thức bậc Hà Nội!")
            
        content = s4.encode('utf-8')
    z_out.writestr(item, content)

z_in.close()
z_out.close()

# Overwrite original
shutil.move(new_zip_name, src_file)
shutil.copy2(src_file, 'thang8/baocaokpi_thang8_hoanthien.xlsx')
print("✅ Đã ghi đè thành công cả 2 file báo cáo!")

# Run recalc via Excel COM
print("--> Chạy Excel COM tính toán và lưu sẵn các giá trị mới...")
cmd = ['powershell', '-ExecutionPolicy', 'Bypass', '-File', 'recalc_workbook.ps1', '-FilePath', os.path.abspath(src_file)]
res = subprocess.run(cmd, capture_output=True, text=True)
print(res.stdout)

# Also recalc thang8 copy
cmd2 = ['powershell', '-ExecutionPolicy', 'Bypass', '-File', 'recalc_workbook.ps1', '-FilePath', os.path.abspath('thang8/baocaokpi_thang8_hoanthien.xlsx')]
res2 = subprocess.run(cmd2, capture_output=True, text=True)
print("Recalc thang8 copy done.")
