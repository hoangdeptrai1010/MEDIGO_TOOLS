import zipfile, os, shutil, subprocess

src_file = 'baocaokpi_thang8_hoanthien.xlsx'

z_in = zipfile.ZipFile(src_file, 'r')
new_zip_name = 'temp_mod2.xlsx'
z_out = zipfile.ZipFile(new_zip_name, 'w', zipfile.ZIP_DEFLATED)

nested_if_j3 = (
    'IFERROR('
    'IF(OR($A3=&quot;Hàng Bông&quot;,$A3=&quot;Đường Láng&quot;),'
    'IF(AND($H3&gt;=3400000,$I3&gt;=900000),4000000,'
    'IF(AND($H3&gt;=3200000,$I3&gt;=750000),2800000,'
    'IF(AND($H3&gt;=2800000,$I3&gt;=650000),2200000,'
    'IF(AND($H3&gt;=2500000,$I3&gt;=520000),1600000,'
    'IF(AND($H3&gt;=2200000,$I3&gt;=450000),1000000,0))))),'
    'IF(AND($H3&gt;=1500000,$I3&gt;=700000),4000000,'
    'IF(AND($H3&gt;=1200000,$I3&gt;=650000),2800000,'
    'IF(AND($H3&gt;=1000000,$I3&gt;=580000),2200000,'
    'IF(AND($H3&gt;=850000,$I3&gt;=480000),1600000,'
    'IF(AND($H3&gt;=630000,$I3&gt;=350000),1000000,0)))))'
    '),0)'
)

import re

for item in z_in.infolist():
    content = z_in.read(item.filename)
    if item.filename == 'xl/worksheets/sheet4.xml':
        s4 = content.decode('utf-8')
        
        # Replace J3 formula with nested_if_j3
        s4 = re.sub(r'<f t="shared" ref="J3:J53" si="4">.*?</f>', f'<f t="shared" ref="J3:J53" si="4">{nested_if_j3}</f>', s4, flags=re.DOTALL)
        
        content = s4.encode('utf-8')
    z_out.writestr(item, content)

z_in.close()
z_out.close()

shutil.move(new_zip_name, src_file)
shutil.copy2(src_file, 'thang8/baocaokpi_thang8_hoanthien.xlsx')

print("--> Chạy recalc Excel COM...")
cmd = ['powershell', '-ExecutionPolicy', 'Bypass', '-File', 'recalc_workbook.ps1', '-FilePath', os.path.abspath(src_file)]
res = subprocess.run(cmd, capture_output=True, text=True)
print(res.stdout)

shutil.copy2(src_file, 'thang8/baocaokpi_thang8_hoanthien.xlsx')
print("Done!")
