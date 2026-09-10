import zipfile, os, shutil, re

def convert_to_nested_if(file_path):
    print(f"Converting {file_path} to native nested IF...")
    z_in = zipfile.ZipFile(file_path, 'r')
    temp_zip = file_path + '.tmp'
    z_out = zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED)
    
    for item in z_in.infolist():
        content = z_in.read(item.filename)
        if item.filename == 'xl/worksheets/sheet2.xml':
            xml = content.decode('utf-8')
            
            def repl_k(m):
                r = m.group(1)
                nested = (
                    f'IF(AND(I{r}&gt;=T{r},E{r}&gt;=D{r}),AD{r}*1.5%,'
                    f'IF(AND(I{r}&gt;=T{r},E{r}&lt;D{r}),AD{r}*1.3%,'
                    f'IF(AND(I{r}&gt;=S{r},E{r}&gt;=D{r}),AD{r}*1.3%,'
                    f'IF(AND(I{r}&gt;=S{r},E{r}&lt;D{r}),AD{r}*1.1%,'
                    f'IF(AND(I{r}&gt;=R{r},E{r}&gt;=D{r}),AD{r}*1.2%,'
                    f'IF(AND(I{r}&gt;=R{r},E{r}&lt;D{r}),AD{r}*1%,""))))))'
                )
                return f'G{r}&gt;=F{r}),{nested},""'
            
            # Match flexibly
            pattern_k = r'G(\d+)&gt;=F\1\s*\),\s*(?:_xludf\.)?(?:_xlfn\.)?IFS\(.*?\bTRUE\s*,\s*""\s*\)\s*,\s*""'
            xml, count_k = re.subn(pattern_k, repl_k, xml, flags=re.DOTALL)
            print(f"  sheet2.xml: replaced {count_k} occurrences")
            content = xml.encode('utf-8')

        elif item.filename == 'xl/worksheets/sheet3.xml':
            xml = content.decode('utf-8')
            def repl_g(m):
                r = m.group(1)
                nested = (
                    f'IF(E{r}&gt;=T{r},D{r}*0.5%+1000000,'
                    f'IF(E{r}&gt;=P{r},D{r}*0.4%+700000,'
                    f'IF(E{r}&gt;=K{r},D{r}*0.35%+500000,0)))'
                )
                return f'W{r}&lt;&gt;"Đạt",0,{nested}'
                
            pattern_g = r'W(\d+)&lt;&gt;"Đạt",\s*0,\s*(?:_xludf\.)?(?:_xlfn\.)?IFS\(.*?\bTRUE\s*,\s*0\s*\)'
            xml, count_g = re.subn(pattern_g, repl_g, xml, flags=re.DOTALL)
            print(f"  sheet3.xml: replaced {count_g} occurrences")
            content = xml.encode('utf-8')

        z_out.writestr(item, content)
        
    z_in.close()
    z_out.close()
    shutil.move(temp_zip, file_path)
    print(f"✅ Converted {file_path}")

for p in ['goc/NHÀ THUỐC THÁNG 8 2026.xlsx', 'baocaokpi_thang8_hoanthien.xlsx']:
    if os.path.exists(p):
        convert_to_nested_if(p)
