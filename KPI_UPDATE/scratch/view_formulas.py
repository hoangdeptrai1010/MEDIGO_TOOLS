import zipfile, re

z = zipfile.ZipFile('baocaokpi_thang8_hoanthien.xlsx')
s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
s3 = z.read('xl/worksheets/sheet3.xml').decode('utf-8')

m_k3 = re.search(r'<c r="K3"[^>]*><f[^>]*>(.*?)</f>', s2, re.DOTALL)
if m_k3:
    print('K3 formula:\n', m_k3.group(1))

m_g10 = re.search(r'<c r="G10"[^>]*><f[^>]*>(.*?)</f>', s3, re.DOTALL)
if m_g10:
    print('\nG10 formula:\n', m_g10.group(1))
