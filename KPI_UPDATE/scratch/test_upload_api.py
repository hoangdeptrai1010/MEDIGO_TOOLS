import urllib.request
import urllib.parse
import threading
import time
import os
import sys
import json

if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

sys.path.insert(0, 'TOOL_KPISHEET')
from app import KPISheetHandler, HTTPServer

server = HTTPServer(('127.0.0.1', 8771), KPISheetHandler)
t = threading.Thread(target=server.serve_forever, daemon=True)
t.start()
time.sleep(1)

boundary = '----WebKitFormBoundaryTest12345678'
hcm_bytes = open('TOOL_KPISHEET/KPI CNT HCM Tháng 09.xlsx', 'rb').read()
hn_bytes = open('TOOL_KPISHEET/e_xuat_KPI_Quy_3.26.xlsx', 'rb').read()

body = bytearray()
body.extend(f'--{boundary}\r\n'.encode('utf-8'))
body.extend(b'Content-Disposition: form-data; name="month"\r\n\r\n9\r\n')

body.extend(f'--{boundary}\r\n'.encode('utf-8'))
body.extend(b'Content-Disposition: form-data; name="hcm_file"; filename="MY_CUSTOM_HCM_T9.xlsx"\r\n')
body.extend(b'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n')
body.extend(hcm_bytes)
body.extend(b'\r\n')

body.extend(f'--{boundary}\r\n'.encode('utf-8'))
body.extend(b'Content-Disposition: form-data; name="hn_file"; filename="MY_CUSTOM_HN_T9.xlsx"\r\n')
body.extend(b'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n')
body.extend(hn_bytes)
body.extend(b'\r\n')

body.extend(f'--{boundary}--\r\n'.encode('utf-8'))

req = urllib.request.Request(
    'http://127.0.0.1:8771/api/build',
    data=bytes(body),
    headers={'Content-Type': f'multipart/form-data; boundary={boundary}'},
    method='POST'
)

try:
    with urllib.request.urlopen(req) as resp:
        res = json.loads(resp.read().decode('utf-8'))
        print('Upload API Response Success:', res.get('success'))
        print('Filename generated:', res['data']['filename'])
        print('Staff Count in result:', res['data']['staff_count'])
        print('Stores Count in result:', res['data']['stores_count'])
        print('Staff sample (first 3):')
        for s in res['data']['staff_sample'][:3]:
            print('  ', s)
except Exception as e:
    print('Upload API Error:', e)
finally:
    server.shutdown()
