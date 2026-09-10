import urllib.request
import json
import subprocess
import time
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

app_path = 'd:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET/app.py'
proc = subprocess.Popen(['uv', 'run', '--with', 'openpyxl', 'python', app_path])
time.sleep(2)

try:
    # 1. Test GET /api/status
    req = urllib.request.urlopen('http://127.0.0.1:8767/api/status')
    data = json.loads(req.read().decode('utf-8'))
    print('✅ API Status:', data)
    assert data['status'] == 'ready'

    # 2. Test POST /api/build (JSON mode)
    build_req = urllib.request.Request(
        'http://127.0.0.1:8767/api/build',
        data=json.dumps({'month': 9}).encode('utf-8'),
        headers={'Content-Type': 'application/json'}
    )
    build_res = urllib.request.urlopen(build_req)
    build_data = json.loads(build_res.read().decode('utf-8'))
    print('✅ API Build:', build_data['success'])
    print('   Filename:', build_data['data']['filename'])
    print('   Staff Count:', build_data['data']['staff_count'])
    print('   Stores Count:', build_data['data']['stores_count'])

    # 3. Test Download Endpoint
    dl_url = 'http://127.0.0.1:8767' + build_data['data']['download_url']
    dl_res = urllib.request.urlopen(dl_url)
    dl_bytes = dl_res.read()
    print('✅ Download Endpoint Success! File size:', len(dl_bytes), 'bytes')
    assert len(dl_bytes) > 10000

    print('\n🎉 TẤT CẢ CÁC API ENDPOINTS CỦA TOOL_KPISHEET ĐÃ TEST THÀNH CÔNG 100%!')
finally:
    proc.terminate()
