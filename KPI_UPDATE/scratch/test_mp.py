import sys, io
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

sys.path.insert(0, 'TOOL_KPISHEET')
from app import parse_multipart_form

boundary = "----WebKitFormBoundaryABC123"
body_parts = [
    f"--{boundary}",
    'Content-Disposition: form-data; name="month"',
    '',
    '9',
    f"--{boundary}",
    'Content-Disposition: form-data; name="hcm_file"; filename="KPI CNT HCM Tháng 09.xlsx"',
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '',
    'FAKE_EXCEL_DATA',
    f"--{boundary}--",
    ''
]
body_bytes = "\r\n".join(body_parts).encode('utf-8')

headers = {
    'Content-Type': f'multipart/form-data; boundary={boundary}',
    'Content-Length': str(len(body_bytes))
}

fields, files = parse_multipart_form(headers, io.BytesIO(body_bytes))
print("Parsed fields:", fields)
print("Parsed files:", list(files.keys()))
