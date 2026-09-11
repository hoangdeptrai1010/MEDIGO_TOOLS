import os
import sys
import io
import json
import calendar
import unicodedata
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.parse
import email
from email.message import EmailMessage

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _CURRENT_DIR)

from builder_engine import generate_kpisheet_package, parse_hcm_proposal, parse_hn_proposal

PORT = 8767
UPLOAD_DIR = os.path.join(_CURRENT_DIR, 'uploads')
OUTPUT_DIR = os.path.join(_CURRENT_DIR, 'output')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def parse_multipart_form(headers, rfile):
    """
    Zero-dependency multipart/form-data parser compatible with Python 3.9 -> 3.14+.
    """
    content_type = headers.get('Content-Type', '')
    if not content_type.startswith('multipart/form-data'):
        return {}, {}

    try:
        content_length = int(headers.get('Content-Length', 0))
    except (TypeError, ValueError):
        content_length = 0

    body = rfile.read(content_length)
    
    # Extract boundary
    boundary_marker = None
    for part in content_type.split(';'):
        part = part.strip()
        if part.startswith('boundary='):
            boundary_marker = part.split('=', 1)[1].strip('"').encode('utf-8')
            break
            
    if not boundary_marker:
        return {}, {}

    parts = body.split(b'--' + boundary_marker)
    fields = {}
    files = {}

    for p in parts:
        if not p or p == b'--\r\n' or p == b'--':
            continue
        if b'\r\n\r\n' not in p:
            continue
            
        header_part, content_part = p.split(b'\r\n\r\n', 1)
        content_part = content_part.rstrip(b'\r\n')
        
        header_text = header_part.decode('utf-8', errors='ignore')
        
        # Parse Content-Disposition
        cd_match = re.search(r'Content-Disposition:\s*form-data;\s*name="([^"]+)"(?:;\s*filename="([^"]+)")?', header_text, re.IGNORECASE)
        if cd_match:
            name = cd_match.group(1)
            filename = cd_match.group(2)
            
            if filename:
                files[name] = {
                    'filename': filename,
                    'content': content_part
                }
            else:
                fields[name] = content_part.decode('utf-8', errors='ignore')

    return fields, files

import re

class KPISheetHandler(SimpleHTTPRequestHandler):
    def translate_path(self, path):
        clean_path = urllib.parse.unquote(path.split('?')[0])
        if clean_path in ('/', '/index.html'):
            return os.path.join(_CURRENT_DIR, 'index.html')
        return os.path.join(_CURRENT_DIR, clean_path.lstrip('/'))

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def send_json(self, data, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        body = json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        
        if path == '/api/status':
            hcm_files = [f for f in os.listdir(_CURRENT_DIR) if f.endswith('.xlsx') and ('hcm' in f.lower() or 'cnt' in f.lower())]
            hn_files = [f for f in os.listdir(_CURRENT_DIR) if f.endswith('.xlsx') and ('hn' in f.lower() or 'quy' in f.lower() or 'e_xuat' in f.lower())]
            self.send_json({
                'status': 'ready',
                'port': PORT,
                'default_hcm_file': hcm_files[0] if hcm_files else None,
                'default_hn_file': hn_files[0] if hn_files else None
            })
            return

        if path.startswith('/api/download/'):
            fname = os.path.basename(urllib.parse.unquote(path[len('/api/download/'):]))
            cand_paths = [
                os.path.join(OUTPUT_DIR, fname),
                os.path.join(_CURRENT_DIR, '..', 'thang9', 'output', fname),
                os.path.join(_CURRENT_DIR, '..', 'thang8', 'output', fname),
                os.path.join(_CURRENT_DIR, '..', 'thang7', 'output', fname)
            ]
            fpath = next((p for p in cand_paths if os.path.exists(p)), None)
            if fpath and os.path.exists(fpath):
                safe_name = urllib.parse.quote(fname)
                ascii_name = unicodedata.normalize('NFKD', fname).encode('ascii', 'ignore').decode('ascii') or 'kpisheet.xlsx'
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{safe_name}')
                self.send_header('Content-Length', str(os.path.getsize(fpath)))
                self.end_headers()
                with open(fpath, 'rb') as f:
                    self.wfile.write(f.read())
                return
            else:
                self.send_json({'error': 'File not found'}, status=404)
                return

        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/api/build':
            content_type = self.headers.get('Content-Type', '')
            hcm_file_path = None
            hn_file_path = None
            month_num = 9

            if 'multipart/form-data' in content_type:
                fields, files = parse_multipart_form(self.headers, self.rfile)
                if 'month' in fields:
                    try:
                        month_num = int(fields['month'])
                    except Exception:
                        pass
                        
                if 'hcm_file' in files:
                    item = files['hcm_file']
                    clean_name = os.path.basename(item['filename'])
                    save_path = os.path.join(UPLOAD_DIR, f"hcm_proposal_T{month_num}_{clean_name}")
                    with open(save_path, 'wb') as f:
                        f.write(item['content'])
                    hcm_file_path = save_path
                    print(f"--> [Upload] Đã nhận file đề xuất HCM tải lên: {clean_name} ({len(item['content'])} bytes)")

                if 'hn_file' in files:
                    item = files['hn_file']
                    clean_name = os.path.basename(item['filename'])
                    save_path = os.path.join(UPLOAD_DIR, f"hn_proposal_T{month_num}_{clean_name}")
                    with open(save_path, 'wb') as f:
                        f.write(item['content'])
                    hn_file_path = save_path
                    print(f"--> [Upload] Đã nhận file đề xuất HN tải lên: {clean_name} ({len(item['content'])} bytes)")
            else:
                length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(length).decode('utf-8')
                data = json.loads(body) if body else {}
                month_num = int(data.get('month', 9))

            if not hcm_file_path:
                hcm_cand = os.path.join(_CURRENT_DIR, 'KPI CNT HCM Tháng 09.xlsx')
                if os.path.exists(hcm_cand): 
                    hcm_file_path = hcm_cand
                    print(f"--> [Default] Sử dụng file đề xuất HCM mặc định: {hcm_file_path}")
            if not hn_file_path:
                hn_cand = os.path.join(_CURRENT_DIR, 'e_xuat_KPI_Quy_3.26.xlsx')
                if os.path.exists(hn_cand): 
                    hn_file_path = hn_cand
                    print(f"--> [Default] Sử dụng file đề xuất HN mặc định: {hn_file_path}")

            if not hcm_file_path or not hn_file_path:
                self.send_json({'error': 'Thiếu file đề xuất KPI HCM hoặc HN'}, status=400)
                return

            print(f"--> [Executing Engine] Tháng {month_num}:")
            print(f"     1. HCM File: {hcm_file_path}")
            print(f"     2. HN File:  {hn_file_path}")

            try:
                project_folder = None
                cand_folders = [
                    os.path.join(_CURRENT_DIR, '..', f'thang{month_num}', f'Pharmacy_retail_Store_KPIs_September 2026'),
                    os.path.join(_CURRENT_DIR, '..', f'thang{month_num}'),
                    os.path.join(_CURRENT_DIR, '..', 'thang8', 'Pharmacy_retail_Store_KPIs_August 2026')
                ]
                for cf in cand_folders:
                    if os.path.exists(cf):
                        project_folder = cf
                        break

                res = generate_kpisheet_package(
                    hcm_file_path=hcm_file_path,
                    hn_file_path=hn_file_path,
                    month_num=month_num,
                    project_folder=project_folder
                )
                res['download_url'] = f'/api/download/{res["filename"]}'
                print(f"--> [SUCCESS] Đã tạo thành công gói kế hoạch: {res['filename']} (Kỳ: {res['period']})")
                self.send_json({'success': True, 'data': res})
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.send_json({'error': str(e)}, status=500)
            return

        self.send_json({'error': 'Not found'}, status=404)

def run_server():
    server = HTTPServer(('127.0.0.1', PORT), KPISheetHandler)
    print(f"\n=======================================================")
    print(f"🌟 TOOL QUẢN LÝ KPI SHEET ĐANG CHẠY TẠI:")
    print(f"👉 http://localhost:{PORT}")
    print(f"=======================================================\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nĐã dừng server.")

if __name__ == '__main__':
    run_server()
