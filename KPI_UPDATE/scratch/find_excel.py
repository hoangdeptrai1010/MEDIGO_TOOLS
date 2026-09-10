import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

for root, dirs, files in os.walk('.'):
    for f in files:
        if f.endswith('.xlsx') or f.endswith('.xlsm'):
            p = os.path.join(root, f)
            print(p)
