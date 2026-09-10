import shutil, os

src = 'thang8/bangluong_thang8_hoanthien.xlsx'
targets = [
    'thang8/BANGLUONGTHANG8.xlsx',
    'thang8/BANGLUONGTHANG8_HOANG_.xlsx',
    'thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx'
]

for t in targets:
    shutil.copy2(src, t)
    print(f"Synced {src} -> {t}")
