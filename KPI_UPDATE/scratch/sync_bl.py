import shutil, os

src = 'thang8/bangluong_thang8_hoanthien.xlsx'
dst = 'thang8/BANGLUONGTHANG8.xlsx'

try:
    shutil.copy2(src, dst)
    print("Synced successfully to BANGLUONGTHANG8.xlsx")
except Exception as e:
    print(f"Sync warning: {e}")
