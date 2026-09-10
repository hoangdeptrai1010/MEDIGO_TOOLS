@echo off
chcp 65001 > nul
echo ==============================================================================
echo             MEDIGO PRICING ENGINE (TOOL V5) - ĐANG CHẠY TỰ ĐỘNG
echo ==============================================================================
echo.
echo [*] Dang quet du lieu va xuat bao cao dinh gia...
uv run --with pandas --with python-calamine --with openpyxl python "%~dp0TOOL\toolv5.py"
echo.
echo ==============================================================================
echo [*] Hoan tat! File bao cao da duoc luu tai thu muc: OUTPUT\
echo ==============================================================================
pause
