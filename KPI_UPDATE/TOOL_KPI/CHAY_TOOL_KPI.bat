@echo off
chcp 65001 > nul
title MEDIGO - TOOL TẠO BÁO CÁO KPI & DỰ ÁN

echo ========================================================
echo        MEDIGO TOOL TẠO BÁO CÁO KPI & DỰ ÁN
echo ========================================================
echo.
echo Đang khởi động máy chủ Web App Tool KPI...
echo Trình duyệt sẽ tự động mở tại http://localhost:8765
echo.

set UV_PATH="C:\Users\10102\.local\bin\uv.exe"
set PYTHON_PATH="C:\Users\10102\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist %UV_PATH% (
    %UV_PATH% run --with openpyxl --with pandas python "%~dp0app.py"
) else if exist %PYTHON_PATH% (
    %PYTHON_PATH% "%~dp0app.py"
) else (
    python "%~dp0app.py"
)

pause
