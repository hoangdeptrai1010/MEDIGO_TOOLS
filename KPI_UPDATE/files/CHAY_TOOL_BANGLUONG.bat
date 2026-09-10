@echo off
chcp 65001 > nul
title MEDIGO PAYROLL AUTOMATION TOOL

echo ========================================================
echo        HỆ THỐNG TỰ ĐỘNG HÓA BẢNG LƯƠNG - MEDIGO
echo ========================================================
echo.
echo Vui lòng chọn tháng cần tính Bảng Lương:
echo   [1] Bảng Lương Tháng 8 (Mới nhất)
echo   [2] Bảng Lương Tháng 7
echo   [3] Chạy tùy chỉnh theo tham số
echo.

set /p choice="Nhập lựa chọn của bạn (1/2/3): "

set PYTHON_PATH="C:\Users\10102\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist %PYTHON_PATH% (
    set PYTHON_PATH=python
)

if "%choice%"=="1" (
    echo.
    echo --> Đang tính Bảng Lương Tháng 8...
    %PYTHON_PATH% "%~dp0build_payroll_report.py" --month 8 --kpi "%~dp0thang8\baocaokpi_thang8_hoanthien.xlsx" --output "%~dp0thang8\bangluong_thang8_hoanthien.xlsx"
) else if "%choice%"=="2" (
    echo.
    echo --> Đang tính Bảng Lương Tháng 7...
    %PYTHON_PATH% "%~dp0build_payroll_report.py" --month 7 --kpi "%~dp0thang7\baocaokpi_thang7_hoanthien.xlsx" --output "%~dp0thang7\bangluong_thang7_hoanthien.xlsx"
) else (
    echo.
    echo Chế độ tùy chỉnh: Vui lòng chạy lệnh trong terminal.
)

echo.
echo ========================================================
echo Hoàn thành! Nhấn phím bất kỳ để thoát.
pause > nul
