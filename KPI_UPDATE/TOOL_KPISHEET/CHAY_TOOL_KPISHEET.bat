@echo off
chcp 65001 > nul
title MEDIGO - TOOL KHOI TAO GOI KPI
echo ========================================================
echo        MEDIGO - TOOL KHOI TAO GOI KPI TU DE XUAT
echo ========================================================
echo.
echo Dang khoi dong Web App tai http://localhost:8767 ...
echo.

start "" "http://localhost:8767"

set UV_PATH="C:\Users\10102\.local\bin\uv.exe"
if exist %UV_PATH% (
    %UV_PATH% run --with openpyxl --with python-docx --with python-pptx python "%~dp0app.py"
) else (
    python "%~dp0app.py"
)

pause
