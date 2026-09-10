@echo off
chcp 65001 > nul
cd /d "%~dp0TOOL_KPI"
call CHAY_TOOL_KPI.bat
