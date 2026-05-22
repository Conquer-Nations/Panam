@echo off
REM Re-extract the customer workbook to data.json. Run this after you update shipments in the .xlsx.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0refresh-data.ps1"
pause
