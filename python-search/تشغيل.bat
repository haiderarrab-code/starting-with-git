@echo off
chcp 65001 >nul
title نظام البحث الموحد - قسم المعلومات المركزية
echo ==================================================
echo   نظام البحث الموحد - قسم المعلومات المركزية
echo ==================================================
echo.
where python >nul 2>nul
if errorlevel 1 (
  echo [!] لم يتم العثور على Python.
  echo     حمله من: https://www.python.org/downloads/
  echo     وفعل خيار "Add Python to PATH" اثناء التنصيب.
  pause
  exit /b
)
echo [*] التحقق من المكتبات...
python -c "import pandas, openpyxl" 2>nul
if errorlevel 1 (
  echo [*] تنصيب المكتبات لاول مرة (يحتاج انترنت)...
  python -m pip install pandas openpyxl
)
echo [*] تشغيل البرنامج...
python "%~dp0بحث_قاعدة_البيانات.py"
if errorlevel 1 pause
