@echo off
chcp 65001 >nul
title تحديث الداشبورد - أعمال المختصون
cd /d "%~dp0"
echo ============================================
echo    تحديث الداشبورد - أعمال المختصون
echo ============================================
echo.
python reader.py
if errorlevel 1 (
  py reader.py
)
if errorlevel 1 (
  echo.
  echo [خطأ] تعذر تشغيل Python.
  echo ثبّت Python وتأكد من توفر الحزم التالية: pandas , openpyxl
  echo قم بتشغيل هذا الأمر لتثبيتها:
  echo    pip install pandas openpyxl
  pause
  exit /b 1
)
echo.
echo تم تحديث بيانات الداشبورد بنجاح... جارٍ فتح اللوحة
start "" "%~dp0index.html"
timeout /t 3 >nul
exit /b 0