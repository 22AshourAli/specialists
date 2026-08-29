@echo off
chcp 65001 >nul
title فتح الداشبورد - أعمال المختصون
cd /d "%~dp0"
start "" "%~dp0index.html"
exit /b 0