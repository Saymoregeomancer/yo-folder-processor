@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo ==========================================================
echo  Обробка папок в C:\Users\Dembe\Desktop\content\ter
echo ==========================================================
echo.

python process_folders.py %*

echo.
pause
