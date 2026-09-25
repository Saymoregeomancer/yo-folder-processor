@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo ==========================================================
echo  ПЕРЕВІРКА (dry-run) - нічого не змінюється
echo ==========================================================
echo.

python process_folders.py --dry-run %*

echo.
pause
