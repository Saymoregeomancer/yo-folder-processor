@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo ==========================================================
echo  Обробка папок
if not "%~1"=="" echo  %~1
echo ==========================================================
echo.

if exist "%~dp0process_folders.exe" (
    "%~dp0process_folders.exe" %*
) else (
    python process_folders.py %*
)

echo.
pause
