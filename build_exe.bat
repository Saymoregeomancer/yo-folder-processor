@echo off
rem Збирає process_folders.exe і release\yo-folder-processor-win64.zip для запуску без Python.
chcp 65001 >nul
cd /d "%~dp0"

rem чисте середовище — інакше PyInstaller затягує все з глобального Python
if not exist .venv-build\Scripts\python.exe python -m venv .venv-build || goto :error
set PY=.venv-build\Scripts\python.exe
%PY% -m pip install --quiet -r requirements.txt pyinstaller || goto :error
%PY% -m PyInstaller --noconfirm --clean --onefile --console --name process_folders ^
    --collect-submodules xlwings --exclude-module xlwings.pro process_folders.py || goto :error

set OUT=release\yo-folder-processor
if exist release rmdir /s /q release
mkdir "%OUT%"
copy /y dist\process_folders.exe "%OUT%\" >nul
for %%F in (process_folders.bat process_folders_dry_run.bat install_context_menu.bat install_context_menu.ps1 uninstall_context_menu.bat db.ini.example README.md) do copy /y %%F "%OUT%\" >nul
powershell -NoProfile -Command "Compress-Archive -Path 'release\yo-folder-processor' -DestinationPath 'release\yo-folder-processor-win64.zip' -Force" || goto :error

echo.
echo Готово: release\yo-folder-processor-win64.zip
goto :eof

:error
echo.
echo [ПОМИЛКА] Збірка не вдалась.
exit /b 1
