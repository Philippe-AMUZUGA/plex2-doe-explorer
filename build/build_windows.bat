@echo off
setlocal
cd /d "%~dp0.."
py -3 -m pip install -r requirements.txt
if errorlevel 1 exit /b 1
py -3 -m pip install pyinstaller
if errorlevel 1 exit /b 1
py -3 -m PyInstaller build\pyinstaller_windows.spec --noconfirm --clean
