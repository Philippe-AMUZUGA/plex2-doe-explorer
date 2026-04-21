@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "APPDIR=%~dp0"
set "RUNTIME=%APPDIR%.runtime"
set "PYW=%RUNTIME%\Scripts\pythonw.exe"
set "PYC=%RUNTIME%\Scripts\python.exe"
set "REQ=%APPDIR%requirements.txt"
set "SCRIPT=%APPDIR%app\PLEX2_Launcher.py"
set "PYBOOT="

if exist "%PYW%" (
    start "" "%PYW%" "%SCRIPT%"
    exit /b 0
)

if exist "%PYC%" (
    start "" "%PYC%" "%SCRIPT%"
    exit /b 0
)

call :find_python
if not defined PYBOOT goto no_python

echo [PLEX2] Initialisation du runtime local Windows...
%PYBOOT% -m venv "%RUNTIME%"
if errorlevel 1 goto venv_error

call "%PYC%" -m pip install --upgrade pip
if errorlevel 1 goto pip_error

call "%PYC%" -m pip install -r "%REQ%"
if errorlevel 1 goto pip_error

start "" "%PYW%" "%SCRIPT%"
exit /b 0

:find_python
where py >nul 2>&1
if %errorlevel%==0 (
    set "PYBOOT=py -3"
    goto :eof
)
where python >nul 2>&1
if %errorlevel%==0 (
    set "PYBOOT=python"
    goto :eof
)
goto :eof

:no_python
powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python 3.10 ou supérieur n''a pas été détecté sur cette machine Windows. Installez Python, puis relancez Lancer_PLEX2.bat.', 'PLEX2')" >nul 2>&1
exit /b 1

:venv_error
powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('La création du runtime local PLEX2 a échoué.', 'PLEX2')" >nul 2>&1
exit /b 1

:pip_error
powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('L''installation des dépendances a échoué. Vérifiez l''accès réseau puis relancez Lancer_PLEX2.bat.', 'PLEX2')" >nul 2>&1
exit /b 1
