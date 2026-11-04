@echo off
setlocal
REM ==========================================================
REM  Construit SignaturesO365.exe (portable, un seul fichier)
REM  Utilise Python 3.12 (compatible avec le composant de
REM  connexion Windows pymsalruntime). L'installe si absent.
REM ==========================================================
cd /d "%~dp0"

set "PYEXE="
for /f "delims=" %%i in ('py -3.12 -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%i"
if not defined PYEXE if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"

if not defined PYEXE (
    echo Python 3.12 absent : installation via winget, pour l'utilisateur, sans droits admin...
    winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
)
if not defined PYEXE (
    echo.
    echo ERREUR : Python 3.12 introuvable.
    echo Installez Python 3.12 depuis https://www.python.org/downloads/ puis relancez.
    pause
    exit /b 1
)
echo Python utilise : %PYEXE%

echo Creation d'un environnement propre...
if exist .venv rmdir /s /q .venv
"%PYEXE%" -m venv .venv || goto :err
set "VPY=%~dp0.venv\Scripts\python.exe"

"%VPY%" -m pip install --upgrade pip
"%VPY%" -m pip install -r requirements.txt || goto :err

echo Verification du composant de connexion Windows...
"%VPY%" -c "import pymsalruntime; print('pymsalruntime OK')"
if errorlevel 1 (
    echo ERREUR : pymsalruntime ne s'installe pas avec ce Python.
    goto :err
)

echo Nettoyage des anciennes constructions...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist SignaturesO365.spec del /q SignaturesO365.spec

"%VPY%" -m PyInstaller --noconfirm --clean --onefile --windowed --collect-all pymsalruntime --name SignaturesO365 sig_o365.py || goto :err

echo.
echo ==========================================================
echo  OK : dist\SignaturesO365.exe
echo  Cet exe fonctionne sur n'importe quel PC Windows sans Python.
echo ==========================================================
pause
exit /b 0

:err
echo.
echo ERREUR pendant la construction, voir les messages ci-dessus.
pause
exit /b 1
