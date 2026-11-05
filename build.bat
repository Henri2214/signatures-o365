@echo off
setlocal
REM ==========================================================
REM  Construit SignaturesO365.exe (portable, un seul fichier)
REM  - environnement Python neuf a chaque fois (Python 3.12,
REM    installe si absent) ;
REM  - tests avant construction ;
REM  - construction dans un dossier temporaire : si quelque
REM    chose echoue, l'ancien dist\SignaturesO365.exe reste ;
REM  - en cas de succes, dist\ ne contient plus que le nouvel
REM    exe et son empreinte SHA-256. Les dossiers dist\data
REM    (config, connexion) et dist\exports (sauvegardes,
REM    journaux) sont conserves ;
REM  - tous les fichiers temporaires sont supprimes a la fin.
REM ==========================================================
cd /d "%~dp0"
set "PYTHONDONTWRITEBYTECODE=1"
set "TMPDIST=build\dist-new"

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
    goto :err
)
echo Python utilise : %PYEXE%

echo.
echo [1/5] Suppression des restes des constructions precedentes...
call :cleanup

echo.
echo [2/5] Creation d'un environnement propre...
"%PYEXE%" -m venv .venv || goto :err
set "VPY=%~dp0.venv\Scripts\python.exe"
"%VPY%" -m pip install --upgrade pip --quiet
"%VPY%" -m pip install -r requirements-dev.txt --quiet || goto :err
"%VPY%" -c "import pymsalruntime; print('pymsalruntime OK')"
if errorlevel 1 (
    echo ERREUR : pymsalruntime ne s'installe pas avec ce Python.
    goto :err
)

echo.
echo [3/5] Tests automatiques...
"%VPY%" -m pytest -q -p no:cacheprovider tests
if errorlevel 1 (
    echo ERREUR : des tests echouent, construction annulee.
    goto :err
)

echo.
echo [4/5] Construction de l'exe dans un dossier temporaire...
"%VPY%" -m PyInstaller --noconfirm --clean --onefile --windowed --collect-all pymsalruntime ^
    --name SignaturesO365 --distpath "%TMPDIST%" --workpath build\work --specpath build ^
    sig_o365.py || goto :err
if not exist "%TMPDIST%\SignaturesO365.exe" (
    echo ERREUR : l'exe n'a pas ete produit.
    goto :err
)

echo.
echo [5/5] Remplacement de l'ancien exe dans dist\...
if not exist dist mkdir dist
copy /y "%TMPDIST%\SignaturesO365.exe" "dist\SignaturesO365.exe" >nul
if errorlevel 1 (
    echo ERREUR : impossible de remplacer dist\SignaturesO365.exe.
    echo Fermez SignaturesO365 s'il est ouvert, puis relancez build.bat.
    goto :err
)
REM dist\ : on ne garde que l'exe, son empreinte, data\ et exports\
for %%F in (dist\*) do if /i not "%%~nxF"=="SignaturesO365.exe" del /q "%%F"
for /d %%D in (dist\*) do if /i not "%%~nxD"=="data" if /i not "%%~nxD"=="exports" rmdir /s /q "%%D"
powershell -NoProfile -Command "$h=(Get-FileHash 'dist\SignaturesO365.exe' -Algorithm SHA256).Hash; ($h + '  SignaturesO365.exe') | Out-File -Encoding ascii 'dist\SignaturesO365.exe.sha256'; Write-Host ('SHA-256 : ' + $h)"

call :cleanup

echo.
echo ==========================================================
echo  OK : dist\SignaturesO365.exe
echo  Cet exe fonctionne sur n'importe quel PC Windows sans Python.
echo  dist\data et dist\exports ont ete conserves s'ils existaient.
echo ==========================================================
pause
exit /b 0

:err
echo.
echo ERREUR pendant la construction, voir les messages ci-dessus.
echo L'exe precedent (dist\SignaturesO365.exe), s'il existait, n'a pas ete modifie.
call :cleanup
pause
exit /b 1

REM ---------------------------------------------------------- nettoyage
:cleanup
if exist build rmdir /s /q build
if exist .venv rmdir /s /q .venv
if exist SignaturesO365.spec del /q SignaturesO365.spec
if exist .pytest_cache rmdir /s /q .pytest_cache
for /d /r %%D in (__pycache__) do if exist "%%D" rmdir /s /q "%%D"
exit /b 0
