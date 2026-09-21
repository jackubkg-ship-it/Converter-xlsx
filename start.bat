@echo off
title Format Konvertoru
cd /d "%~dp0"

echo ============================================
echo   Format Konvertoru ishe salinir...
echo ============================================
echo.

set "PYTHON="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON=py -3"
if "%PYTHON%"=="" (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON=python"
)

if "%PYTHON%"=="" (
    echo [XETA] Python tapilmadi. Qurashdirin: https://www.python.org/downloads/
    echo Qurashdirma zamani "Add python.exe to PATH" qutusunu isaretleyin.
    pause
    exit /b 1
)

if not exist ".venv" (
    echo Ilk defe ishe salinir, muhit hazirlanir...
    %PYTHON% -m venv .venv
)

call ".venv\Scripts\activate.bat"

echo Lazimi paketler yoxlanilir...
python -m pip install --quiet --disable-pip-version-check --upgrade pip
python -m pip install --quiet --disable-pip-version-check -r requirements.txt

echo.
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://localhost:8010"

echo ============================================
echo   Panel bu unvanda achilacaq: http://localhost:8010
echo   BU PENCERENI BAGLAMAYIN.
echo ============================================
echo.

python -m uvicorn main:app --host 127.0.0.1 --port 8010

pause
