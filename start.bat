@echo off
chcp 65001 > nul
setlocal

echo ============================
echo STARTING PROJECT
echo ============================
echo.

REM check venv
if not exist venv (
    echo ERROR: venv not found. Run setup.bat first
    pause
    exit /b
)

echo [1/3] Installing requirements...

call venv\Scripts\python.exe -m pip install --upgrade pip
call venv\Scripts\python.exe -m pip install -r requirements.txt

if errorlevel 1 (
    echo ERROR: pip install failed
    pause
    exit /b
)

echo.
echo [2/3] Running project...

call venv\Scripts\python.exe main.py

echo.
echo ============================
echo FINISHED
echo ============================
pause
endlocal