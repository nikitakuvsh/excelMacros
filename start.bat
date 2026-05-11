@echo off
chcp 65001 > nul
setlocal

echo ============================
echo STARTING PROJECT
echo ============================
echo.

REM ============================
REM 1. CHECK / CREATE venv
REM ============================

if not exist venv (
    echo [1/4] venv not found. Creating...

    python -m venv venv

    if errorlevel 1 (
        echo ERROR: Failed to create venv
        pause
        exit /b
    )

    echo venv created.
) else (
    echo [1/4] venv already exists
)

echo.

REM ============================
REM 2. INSTALL DEPENDENCIES
REM ============================

echo [2/4] Installing requirements...

call venv\Scripts\python.exe -m pip install --upgrade pip
call venv\Scripts\python.exe -m pip install -r requirements.txt

if errorlevel 1 (
    echo ERROR: pip install failed
    pause
    exit /b
)

echo.

REM ============================
REM 3. RUN PROJECT
REM ============================

echo [3/4] Running main.py...

call venv\Scripts\python.exe main.py

echo.

echo ============================
echo FINISHED
echo ============================
pause

endlocal