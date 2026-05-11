@echo off
chcp 65001 > nul
setlocal

set REPO_URL=https://github.com/nikitakuvsh/excelMacros/archive/refs/heads/main.zip
set ZIP_FILE=repo.zip
set TEMP_DIR=_temp_extract

echo ============================
echo SETUP PROJECT
echo ============================
echo.

REM ============================
REM DOWNLOAD
REM ============================

echo [1/3] Downloading repository...

powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest '%REPO_URL%' -OutFile '%ZIP_FILE%'"

if not exist "%ZIP_FILE%" (
    echo ERROR: Download failed
    pause
    exit /b
)

echo Download OK
echo.

REM ============================
REM EXTRACT
REM ============================

echo [2/3] Extracting...

if exist "%TEMP_DIR%" rmdir /s /q "%TEMP_DIR%"
mkdir "%TEMP_DIR%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive '%ZIP_FILE%' '%TEMP_DIR%' -Force"

if not exist "%TEMP_DIR%\excelMacros-main" (
    echo ERROR: Extract failed
    pause
    exit /b
)

echo Moving files...

xcopy "%TEMP_DIR%\excelMacros-main\*" "." /E /H /Y > nul

REM cleanup ALWAYS safe
if exist "%TEMP_DIR%" rmdir /s /q "%TEMP_DIR%"
if exist "%ZIP_FILE%" del "%ZIP_FILE%"

echo.

REM ============================
REM INSTALL DEPENDENCIES
REM ============================

echo [3/3] Installing dependencies...

if exist requirements.txt (
    python -m venv venv
    call venv\Scripts\python.exe -m pip install --upgrade pip
    call venv\Scripts\python.exe -m pip install -r requirements.txt
) else (
    echo No requirements.txt found
)

echo.
echo ============================
echo DONE
echo ============================
pause
endlocal