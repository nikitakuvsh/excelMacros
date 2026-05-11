@echo off
chcp 65001 > nul

set REPO_URL=https://github.com/nikitakuvsh/excelMacros/archive/refs/heads/main.zip
set ZIP_FILE=repo.zip
set FOLDER_NAME=excelMacros-main

echo ============================
echo     SETUP PROJECT
echo ============================
echo.

REM Скачивание архива
echo [1/3] Downloading repository...

powershell -Command ^
"Invoke-WebRequest -Uri '%REPO_URL%' -OutFile '%ZIP_FILE%'"

if not exist %ZIP_FILE% (
    echo ERROR: Download failed
    pause
    exit /b
)

echo Download complete.
echo.

REM Распаковка
echo [2/3] Extracting files...

powershell -Command ^
"Expand-Archive -Path '%ZIP_FILE%' -DestinationPath '.' -Force"

echo Extracted.
echo.

REM Очистка zip
del %ZIP_FILE%

REM Переход в папку проекта
cd %FOLDER_NAME%

echo [3/3] Installing dependencies...

if exist requirements.txt (
    python -m venv venv
    call venv\Scripts\activate
    pip install -r requirements.txt
) else (
    echo No requirements.txt found
)

echo.
echo ============================
echo        DONE
echo ============================
echo.

pause