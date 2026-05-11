@echo off
chcp 65001 > nul

echo ============================
echo      STARTING PROJECT
echo ============================
echo.

REM Создание venv
if not exist venv (
    echo [1/4] Creating virtual environment...
    python -m venv venv
) else (
    echo [1/4] venv already exists
)

echo.

REM Активация venv
echo [2/4] Activating virtual environment...
call venv\Scripts\activate

echo.

REM Установка зависимостей
echo [3/4] Installing requirements...
pip install -r requirements.txt

echo.

REM Запуск проекта
echo [4/4] Starting main.py...
python main.py

echo.
echo ============================
echo      SCRIPT FINISHED
echo ============================
echo.

pause