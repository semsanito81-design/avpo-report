@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python не найден. Установите Python 3.10+ с python.org
    pause
    exit /b 1
)

if not exist ".venv" (
    echo Создаю виртуальное окружение...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    echo Устанавливаю зависимости...
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
) else (
    call .venv\Scripts\activate.bat
)

python avpo.py
pause