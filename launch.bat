@echo off
REM ============================================================
REM  Email App — двойной клик для запуска GUI (Windows)
REM  Требования: Python 3.10+ установлен, venv создан
REM ============================================================

cd /d "%~dp0"

REM Проверяем наличие .venv
if not exist ".venv\Scripts\activate.bat" (
    echo [ОШИБКА] Виртуальное окружение не найдено.
    echo.
    echo Создайте его один раз:
    echo   python -m venv .venv
    echo   .venv\Scripts\activate
    echo   pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

REM Копируем конфиг-пример, если конфига нет
if not exist "config\settings.yaml" (
    copy "config\settings.example.yaml" "config\settings.yaml" >nul
    echo [INFO] Создан config\settings.yaml из примера.
    echo Заполните SMTP-данные перед отправкой.
    echo.
)

call .venv\Scripts\activate.bat
pythonw -m email_app --modern-gui
