@echo off
REM ============================================================
REM  build_windows.bat — сборка standalone Email App.exe
REM  Запустить ОДИН РАЗ на Windows-машине.
REM  Результат: dist\Email App\Email App.exe
REM ============================================================

cd /d "%~dp0"

echo === Email App — сборка .exe через PyInstaller ===
echo.

REM Активируем venv
if not exist ".venv\Scripts\activate.bat" (
    echo [ОШИБКА] Сначала создайте venv и установите зависимости:
    echo   python -m venv .venv
    echo   .venv\Scripts\activate
    echo   pip install -r requirements.txt
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

REM Устанавливаем PyInstaller если нет
python -c "import PyInstaller" 2>nul || pip install pyinstaller

echo.
echo Собираем...
echo.

pyinstaller ^
  --onedir ^
  --windowed ^
  --name "Email App" ^
  --add-data "templates;templates" ^
  --add-data "config;config" ^
  --add-data "files;files" ^
  --add-data "presets;presets" ^
  --add-data "recipients.csv;." ^
  --collect-all customtkinter ^
  --collect-all tkhtmlview ^
  --hidden-import "email_app.gui" ^
  --hidden-import "email_app.modern_gui" ^
  --hidden-import "email_app.service" ^
  --hidden-import "email_app.smtp_client" ^
  --hidden-import "email_app.config" ^
  --hidden-import "email_app.models" ^
  --hidden-import "email_app.recipients" ^
  --hidden-import "email_app.renderer" ^
  --hidden-import "email_app.presets" ^
  --hidden-import "email_app.stats" ^
  --hidden-import "email_app.campaign_queue" ^
  email_app\__main__.py

if %errorlevel% neq 0 (
    echo.
    echo [ОШИБКА] Сборка завершилась с ошибкой.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Готово! Приложение: dist\Email App\Email App.exe
echo.
echo  ВАЖНО: папка dist\Email App\ содержит всё необходимое.
echo  Скопируйте её целиком — не только .exe.
echo.
echo  Перед запуском убедитесь, что рядом с .exe есть:
echo    config\settings.yaml  (заполненный SMTP-конфиг)
echo    recipients.csv        (список получателей)
echo ============================================================
echo.
pause
