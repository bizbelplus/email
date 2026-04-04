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
  --noconfirm ^
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
  --hidden-import "socks" ^
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

  REM Гарантируем структуру portable-папки даже если каталоги пустые
  set "DIST_DIR=dist\Email App"
  if not exist "%DIST_DIR%\config" mkdir "%DIST_DIR%\config"
  if not exist "%DIST_DIR%\templates" mkdir "%DIST_DIR%\templates"
  if not exist "%DIST_DIR%\files" mkdir "%DIST_DIR%\files"
  if not exist "%DIST_DIR%\presets" mkdir "%DIST_DIR%\presets"
  if not exist "%DIST_DIR%\preview" mkdir "%DIST_DIR%\preview"
  if not exist "%DIST_DIR%\history" mkdir "%DIST_DIR%\history"
  if not exist "%DIST_DIR%\logs" mkdir "%DIST_DIR%\logs"

  REM Кладем portable-ресурсы рядом с .exe, потому что приложение ищет их в корне dist\Email App
  if exist "templates" xcopy "templates\*" "%DIST_DIR%\templates\" /E /I /Y >nul
  if exist "config" xcopy "config\*" "%DIST_DIR%\config\" /E /I /Y >nul
  if exist "files" xcopy "files\*" "%DIST_DIR%\files\" /E /I /Y >nul
  if exist "presets" xcopy "presets\*" "%DIST_DIR%\presets\" /E /I /Y >nul
  if exist "preview" xcopy "preview\*" "%DIST_DIR%\preview\" /E /I /Y >nul
  if exist "history" xcopy "history\*" "%DIST_DIR%\history\" /E /I /Y >nul
  if exist "logs" xcopy "logs\*" "%DIST_DIR%\logs\" /E /I /Y >nul
  if exist "recipients.csv" copy /Y "recipients.csv" "%DIST_DIR%\recipients.csv" >nul
  if exist "README.md" copy /Y "README.md" "%DIST_DIR%\README.md" >nul
  if exist "GUIDE.md" copy /Y "GUIDE.md" "%DIST_DIR%\GUIDE.md" >nul
  if exist "*.pdf" copy /Y "*.pdf" "%DIST_DIR%\" >nul
  if exist "*.docx" copy /Y "*.docx" "%DIST_DIR%\" >nul

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
