@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PY_CMD="
py -3 --version > build_log.txt 2>&1
if not errorlevel 1 set "PY_CMD=py -3"
if "%PY_CMD%"=="" (
  python --version >> build_log.txt 2>&1
  if not errorlevel 1 set "PY_CMD=python"
)

if "%PY_CMD%"=="" (
  echo Python 3 was not found.
  pause
  exit /b 1
)

echo Installing required packages...
%PY_CMD% -m pip install -r requirements.txt >> build_log.txt 2>&1
if errorlevel 1 (
  echo Package installation failed. Check build_log.txt.
  pause
  exit /b 1
)

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist integrated_counsel_program_v1.spec del /q integrated_counsel_program_v1.spec

echo Building unified program...
%PY_CMD% -m PyInstaller --noconfirm --clean --onedir --windowed ^
  --name "integrated_counsel_program_v1" ^
  --add-data "selection_standard_2026.pdf;." ^
  --add-data "basic_pension_review_template.xlsx;." ^
  --hidden-import team1_app ^
  --hidden-import team2_app ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32com.client ^
  launcher.py >> build_log.txt 2>&1

if errorlevel 1 (
  echo Build failed. Check build_log.txt.
  pause
  exit /b 1
)

echo.
echo Build complete.
echo Run: dist\integrated_counsel_program_v1\integrated_counsel_program_v1.exe
pause
