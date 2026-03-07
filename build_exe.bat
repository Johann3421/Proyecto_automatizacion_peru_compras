@echo off
setlocal

REM Build script for Peru Compras Bot GUI (Windows)
REM Generates a GUI-only .exe (without terminal window)

cd /d "%~dp0"

echo [1/3] Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 goto :error

echo [2/3] Cleaning previous build artifacts...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist peru_compras_bot.spec del /f /q peru_compras_bot.spec

echo [3/3] Building executable...
pyinstaller --noconsole --onefile --add-data "productos.xlsx;." --exclude-module sqlalchemy --collect-all numpy --collect-all pandas peru_compras_bot.py
if errorlevel 1 goto :error

echo.
echo Build finished successfully.
echo Executable: dist\peru_compras_bot.exe
pause
exit /b 0

:error
echo.
echo Build failed.
pause
exit /b 1