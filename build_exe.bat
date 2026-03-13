@echo off
setlocal

REM Build script for Peru Compras Bot GUI (Windows)
REM Generates a GUI-only .exe (without terminal window)

cd /d "%~dp0"

set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

echo [1/3] Installing dependencies...
"%PYTHON_EXE%" -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [2/3] Cleaning previous build artifacts...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [3/3] Building executable...
"%PYTHON_EXE%" -m PyInstaller --noconfirm peru_compras_bot.spec
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