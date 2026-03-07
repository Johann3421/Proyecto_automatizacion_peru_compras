@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Build instalador Peru Compras Bot
echo ============================================================

if not exist "dist\peru_compras_bot.exe" (
    echo ERROR: No existe dist\peru_compras_bot.exe
    echo Primero genera el ejecutable con build_exe.bat
    exit /b 1
)

where ISCC >nul 2>&1
if errorlevel 1 (
    set "ISCC_EXE=C:\Program Files\Inno Setup 6\ISCC.exe"
    if not exist "%ISCC_EXE%" set "ISCC_EXE=%USERPROFILE%\AppData\Local\Programs\Inno Setup 6\ISCC.exe"
    if not exist "%ISCC_EXE%" (
        for /f "usebackq delims=" %%I in (`powershell -NoProfile -Command "(Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue ^| Where-Object { $_.DisplayName -like '*Inno Setup*' } ^| Select-Object -First 1 -ExpandProperty InstallLocation) + 'ISCC.exe'"`) do set "ISCC_EXE=%%I"
    )
    if not exist "%ISCC_EXE%" (
        echo ERROR: Inno Setup no esta instalado o ISCC no esta en PATH.
        echo Instala Inno Setup 6: https://jrsoftware.org/isinfo.php
        echo Luego vuelve a ejecutar este archivo.
        exit /b 1
    )
) else (
    set "ISCC_EXE=ISCC"
)

echo Compilando instalador con Inno Setup...
"%ISCC_EXE%" installer.iss
if errorlevel 1 (
    echo ERROR: Fallo la compilacion del instalador.
    exit /b 1
)

echo.
echo OK: Instalador generado en installer_output\PeruComprasBot_Setup.exe
exit /b 0
