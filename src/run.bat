@echo off
echo ========================================================
echo    Running... Gerardo Zermeño - Suite de Utilidades
echo ========================================================
echo.

if not exist "C:\Temp" (
    mkdir C:\Temp
    echo Carpeta C:\Temp creada
)

powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0main.ps1"

if %errorlevel% neq 0 (
    echo.
    echo Error al ejecutar la herramienta
    pause
)
