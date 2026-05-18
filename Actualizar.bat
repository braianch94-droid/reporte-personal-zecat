@echo off
echo ================================================
echo  Actualizando Reporte de Personal...
echo ================================================
echo.

echo Cerrando Excel si esta abierto...
taskkill /f /im EXCEL.EXE 2>nul
timeout /t 2 /nobreak >nul

echo Generando reporte...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0GenerarHTML.ps1"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: No se pudo generar el reporte.
    echo Verificar que Asistencia.xlsx exista y no este bloqueado.
    pause
    exit /b 1
)

echo.
echo Abriendo reporte actualizado...
start "" "%~dp0Reporte_Personal.html"

echo Listo!
timeout /t 2 /nobreak >nul
