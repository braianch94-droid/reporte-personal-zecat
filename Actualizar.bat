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
echo Actualizando GitHub Pages...
copy /Y "%~dp0Reporte_Personal.html" "%~dp0index.html" >nul
cd /d "%~dp0"
git add Reporte_Personal.html index.html
git commit -m "Actualizar reporte %DATE% %TIME%"
git push origin master

echo.
echo Abriendo reporte actualizado...
start "" "%~dp0Reporte_Personal.html"

echo.
echo Listo! El reporte tambien se subio a GitHub Pages.
timeout /t 3 /nobreak >nul
