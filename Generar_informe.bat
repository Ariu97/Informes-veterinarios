@echo off
REM Ejecuta el generador desde la carpeta del proyecto, independientemente del cwd actual
cd /d "%~dp0"
python generar_informe.py
echo.
echo Pulsa una tecla para cerrar...
pause >nul