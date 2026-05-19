@echo off
title Conciliacion Bancaria - Seikou SA
cd /d "%~dp0"
color 0A

:: Leer la ruta de Python guardada por el instalador
set PYTHON=python
if exist "%~dp0python_path.txt" (
    set /p PYTHON=<"%~dp0python_path.txt"
)

:: Verificar que Python funciona
%PYTHON% --version >nul 2>&1
if %errorlevel% neq 0 (
    :: Intentar con py
    py --version >nul 2>&1
    if %errorlevel%==0 (
        set PYTHON=py
    ) else (
        color 0C
        echo  [ERROR] No se puede iniciar. Ejecuta instalar.bat primero.
        pause
        exit
    )
)

:: Verificar streamlit instalado
%PYTHON% -c "import streamlit" >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo  [ERROR] Librerias no instaladas. Ejecuta instalar.bat primero.
    pause
    exit
)

taskkill /f /im streamlit.exe >nul 2>&1
timeout /t 1 /nobreak >nul

echo.
echo  Iniciando aplicacion...
echo  El navegador abre en unos segundos.
echo  Para cerrar: cierra esta ventana.
echo.

%PYTHON% -m streamlit run "%~dp0app.py" --server.port 8501 --browser.gatherUsageStats false

pause
