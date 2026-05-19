@echo off
title Instalador - Conciliacion Bancaria Seikou SA
cd /d "%~dp0"
color 0A

echo.
echo  ================================================
echo   INSTALADOR - Conciliacion Bancaria Seikou SA
echo  ================================================
echo.

:: Buscar Python con varios nombres posibles
set PYTHON=
python --version >nul 2>&1
if %errorlevel%==0 set PYTHON=python

if "%PYTHON%"=="" (
    py --version >nul 2>&1
    if %errorlevel%==0 set PYTHON=py
)

if "%PYTHON%"=="" (
    python3 --version >nul 2>&1
    if %errorlevel%==0 set PYTHON=python3
)

:: Buscar en rutas comunes si todavia no encontro
if "%PYTHON%"=="" (
    for %%P in (
        "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
        "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
        "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
        "C:\Python312\python.exe"
        "C:\Python311\python.exe"
        "C:\Program Files\Python312\python.exe"
        "C:\Program Files\Python311\python.exe"
    ) do (
        if exist %%P (
            set PYTHON=%%P
            goto :encontrado
        )
    )
)

:encontrado
if "%PYTHON%"=="" (
    color 0C
    echo  [ERROR] No se encontro Python en este equipo.
    echo.
    echo  1. Ve a: https://www.python.org/downloads/
    echo  2. Descarga e instala Python
    echo  3. IMPORTANTE: marca "Add Python to PATH" al instalar
    echo  4. Reinicia el equipo
    echo  5. Vuelve a ejecutar este instalador
    echo.
    pause
    start "" "https://www.python.org/downloads/"
    exit
)

for /f "tokens=*" %%v in ('%PYTHON% --version 2^>^&1') do set PYVER=%%v
echo  Python encontrado: %PYVER%
echo  Ruta: %PYTHON%
echo.
echo  Instalando librerias (puede tardar 1-2 minutos)...
echo.

%PYTHON% -m pip install --upgrade pip --quiet
%PYTHON% -m pip install streamlit pandas pdfplumber openpyxl reportlab

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo  [ERROR] Fallo la instalacion de librerias.
    echo  Verifica que tengas internet y vuelve a intentar.
    pause
    exit
)

:: Guardar la ruta de Python para que run.bat la use
echo %PYTHON%> "%~dp0python_path.txt"

echo.
color 0A
echo  ================================================
echo   Instalacion completada!
echo   Ahora cierra esta ventana y abre: run.bat
echo  ================================================
echo.
pause
