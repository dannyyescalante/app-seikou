@echo off
title Conciliacion Bancaria - Acceso Remoto
cd /d "%~dp0"

echo.
echo  Iniciando app + acceso remoto...
echo.

:: Iniciar Streamlit en segundo plano
start "Streamlit" cmd /c "python -m streamlit run app.py --server.port 8501 --server.headless true --browser.gatherUsageStats false"

timeout /t 4 /nobreak >nul

:: Verificar si ngrok esta instalado
ngrok version >nul 2>&1
if %errorlevel% neq 0 (
    echo  Descargando ngrok...
    powershell -Command "Invoke-WebRequest -Uri 'https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-v3-stable-windows-amd64.zip' -OutFile 'ngrok.zip'"
    powershell -Command "Expand-Archive -Path 'ngrok.zip' -DestinationPath '.' -Force"
    del ngrok.zip >nul 2>&1
)

echo.
echo  ================================================
echo   La URL aparece abajo (linea "Forwarding")
echo   Comparte esa URL con quien necesite acceder
echo   La URL cambia cada vez que abres esto
echo  ================================================
echo.

ngrok http 8501

pause
