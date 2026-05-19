@echo off
title Conciliacion Bancaria - Seikou SA (RED LOCAL)
cd /d "%~dp0"

taskkill /f /im streamlit.exe >nul 2>&1
timeout /t 2 /nobreak >nul

echo.
echo  ========================================
echo   Conciliacion Bancaria - Seikou SA
echo   Acceso desde otros equipos en la red:
echo  ========================================

for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do (
    set IP=%%a
    goto :found
)
:found
set IP=%IP: =%
echo.
echo   URL para compartir: http://%IP%:8501
echo.
echo  ========================================
echo.

streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --browser.gatherUsageStats false
pause
