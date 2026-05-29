@echo off
title Qualesce AI Agent
color 0A
echo.
echo  =====================================================
echo    Qualesce AI Project Manager
echo  =====================================================
echo.

echo  Stopping any previous instance...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8502 "') do (
    taskkill /F /PID %%a >nul 2>&1
)
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8503 "') do (
    taskkill /F /PID %%a >nul 2>&1
)
timeout /t 2 /nobreak >nul

echo  Starting app... please wait.
echo.
cd /d "%~dp0"
python -m streamlit run app.py --server.port=8502 --browser.gatherUsageStats=false
echo.
echo  App stopped. Press any key to close.
pause > nul
