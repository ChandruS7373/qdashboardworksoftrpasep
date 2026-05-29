@echo off
echo Installing QDashboard requirements...
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo.
    echo Install failed. Make sure Python and pip are installed and in PATH.
    pause
    exit /b 1
)
echo.
echo All requirements installed successfully.
pause
