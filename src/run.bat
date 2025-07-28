
@echo off
echo Starting Car Inspection Management System...
python main.py
if %errorlevel% neq 0 (
    echo.
    echo An error occurred while running the application.
    echo Please check that all requirements are installed.
    echo Run install.bat if you haven't already.
    pause
)
