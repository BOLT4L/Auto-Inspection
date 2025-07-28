
@echo off
echo ======================================
echo Car Inspection Management System Setup
echo ======================================
echo.

echo Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://python.org
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo Python found!
echo.

echo Installing required packages...
pip install pyodbc
if %errorlevel% neq 0 (
    echo ERROR: Failed to install pyodbc
    echo Please check your internet connection and try again
    pause
    exit /b 1
)

echo.
echo Testing database connectivity...
python database_utils.py

echo.
echo ======================================
echo Setup Complete!
echo ======================================
echo.
echo To start the application, run: python main.py
echo Or double-click on main.py
echo.
echo First time users:
echo 1. Click "Select Database" to create a new database
echo 2. Choose a location and filename for your database
echo 3. Start adding inspection records!
echo.
pause
