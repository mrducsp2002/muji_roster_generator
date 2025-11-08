@echo off
title Build Windows Executable
echo.
echo ========================================
echo   Building Roster Generator for Windows
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://python.org
    echo.
    pause
    exit /b 1
)

echo Python found! Installing PyInstaller...
python -m pip install pyinstaller

echo.
echo Building executable...
python build_windows_exe.py

echo.
echo Build process complete!
echo Check the output above for any errors.
echo.
pause
