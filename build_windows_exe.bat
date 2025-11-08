@echo off
echo Building Windows executable for Roster Generator...
echo.

REM Clean up previous builds
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del "*.spec"

echo Creating executable...
pyinstaller --onefile --name "RosterGenerator" --add-data "*.xlsx;." roster_claude.py

echo.
echo Build complete! 
echo The executable is located in the 'dist' folder.
echo You can copy 'dist\RosterGenerator.exe' to any Windows computer and run it.
echo.
pause
