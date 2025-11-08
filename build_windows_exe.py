#!/usr/bin/env python3
"""
Windows Executable Builder for Roster Generator
Run this script on a Windows machine to create the executable
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def build_windows_executable():
    """Build Windows executable using PyInstaller"""
    
    print("🪟 Building Windows Executable for Roster Generator")
    print("=" * 60)
    
    # Check if we're on Windows
    if sys.platform != "win32":
        print("⚠️  Warning: This script is designed for Windows.")
        print("   You're currently on:", sys.platform)
        print("   The executable will be created for your current platform.")
        print()
    
    # Clean up previous builds
    print("🧹 Cleaning up previous builds...")
    for dir_name in ["dist", "build"]:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"   ✅ Removed {dir_name}/")
    
    for spec_file in Path(".").glob("*.spec"):
        spec_file.unlink()
        print(f"   ✅ Removed {spec_file}")
    
    # Check if PyInstaller is installed
    print("\n🔍 Checking PyInstaller installation...")
    try:
        import PyInstaller
        print(f"   ✅ PyInstaller {PyInstaller.__version__} is installed")
    except ImportError:
        print("   ❌ PyInstaller not found. Installing...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], 
                         check=True, capture_output=True)
            print("   ✅ PyInstaller installed successfully")
        except subprocess.CalledProcessError as e:
            print(f"   ❌ Failed to install PyInstaller: {e}")
            return False
    
    # Create the executable
    print("\n📦 Creating executable...")
    try:
        # PyInstaller command for Windows
        cmd = [
            "pyinstaller",
            "--onefile",                    # Single file executable
            "--name", "RosterGenerator",    # Name of the executable
            "--console",                    # Keep console window for input/output
            "--hidden-import", "pandas",    # Ensure pandas is included
            "--hidden-import", "openpyxl",  # Ensure openpyxl is included
            "--hidden-import", "xlrd",      # Additional Excel support
            "--clean",                      # Clean cache
            "roster_claude.py"
        ]
        
        print("   Running PyInstaller...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("   ✅ Executable created successfully!")
        else:
            print("   ❌ Error creating executable:")
            print(result.stderr)
            return False
            
    except Exception as e:
        print(f"   ❌ Error: {e}")
        return False
    
    # Check if executable was created
    exe_name = "RosterGenerator.exe" if sys.platform == "win32" else "RosterGenerator"
    exe_path = f"dist/{exe_name}"
    
    if os.path.exists(exe_path):
        file_size = os.path.getsize(exe_path) / (1024 * 1024)  # Size in MB
        print(f"\n🎉 Build Complete!")
        print(f"📁 Executable location: {os.path.abspath(exe_path)}")
        print(f"📏 File size: {file_size:.1f} MB")
        
        # Create a simple test
        print(f"\n🧪 Testing executable...")
        try:
            # Quick test - just check if it starts without errors
            test_cmd = [exe_path, "--help"] if sys.platform == "win32" else [exe_path, "--help"]
            test_result = subprocess.run(test_cmd, capture_output=True, text=True, timeout=10)
            if test_result.returncode == 0 or "usage" in test_result.stdout.lower():
                print("   ✅ Executable starts successfully")
            else:
                print("   ⚠️  Executable may have issues, but it was created")
        except Exception as e:
            print(f"   ⚠️  Could not test executable: {e}")
        
        return True
    else:
        print(f"   ❌ Executable not found at {exe_path}")
        return False

def create_distribution_package():
    """Create a complete distribution package"""
    
    print("\n📦 Creating distribution package...")
    
    # Create package directory
    package_name = "RosterGenerator_Windows"
    if os.path.exists(package_name):
        shutil.rmtree(package_name)
    os.makedirs(package_name)
    
    # Copy executable
    exe_name = "RosterGenerator.exe" if sys.platform == "win32" else "RosterGenerator"
    exe_source = f"dist/{exe_name}"
    exe_dest = f"{package_name}/{exe_name}"
    
    if os.path.exists(exe_source):
        shutil.copy2(exe_source, exe_dest)
        print(f"   ✅ Copied executable to {package_name}/")
    
    # Create README for Windows users
    readme_content = """# Roster Generator for Windows

## Quick Start

1. **Double-click** `RosterGenerator.exe` to run the program
2. Follow the on-screen prompts to select Excel file and day
3. Your roster will be generated automatically!

## Requirements

- Windows 10 or later
- No additional software installation required!

## How to Use

1. **Place your Excel files** in the same folder as `RosterGenerator.exe`
2. **Run the program** by double-clicking the executable
3. **Select a file** from the list of available Excel files
4. **Choose a day** (M, T, W, Th, F, Sa, Su)
5. **Wait for generation** - the program will create your roster
6. **Export to Excel** if prompted (optional)

## Excel File Format

Your Excel file should have:
- A "Weekly" sheet with employee schedule data
- A "Team" sheet with department information

## Output

The program creates:
- Detailed roster schedules for each employee
- Task assignments (FR, GR, R, departments)
- Break assignments (40-min for shifts >6h, 10-min for all)
- Excel export files in "roster_output" folder

## Troubleshooting

**If the program doesn't start:**
- Make sure you're on Windows 10 or later
- Try running as administrator
- Check if antivirus software is blocking it

**If you get file errors:**
- Make sure Excel files are in the same folder
- Close Excel if the file is open
- Check file permissions

**If you get "command not found" errors:**
- The executable might be corrupted
- Try re-downloading the package

## Features

✅ Automatic file selection
✅ Day-based roster generation  
✅ Break management (40-min/10-min rules)
✅ Role-based task restrictions (ADM, SPV, ASM, SM)
✅ Excel export functionality
✅ Cross-platform compatibility

## Support

For technical issues, refer to the original Python source code.
"""
    
    with open(f"{package_name}/README.txt", "w", encoding="utf-8") as f:
        f.write(readme_content)
    print(f"   ✅ Created README.txt")
    
    # Create a Windows batch file launcher
    if sys.platform == "win32":
        batch_content = """@echo off
title Roster Generator
echo.
echo ========================================
echo    Roster Generator for Windows
echo ========================================
echo.
echo Starting program...
echo.
RosterGenerator.exe
echo.
echo Program finished.
echo Press any key to exit...
pause > nul
"""
        with open(f"{package_name}/Run_Roster_Generator.bat", "w") as f:
            f.write(batch_content)
        print(f"   ✅ Created Run_Roster_Generator.bat")
    
    # Copy sample Excel files if they exist
    excel_files = list(Path(".").glob("*.xlsx"))
    if excel_files:
        print(f"   📊 Found {len(excel_files)} Excel file(s), copying...")
        for excel_file in excel_files:
            if not excel_file.name.startswith("~$"):  # Skip temp files
                shutil.copy2(excel_file, f"{package_name}/{excel_file.name}")
                print(f"      ✅ Copied {excel_file.name}")
    
    print(f"\n🎉 Distribution package created!")
    print(f"📁 Package location: {os.path.abspath(package_name)}")
    print(f"📦 Contents:")
    for item in sorted(os.listdir(package_name)):
        item_path = os.path.join(package_name, item)
        if os.path.isfile(item_path):
            size = os.path.getsize(item_path) / 1024  # KB
            print(f"   📄 {item} ({size:.1f} KB)")
        else:
            print(f"   📁 {item}/")
    
    print(f"\n🚀 To distribute:")
    print(f"   1. Zip the '{package_name}' folder")
    print(f"   2. Copy to any Windows computer")
    print(f"   3. Extract and run {exe_name}")
    
    return True

if __name__ == "__main__":
    print("Starting Windows executable build process...")
    print()
    
    # Build the executable
    success = build_windows_executable()
    
    if success:
        # Create distribution package
        create_distribution_package()
        
        print(f"\n✅ All done! Your Windows executable is ready.")
        print(f"💡 Tip: Test it on your Windows machine before distributing.")
    else:
        print(f"\n❌ Build failed. Check the error messages above.")
        sys.exit(1)
    
    input("\nPress Enter to exit...")
