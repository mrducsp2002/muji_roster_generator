#!/usr/bin/env python3
"""
Script to create a Windows executable package for the Roster Generator
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def create_windows_package():
    """Create a complete Windows package with executable and documentation"""
    
    print("🚀 Creating Windows Package for Roster Generator")
    print("=" * 50)
    
    # Clean up previous builds
    print("🧹 Cleaning up previous builds...")
    for dir_name in ["dist", "build"]:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"   Removed {dir_name}/")
    
    for spec_file in Path(".").glob("*.spec"):
        spec_file.unlink()
        print(f"   Removed {spec_file}")
    
    # Create the executable
    print("\n📦 Creating executable...")
    try:
        cmd = [
            "pyinstaller",
            "--onefile",                    # Single file executable
            "--name", "RosterGenerator",    # Name of the executable
            "--console",                    # Keep console window for input/output
            "--add-data", "*.xlsx:.",       # Include Excel files
            "--hidden-import", "pandas",    # Ensure pandas is included
            "--hidden-import", "openpyxl",  # Ensure openpyxl is included
            "--hidden-import", "xlrd",      # Additional Excel support
            "roster_claude.py"
        ]
        
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
    
    # Create distribution folder
    print("\n📁 Creating distribution package...")
    dist_folder = "RosterGenerator_Windows"
    if os.path.exists(dist_folder):
        shutil.rmtree(dist_folder)
    os.makedirs(dist_folder)
    
    # Copy executable
    exe_source = "dist/RosterGenerator.exe"
    exe_dest = f"{dist_folder}/RosterGenerator.exe"
    if os.path.exists(exe_source):
        shutil.copy2(exe_source, exe_dest)
        print(f"   ✅ Copied executable to {dist_folder}/")
    
    # Create README for Windows users
    readme_content = """# Roster Generator for Windows

## How to Use

1. **Double-click** `RosterGenerator.exe` to run the program
2. The program will open in a command window
3. Follow the on-screen prompts to:
   - Select an Excel file from the available files
   - Choose a day (M, T, W, Th, F, Sa, Su)
   - Generate your roster

## Requirements

- Windows 10 or later
- No additional software installation required!

## Excel File Format

The program works with Excel files that have:
- A "Weekly" sheet with employee schedule data
- A "Team" sheet with department information

## Troubleshooting

If you get an error message:
1. Make sure your Excel file is in the same folder as RosterGenerator.exe
2. Make sure the Excel file is not open in another program
3. Try running as administrator if you get permission errors

## Output

The program will:
- Generate a roster for the selected day
- Show employee schedules and task assignments
- Optionally export to Excel format
- Save output files in a "roster_output" folder

## Support

For issues or questions, check the original Python files for technical details.
"""
    
    with open(f"{dist_folder}/README.txt", "w") as f:
        f.write(readme_content)
    print(f"   ✅ Created README.txt")
    
    # Create a simple launcher script
    launcher_content = """@echo off
title Roster Generator
echo Starting Roster Generator...
echo.
RosterGenerator.exe
echo.
echo Press any key to exit...
pause > nul
"""
    
    with open(f"{dist_folder}/Run_Roster_Generator.bat", "w") as f:
        f.write(launcher_content)
    print(f"   ✅ Created Run_Roster_Generator.bat")
    
    # Copy any Excel files in the current directory
    excel_files = list(Path(".").glob("*.xlsx"))
    if excel_files:
        print(f"   📊 Found {len(excel_files)} Excel file(s), copying...")
        for excel_file in excel_files:
            shutil.copy2(exe_file, f"{dist_folder}/{excel_file.name}")
            print(f"      Copied {excel_file.name}")
    
    print(f"\n🎉 Windows package created successfully!")
    print(f"📁 Package location: {os.path.abspath(dist_folder)}")
    print(f"📦 Contents:")
    for item in os.listdir(dist_folder):
        print(f"   - {item}")
    
    print(f"\n🚀 To distribute:")
    print(f"   1. Zip the '{dist_folder}' folder")
    print(f"   2. Copy to any Windows computer")
    print(f"   3. Extract and run RosterGenerator.exe")
    
    return True

if __name__ == "__main__":
    success = create_windows_package()
    if not success:
        sys.exit(1)
    
    input("\nPress Enter to exit...")
