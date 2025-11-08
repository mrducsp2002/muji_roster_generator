# Windows Deployment Guide for Roster Generator

## 🚀 Quick Start (For End Users)

If you just want to **run** the roster generator on Windows:

1. **Download** the `RosterGenerator_Windows` folder
2. **Extract** it to your desired location
3. **Double-click** `RosterGenerator.exe`
4. **Follow** the on-screen prompts

**That's it!** No Python installation required.

---

## 🔧 Building Executable (For Developers)

If you need to **create** the Windows executable:

### Option 1: Automated Build (Recommended)

1. **Open Command Prompt** as Administrator
2. **Navigate** to your project folder
3. **Double-click** `build_windows.bat`
4. **Wait** for the build to complete
5. **Find** your executable in `dist/RosterGenerator.exe`

### Option 2: Manual Build

1. **Install PyInstaller**:
   ```cmd
   pip install pyinstaller
   ```

2. **Build the executable**:
   ```cmd
   python build_windows_exe.py
   ```

3. **Or use direct PyInstaller**:
   ```cmd
   pyinstaller --onefile --name "RosterGenerator" roster_claude.py
   ```

---

## 📦 Distribution Package Contents

After building, you'll have:

```
RosterGenerator_Windows/
├── RosterGenerator.exe          # Main executable
├── Run_Roster_Generator.bat     # Windows launcher
├── README.txt                   # User instructions
├── CHW October Week 44 Roster.xlsx  # Sample Excel file
└── modify_this.xlsx            # Another sample file
```

---

## 🖥️ System Requirements

### For Running the Executable:
- **Windows 10** or later (64-bit recommended)
- **No Python installation** required
- **~50-100 MB** disk space

### For Building the Executable:
- **Windows 10** or later
- **Python 3.8+** installed
- **pip** package manager
- **Internet connection** (for downloading PyInstaller)

---

## 🎯 Usage Instructions

### Step-by-Step Guide:

1. **Prepare Excel Files**
   - Place your Excel files in the same folder as `RosterGenerator.exe`
   - Ensure files have "Weekly" and "Team" sheets

2. **Run the Program**
   - Double-click `RosterGenerator.exe`
   - A command window will open

3. **Select File**
   - Choose from available Excel files
   - Press the number corresponding to your file

4. **Choose Day**
   - Enter day code: M, T, W, Th, F, Sa, Su
   - Press Enter

5. **Generate Roster**
   - Wait for processing
   - Review the generated roster
   - Choose to export to Excel if desired

6. **Output Files**
   - Generated rosters saved in `roster_output/` folder
   - Files named with day and timestamp

---

## 🔧 Troubleshooting

### Common Issues:

**"Windows protected your PC"**
- Click "More info" → "Run anyway"
- Or right-click → "Properties" → "Unblock"

**"Python not found" (when building)**
- Install Python from https://python.org
- Check "Add Python to PATH" during installation
- Restart Command Prompt

**"Permission denied"**
- Run Command Prompt as Administrator
- Check file permissions
- Disable antivirus temporarily

**"Excel file not found"**
- Ensure Excel files are in the same folder
- Close Excel if files are open
- Check file extensions (.xlsx)

**"Import error" (when building)**
- Install missing packages: `pip install pandas openpyxl`
- Update pip: `python -m pip install --upgrade pip`

---

## 📋 Features Included

✅ **File Selection**: Automatic detection of Excel files
✅ **Day Selection**: Support for all 7 days of the week
✅ **Break Management**: 
   - 40-minute breaks for shifts >6 hours
   - 10-minute breaks for all employees
✅ **Role Restrictions**:
   - ADM: No customer service tasks
   - SPV/ASM/SM: No FR/GR tasks (can do Register)
✅ **Excel Export**: Automatic roster export functionality
✅ **Department Mapping**: Correct department codes from Team sheet

---

## 🚀 Deployment Options

### Option 1: Direct Distribution
- Zip the `RosterGenerator_Windows` folder
- Send to users via email/cloud storage
- Users extract and run

### Option 2: Network Deployment
- Place on shared network drive
- Create desktop shortcuts
- Users run from network location

### Option 3: USB Distribution
- Copy folder to USB drive
- Users can run directly from USB
- No installation required

---

## 📞 Support

### For End Users:
- Check the `README.txt` in the distribution package
- Ensure Excel files are properly formatted
- Try running as administrator if issues occur

### For Developers:
- Check Python and pip installation
- Verify all dependencies are installed
- Review PyInstaller documentation for advanced options

---

## 🔄 Updates

To update the executable:
1. Make changes to Python source code
2. Re-run the build process
3. Distribute new executable to users
4. No reinstallation required on user machines

---

*This deployment package was created to make the Roster Generator accessible to Windows users without requiring Python knowledge or installation.*
