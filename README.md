# PowerShell Parking Report Generator

Advanced PowerShell solution for generating comprehensive parking reports from ABACUS DataExport files with integrated VBA analysis tools.

## Key Features

- **️ User-Friendly GUI:** Simple graphical interface - just click and generate
- ** VBA Analysis Tool:** Interactive Excel dashboard for advanced data analysis
- ** Automatic Processing:** Processes ALL EVT and INV files in folder automatically
- ** Flexible Date Filtering:** Optional time range filters or process all data
- ** Dual Output Format:** Summary overview + detailed parking events
- ** Auto Installation:** ImportExcel module installed automatically
- ** Smart Error Handling:** Clear error messages and robust file processing

## Quick Start

### For End Users (GUI)

1. **Double-click:** `Run-ParkingReport-GUI.bat`
2. **Select:** Your data folder (containing EVT and INV files)
3. **Choose:** Output location (.xlsm recommended for VBA features)
4. **Optional:** Set date filters or leave blank for all data
5. **Click:** "Generate Report"

### For Advanced Users (CLI)

```powershell
.\scripts\Get-ParkingReport_CLI.ps1 -InputFolder "C:\Data" -OutputFile "C:\Reports\report.xlsm"
```

## Project Structure

```
Parking Report Generator
├── Run-ParkingReport-GUI.bat     # Main launcher
├── README.md                     # This documentation
├── scripts/                      # PowerShell scripts
│├── Get-ParkingReport_GUI.ps1    # GUI version
 ├── Get-ParkingReport_CLI.ps1    # Command-line version
│├── Add-VBAToExcel.ps1           # VBA integration
│└── ParkingAnalysis_VBA.bas      # Excel analysis toolkit
└── docs/                         # Language variants (CZ/EN)
 └── Excel_Macros_Setup_Guide*.md # Macro setup (multiple languages)
```

## Output Features

### Excel Report (.xlsm)

- **Summary Sheet:** Overview with entry/exit counts per device
- **Parking Report Sheet:** Detailed event log with timestamps
- **VBA Analysis Tool:** Interactive dashboard for advanced filtering

### VBA Analysis Capabilities

- ** Time Period Filtering:** Custom date/time ranges
- **️ Analysis Intervals:** Hourly, daily, weekly, monthly grouping
- ** Device-Specific Filtering:** Entry/exit point analysis
- ** Traffic Flow Analysis:** Entry → Exit device patterns
- ** Built-in Help:** Integrated user documentation

## Security & Macros

When opening the generated `.xlsm` file, Excel will show a security warning:

```
SECURITY WARNING Macros have been disabled. [Enable Content]
```

**→ Click "Enable Content"** to activate the VBA analysis features.

### Permanent Macro Settings (Optional)
1. **File** → **Options** → **Trust Center**
2. **Trust Center Settings** → **Macro Settings**
3. Select: **"Disable all macros with notification"**

## CLI Usage (Advanced)

### Prerequisites
Open PowerShell and navigate to the project folder. If you encounter execution policy issues:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Parameters
- **-InputFolder** (required): Path to folder containing EVT and INV files
- **-OutputFile** (required): Path for output .xlsm file
- **-StartDate** (optional): Start date in format "dd.MM.yyyy HH:mm:ss"
- **-EndDate** (optional): End date in format "dd.MM.yyyy HH:mm:ss"

### Basic Export
```powershell
.\scripts\Get-ParkingReport_CLI.ps1 -InputFolder "C:\ParkingData" -OutputFile "C:\Reports\report.xlsm"
```

### Filtered Export (August 2025)
```powershell
.\scripts\Get-ParkingReport_CLI.ps1 -InputFolder "C:\Data" -OutputFile "C:\Reports\august.xlsm" -StartDate "01.08.2025 00:00:00" -EndDate "31.08.2025 23:59:59"
```

## System Requirements

- **Windows:** 10/11 or Windows Server 2016+
- **PowerShell:** 5.0+ (included in Windows 10/11)
- **Excel:** 2016+ (for .xlsm support and VBA features)
- **ImportExcel Module:** Installed automatically

### Installation & Setup
1. **Extract all files** to a folder on your computer
2. **Ensure all files** are in correct structure (see Project Structure above)
3. **Run once:** Double-click `Run-ParkingReport-GUI.bat`

## Data File Requirements

### Input Files (in data folder):
- **Event Files:** `*_EVT_*.txt` - parking events data
- **Inventory Files:** `*_INV_*.txt` - device mapping information

## VBA Analysis Tool Guide

### Running Analysis
1. Open the generated `.xlsm` file
2. Enable macros when prompted
3. Go to "Summary" worksheet
4. Use the analysis controls to:
   - Set time period filters
   - Choose analysis interval (hourly/daily/weekly/monthly)
   - Filter by specific entry/exit devices
   - Run Analysis

### Analysis Features
- **No Exit Recorded:** Highlighted in **bold** for incomplete trips
- **Device Combinations:** Shows traffic flow between entry/exit points
- **Time-based Grouping:** Flexible interval analysis
- **Total Counts:** Comprehensive reporting with summaries

## Troubleshooting

### Common Issues
**Macros not working:**
- Ensure you clicked "Enable Content"
- Restart Excel if needed

**GUI not responding:**
- Check if all script files are in `/scripts` folder
- Verify file permissions

**CLI execution errors:**
- Run PowerShell as administrator
- Check ExecutionPolicy settings

**No data in report:**
- Verify EVT and INV files exist in input folder
- Check date filters aren't too restrictive
- Ensure files contain valid data

### Error Messages
The tool provides clear English error messages for:
- **Success:** "Export completed successfully!"
-  **Warning:** "VBA integration failed" or "VBA file not found"
- **Error:** "No records found" or file access issues

## Advanced Macro Configuration

If you frequently work with .xlsm files and want to avoid the security warning each time:

### Option 1: Trust Center Settings
1. Open Excel → **File** → **Options**
2. Go to **Trust Center** → **Trust Center Settings**
3. Select **Macro Settings**
4. Choose: **"Disable all macros with notification"** (recommended)
5. Click **OK** to save

### Option 2: Trusted Locations
1. In Trust Center Settings, go to **Trusted Locations**
2. Click **Add new location**
3. Browse to the folder where you save reports
4. Check **"Subfolders of this location are also trusted"**
5. Click **OK**

**Note:** Files in trusted locations will automatically enable macros without warnings.

## About

**Latest Version:** Includes VBA integration for interactive analysis
**License:** For DESIGNA GMBH use
**Support:** Check documentation in `/docs` folder