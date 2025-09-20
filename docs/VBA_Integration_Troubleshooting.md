# VBA Integration Troubleshooting

If you encounter the error "You cannot call a method on a null-valued expression" during VBA integration, this usually means Excel cannot access the VBA project. Here are the solutions:

## 0. Quick Setup Script (Recommended)

Run the automated setup script:
```powershell
.\scripts\Setup-ExcelVBAAccess.ps1
```

This script will:
- Check if Excel is installed
- Provide step-by-step manual setup instructions
- Optionally modify registry settings automatically
- Test VBA access configuration

## 1. Enable VBA Project Access in Excel

**Excel 2016/2019/365:**
1. Open Excel
2. Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. Select **Macro Settings**
4. Check **"Trust access to the VBA project object model"**
5. Click **OK** and restart Excel

**Excel 2013:**
1. Open Excel
2. Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. Select **Macro Settings**
4. Enable **"Trust access to the VBA project object model"**

## 2. Adjust Macro Security Settings

1. In Excel Trust Center → **Macro Settings**
2. Select **"Disable all macros with notification"** or **"Enable all macros"**
3. Restart Excel

## 3. PowerShell Execution Policy

If you get security warnings, run PowerShell as Administrator and execute:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 4. Unblock Downloaded Files

If the script was downloaded from the internet:

```powershell
Unblock-File -Path "scripts\Add-VBAToExcel.ps1"
Unblock-File -Path "scripts\*.ps1"
```

## 5. Alternative: Manual VBA Setup

If VBA integration still fails, the script will generate a regular XLSX file. You can manually:

1. Save the XLSX file as XLSM format
2. Open the VBA editor (Alt+F11)
3. Import the VBA module from `scripts\ParkingAnalysis_VBA.bas`

## 6. Check Excel Installation

Ensure Excel is properly installed and COM objects are registered:

```powershell
Get-CimInstance -ClassName Win32_Product | Where-Object Name -like "*Excel*"
```

## Error Messages

- **"Excel is not installed or COM interface is not available"** → Install Microsoft Excel
- **"VBA access may be disabled"** → Follow steps 1-2 above
- **"You cannot call a method on a null-valued expression"** → VBA project access is blocked

## Fallback Behavior

If VBA integration fails, the script will:
- Generate a standard XLSX report
- Show a warning message
- Continue with the export process
- The data will still be available for analysis