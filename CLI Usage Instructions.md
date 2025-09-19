### How to Use the CLI Script

This version of the script is designed to be run directly from the **PowerShell** console and does not have a graphical interface - To implement different way. All necessary information is provided as **command-line parameters** 

#### 1. File Preparation

Ensure the following files are in the same folder:

* `Get-ParkingReport_CLI.ps1`
* **All** your `.txt` data files (e.g., `xx_EVT_xxxx_xxxx.txt` and `xx_INV_xxxx_xxxx.txt`).

**Important:** The script will automatically process **all** EVT and INV files found in the specified folder, combining data from multiple days/periods into a single report.

#### 2. Running from PowerShell

Open PowerShell and navigate to the folder where you have saved the files. Then, run the script with the mandatory parameters **`-InputFolder`** (the path to the folder with the data) and **`-OutputFile`** (the path where the report will be saved).

#### Example 1: Processing Data for the Entire Period

If you don't specify a date range, the script will automatically process all data in the files.

```powershell
.\Get-ParkingReport.ps1 -InputFolder "C:\Users\username\data" -OutputFile "C:\Users\username\Desktop\parking_report.xlsx"
```

#### Example 2: Processing Data for a Specific Time Range

Use the -StartDate and -EndDate parameters to define a specific period. The date and time format must be 'dd.MM.yyyy HH:mm:ss'.

```PowerShell
.\Get-ParkingReport.ps1 -InputFolder "C:\Users\username\data" -OutputFile "C:\Users\username\Desktop\parking_report_january.xlsx" -StartDate "01.01.2024 00:00:00" -EndDate "31.01.2024 23:59:59"
```

The script will automatically find and process **all** EVT and INV files in the input folder, combining data from multiple periods. The progress and any potential errors will be shown directly in the PowerShell console.