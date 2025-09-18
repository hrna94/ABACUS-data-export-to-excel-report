# PowerShell Parking Report Generator

This PowerShell script automatically generates clear and organized reports in Excel (.xlsx) format from raw data from an ABACUS DataExport. It processes raw Events (EVT) and a System Inventory List (INV) files.

## Key Features

* **GUI:** The script now features a user-friendly graphical interface, eliminating the need for command-line parameters.
* **Automatic Installation:** The script automatically checks for and installs the necessary `ImportExcel` module, simplifying the setup process for new users.
* **Flexible Date Filtering:** The filtering mechanism reliably generates reports for a specified date range, but now you can also leave the date/time fields unchecked to process **all data from the entire period** contained within the file.
* **Data Mapping:** The script automatically pairs entry and exit events with corresponding car park names and device names for a comprehensive report.
* **Advanced Excel Output:** The script now creates a single `.xlsx` file with two worksheets:
    * **Summary:** Provides a clear overview of total entries and exits per device within the selected date range.
    * **Parking Report:** Contains the original, detailed list of all entry and exit events.
* **Improved Error Handling:** The script now detects and provides a specific, clear error message if the output file is open.

## How to Use

### 1. File Preparation

Ensure the following files are in the same folder:

-   `Get-ParkingReport.ps1` (this script)
-   The events file (e.g., `01_EVT_xxxx_xxxx.txt`)
-   The device information file (e.g., `01_INV_xxxx_xxxx.txt`)

### 2. Running the Script

Double-click `Get-ParkingReport.ps1` or run it from a PowerShell console.

* **Initial Run:** The first time you run it, the script may prompt you to install the required `ImportExcel` module. Confirm the installation by pressing `Y` and `Enter`.
* **GUI Usage:** A graphical window will appear. Follow these steps:
    1.  Select the **input folder** where your `.txt` data files are located.
    2.  Voluntarily **check** and select the desired **start** and **end dates/times** for the report. If you leave the boxes unchecked, the script will process all available data.
    3.  Specify the **output file** location and name.
    4.  Click the **"Generate Report"** button to start the process.
* **Process Feedback:** You will see the progress of the export displayed in the PowerShell console window. Upon successful completion, a confirmation message will appear.