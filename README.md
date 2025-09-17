# PowerShell Parking Report Generator

This PowerShell script is for automatically generating reports from a ABACUS DataExport. The script processes raw Events (EVT) and a System Inventory List (INV) to create a clear and organized report in Excel (.xlsx) format.

## Key Features

**Data Mapping:** The script automatically pairs entry and exit events with corresponding car park names and device names.
**Time-Based Filtering:** Allows you to generate a report for a specific timeframe (based on the entry date and time).
**Complete Records:** Ensures that the report includes full entry and exit data, even if the exit occurred outside the specified timeframe.
**Excel Output:** Exports all processed data into a single Excel file for easy analysis.

## Prerequisites
- The script requires the ImportExcel module for PowerShell. If you don't have it, install it by running the following command in PowerShell with administrator privileges:
**Install-Module -Name ImportExcel -Scope CurrentUser**

## How to Use

### 1. File Setup

Ensure the following files are located in the same folder:
- `Get-ParkingReport.ps1` (this script)
- The events file (e.g., `01_EVT_xxxx_xxxx.txt`) **- name has to be defined in the script parameters**
- The device information file (e.g., `01_INV_xxxx_xxxx.txt`) **- name has to be defined in the script parameters**

### 2. Running the Script

Open PowerShell and execute the script from its location.

#### Example 1: Using the Default Timeframe
 **If you don't specify a date, the script will use the default values set within the file (01.01.2000-01.01.2045).**
Open powershell and execute **Get-ParkingReport.ps1**

#### Example 2: Specifying a Custom Timeframe
 **To generate a report for a specific period, use the -startDate and -endDate parameters.**
Open powershell and execute **Get-ParkingReport.ps1 -startDate "DD.MM.YYYY hh:mm:ss" -endDate "DD.MM.YYYY hh:mm:ss"**
The script will run and generate a file named report_parking.xlsx in the same directory.