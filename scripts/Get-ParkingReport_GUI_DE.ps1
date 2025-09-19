# --- LOAD VBA INJECTION FUNCTION ---
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbaFunctionPath = Join-Path $scriptPath "Add-VBAToExcel.ps1"
if (Test-Path $vbaFunctionPath) {
    . $vbaFunctionPath
    Write-Host "VBA injection function loaded." -ForegroundColor Green
} else {
    Write-Host "Warning: VBA injection function not found. Excel will be exported without VBA analysis tool." -ForegroundColor Yellow
}

# --- AUTOMATIC MODULE CHECK & INSTALL  ---
Write-Host "Checking for required module 'ImportExcel'..." -ForegroundColor Yellow

$requiredModule = "ImportExcel"
$moduleIsInstalled = (Get-Module -ListAvailable -Name $requiredModule) -ne $null

if (-not $moduleIsInstalled) {
    Write-Host "Module '$requiredModule' not found. Attempting to install..." -ForegroundColor Yellow

    try {
        # The installation will require administrator rights.
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber
        Write-Host "Module '$requiredModule' was successfully installed." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Could not install module '$requiredModule'." -ForegroundColor Red
        Write-Host "Please run PowerShell as an administrator and try again." -ForegroundColor Red
        Write-Host "The command you need is: Install-Module -Name ImportExcel" -ForegroundColor Red
        # Exit the script gracefully if installation fails
        Start-Sleep -Seconds 5
    }
}
Write-Host "Module '$requiredModule' is available. Continuing with the script." -ForegroundColor Green

# --- MAIN SCRIPT FUNCTION ---
# This function contain the logic to process data and export.
function Get-ParkingReport {
    param(
        [string]$inputFolder,
        [string]$outputFile,
        [datetime]$startDate,
        [datetime]$endDate
    )
    # --- CORE SCRIPT LOGIC STARTS HERE ---
    Write-Host "--- Starting script ---"

    if (-not (Test-Path $inputFolder)) {
        Write-Host "ERROR: Input folder not found at path: $inputFolder" -ForegroundColor Red
        return
    }
    Write-Host "Input folder found: $inputFolder" -ForegroundColor Green

    # --- PROCESSING ---
    $ticketTypes = @{
        1="Short term parking-Ticket"; 2="Advanced sale-Ticket"; 3="Congress-Ticket";
        4="Value card"; 5="Season parker card"; 6="Lost ticket"; 8="Hotel ticket";
        12="One use Ticket"; 13="Replacement Ticket"
    }
    $openEntries = @{}
    $allTrips = [System.Collections.Generic.List[object]]::new()

    # Load and process device data from ALL INV files
    Write-Host "--- Processing INV files... ---"
    $deviceFiles = Get-ChildItem -Path $inputFolder -Filter "*_INV_*.txt" | Sort-Object LastWriteTime
    $deviceLookup = @{}

    foreach ($deviceFile in $deviceFiles) {
        Get-Content $deviceFile.FullName | Where-Object { $_.StartsWith("P;") } | ForEach-Object {
            $fields = $_.Split(';')
            if ($fields.Count -ge 5 -and -not [string]::IsNullOrEmpty($fields[4])) {
                $deviceNumber = $fields[4]
                $carParkName = $fields[2]
                $deviceName = $fields[5]

                # Newer files override older entries (sorted by LastWriteTime)
                $deviceLookup[$deviceNumber] = [PSCustomObject]@{
                    CarParkName = $carParkName
                    DeviceName = $deviceName
                }
            }
        }
    }
    Write-Host "Found $($deviceFiles.Count) INV files with total of $($deviceLookup.Count) devices to map."

    # Process ALL event data from ALL EVT files without initial filtering
    Write-Host "--- Processing EVT files... ---"
    $eventFiles = Get-ChildItem -Path $inputFolder -Filter "*_EVT_*.txt" | Sort-Object LastWriteTime
    $events = @()

    foreach ($eventFile in $eventFiles) {
        $fileEvents = Get-Content $eventFile.FullName | Where-Object { $_.StartsWith("P;") } | ForEach-Object {
            $fields = $_.Split(';')
            if ($fields.Count -ge 14 -and $fields[6] -eq '186') {
                try {
                    $eventTime = [datetime]::ParseExact($fields[5], 'dd.MM.yyyy HH:mm:ss', $null)
                    [PSCustomObject]@{
                        DeviceNumber  = $fields[3]
                        EventType     = $fields[7]
                        EventTime     = $eventTime
                        CardNumber    = $fields[9]
                        TicketTypeRaw = $fields[13]
                    }
                } catch {
                    Write-Warning "Skipped a record due to date format error: $($fields[5])"
                }
            }
        }
        $events += $fileEvents
    }
    Write-Host "Found $($events.Count) events in $($eventFiles.Count) EVT files (type 186)."
    
    if ($events.Count -eq 0) {
        Write-Host "No relevant events found. Script is ending."
        return
    }

    # Process entries and exits and link them to device data
    $events | Sort-Object EventTime | ForEach-Object {
        $cardNumber = $_.CardNumber
        $eventType = $_.EventType
        $eventTime = $_.EventTime
        $deviceNumber = $_.DeviceNumber
        
        $deviceInfo = $deviceLookup[$deviceNumber]
        $carParkName = if ($deviceInfo) { $deviceInfo.CarParkName } else { "N/A" }
        $deviceName = if ($deviceInfo) { $deviceInfo.DeviceName } else { "N/A" }

        if ($eventType -eq '1') { # Entry
            if ($openEntries.ContainsKey($cardNumber)) {
                $prevEntry = $openEntries[$cardNumber]
                $prevDeviceInfo = $deviceLookup[$prevEntry.DeviceNumber]
                $prevCarParkName = if ($prevDeviceInfo) { $prevDeviceInfo.CarParkName } else { "N/A" }
                $prevDeviceName = if ($prevDeviceInfo) { $prevDeviceInfo.DeviceName } else { "N/A" }

                $allTrips.Add([PSCustomObject]@{
                    'ISO-Card number' = $prevEntry.CardNumber;
                    'Entry Time' = $prevEntry.EventTime;
                    'ENT number' = $prevEntry.DeviceNumber;
                    'CarPark Name (Entry)' = $prevCarParkName;
                    'Device Name (Entry)' = $prevDeviceName;
                    'Exit Time' = $null;
                    'EXT number' = $null;
                    'CarPark Name (Exit)' = $null;
                    'Device Name (Exit)' = $null;
                    'Ticket type' = if ($ticketTypes.ContainsKey([int]$prevEntry.TicketTypeRaw)) { $ticketTypes[[int]$prevEntry.TicketTypeRaw] } else { $prevEntry.TicketTypeRaw };
                    'Duration (minutes)' = $null
                })
            }
            $openEntries[$cardNumber] = $_
        } elseif ($eventType -eq '2') { # Exit
            if ($openEntries.ContainsKey($cardNumber)) {
                $entryEvent = $openEntries[$cardNumber]
                $entryDeviceInfo = $deviceLookup[$entryEvent.DeviceNumber]
                $entryCarParkName = if ($entryDeviceInfo) { $entryDeviceInfo.CarParkName } else { "N/A" }
                $entryDeviceName = if ($entryDeviceInfo) { $entryDeviceInfo.DeviceName } else { "N/A" }

                $duration = [Math]::Round((New-TimeSpan -Start $entryEvent.EventTime -End $eventTime).TotalMinutes)
                $allTrips.Add([PSCustomObject]@{
                    'ISO-Card number' = $cardNumber;
                    'Entry Time' = $entryEvent.EventTime;
                    'ENT number' = $entryEvent.DeviceNumber;
                    'CarPark Name (Entry)' = $entryCarParkName;
                    'Device Name (Entry)' = $entryDeviceName;
                    'Exit Time' = $eventTime;
                    'EXT number' = $_.DeviceNumber;
                    'CarPark Name (Exit)' = $carParkName;
                    'Device Name (Exit)' = $deviceName;
                    'Ticket type' = if ($ticketTypes.ContainsKey([int]$entryEvent.TicketTypeRaw)) { $ticketTypes[[int]$entryEvent.TicketTypeRaw] } else { $entryEvent.TicketTypeRaw };
                    'Duration (minutes)' = $duration
                })
                $openEntries.Remove($cardNumber)
            }
        }
    }

    # Add remaining open entries
    $openEntries.Values | ForEach-Object {
        $entryEvent = $_
        $entryDeviceInfo = $deviceLookup[$entryEvent.DeviceNumber]
        $entryCarParkName = if ($entryDeviceInfo) { $entryDeviceInfo.CarParkName } else { "N/A" }
        $entryDeviceName = if ($entryDeviceInfo) { $entryDeviceInfo.DeviceName } else { "N/A" }
        
        $allTrips.Add([PSCustomObject]@{
            'ISO-Card number' = $entryEvent.CardNumber;
            'Entry Time' = $entryEvent.EventTime;
            'ENT number' = $entryEvent.DeviceNumber;
            'CarPark Name (Entry)' = $entryCarParkName;
            'Device Name (Entry)' = $entryDeviceName;
            'Exit Time' = $null;
            'EXT number' = $null;
            'CarPark Name (Exit)' = $null;
            'Device Name (Exit)' = $null;
            'Ticket type' = if ($ticketTypes.ContainsKey([int]$entryEvent.TicketTypeRaw)) { $ticketTypes[[int]$entryEvent.TicketTypeRaw] } else { $entryEvent.TicketTypeRaw };
            'Duration (minutes)' = $null
        })
    }

    # --- FILTERING FOR FINAL REPORT ---
    $completedTrips = $allTrips | Where-Object { $_.'Entry Time' -ge $startDate -and $_.'Entry Time' -le $endDate }

    # --- EXPORT TO XLSX ---
    Write-Host "Found $($completedTrips.Count) complete records to export."
    if ($completedTrips.Count -gt 0) {
        # --- Generate and export summary table ---
        Write-Host "Generating summary table..." -ForegroundColor Yellow
        
        # Filter completed records for summary
        $summary = $allTrips | Where-Object { $_.'Entry Time' -ge $startDate -and $_.'Entry Time' -le $endDate }

        # Group entry events and count them
        $entrySummary = $summary | Where-Object { $_.'Entry Time' -ne $null } | Group-Object 'CarPark Name (Entry)', 'Device Name (Entry)' | Select-Object Name, @{N='Entries'; E={$_.Count}}

        # Group exit events and count them
        $exitSummary = $summary | Where-Object { $_.'Exit Time' -ne $null } | Group-Object 'CarPark Name (Exit)', 'Device Name (Exit)' | Select-Object Name, @{N='Exits'; E={$_.Count}}

        # Combine all devices and create a summary table
        $joinedSummary = @()
        $allDevices = ($entrySummary.Name + $exitSummary.Name | Select-Object -Unique).Trim()
        
        $allDevices | ForEach-Object {
            $deviceName = $_
            $entries = ($entrySummary | Where-Object Name -eq $deviceName).Entries
            $exits = ($exitSummary | Where-Object Name -eq $deviceName).Exits
            
            $joinedSummary += [PSCustomObject]@{
                'Parkhaus/Gerät' = $deviceName;
                'Gesamt Einfahrten' = [int]$entries;
                'Gesamt Ausfahrten' = [int]$exits
            }
        }
        
        # Calculate total entries and exits across all devices
        $totalEntries = ($joinedSummary | Measure-Object -Sum 'Gesamt Einfahrten').Sum
        $totalExits = ($joinedSummary | Measure-Object -Sum 'Gesamt Ausfahrten').Sum

        # Add the totals row to the summary table
        $joinedSummary += [PSCustomObject]@{
            'Parkhaus/Gerät' = "GESAMT";
            'Gesamt Einfahrten' = [int]$totalEntries;
            'Gesamt Ausfahrten' = [int]$totalExits
        }

        # Always export as .xlsx first for VBA injection
        $tempXlsxFile = $outputFile -replace '\.xlsm$', '.xlsx'

        # Export the summary table to the first sheet
        $joinedSummary | Export-Excel -Path $tempXlsxFile -WorksheetName "Zusammenfassung" -AutoSize -TableName "Summary" -TableStyle None

        # Export main table to the second sheet (using -Append)
        $finalReport = $completedTrips | Sort-Object 'Entry Time', 'ISO-Card number'
        $finalReport | Export-Excel -Path $tempXlsxFile -WorksheetName "Parking Report" -Append -AutoSize -TableName "ParkingData" -TableStyle None

        # Add VBA analysis tool if function is available
        if (Get-Command "Add-VBAToExcel" -ErrorAction SilentlyContinue) {
            $vbaFilePath = Join-Path $scriptPath "ParkingAnalysis_VBA_DE.bas"
            if (Test-Path $vbaFilePath) {
                try {
                    $finalOutputFile = Add-VBAToExcel -ExcelFilePath $tempXlsxFile -VBAFilePath $vbaFilePath
                    Write-Host "Done! Report with VBA analysis tool generated to: $finalOutputFile" -ForegroundColor Green
                    # Show final success message
                    [System.Windows.Forms.MessageBox]::Show("Export completed successfully!`n`nFile saved to:`n$finalOutputFile", "Export Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                } catch {
                    Write-Host "Error during VBA integration: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "Export completed but VBA integration failed. File saved to: $tempXlsxFile" -ForegroundColor Yellow
                    # Show warning message
                    [System.Windows.Forms.MessageBox]::Show("Export completed but VBA integration failed.`n`nError: $($_.Exception.Message)`n`nFile saved to: $tempXlsxFile", "Export Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            } else {
                Write-Host "Warning: VBA file not found. Report generated without analysis tool: $tempXlsxFile" -ForegroundColor Yellow
                # Show warning message
                [System.Windows.Forms.MessageBox]::Show("Export completed successfully but VBA analysis tool was not added.`n`nFile saved to:`n$tempXlsxFile", "Export Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        } else {
            Write-Host "Done! Report successfully generated to: $tempXlsxFile" -ForegroundColor Green
            # Show success message
            [System.Windows.Forms.MessageBox]::Show("Export completed successfully!`n`nFile saved to:`n$tempXlsxFile", "Export Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } else {
        Write-Host "WARNING: No records found for export. Output file was not created." -ForegroundColor Yellow
        # Show error message as final action
        [System.Windows.Forms.MessageBox]::Show("Export failed!`n`nNo records found for export. Please check your data files and date filters.", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    Write-Host "--- Script finished ---"
}

# --- GUI SECTION OF THE SKRIPTU ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main window
$form = New-Object System.Windows.Forms.Form
$form.Text = "Parking Report Generator"
$form.Size = New-Object System.Drawing.Size(400, 300)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# Input fields for start and end dates
$labelStartDate = New-Object System.Windows.Forms.Label
$labelStartDate.Text = "Start Date:"
$labelStartDate.Location = New-Object System.Drawing.Point(10, 10)
$labelStartDate.AutoSize = $true
$form.Controls.Add($labelStartDate)

# Date & Time picker for start date
$dateTimePickerStartDate = New-Object System.Windows.Forms.DateTimePicker
$dateTimePickerStartDate.Location = New-Object System.Drawing.Point(10, 30)
$dateTimePickerStartDate.Size = New-Object System.Drawing.Size(350, 20)
$dateTimePickerStartDate.Format = "Custom"
$dateTimePickerStartDate.CustomFormat = "dd.MM.yyyy HH:mm:ss"
$dateTimePickerStartDate.ShowUpDown = $false
$dateTimePickerStartDate.ShowCheckBox = $true
$dateTimePickerStartDate.Checked = $false
$form.Controls.Add($dateTimePickerStartDate)

$labelEndDate = New-Object System.Windows.Forms.Label
$labelEndDate.Text = "End Date:"
$labelEndDate.Location = New-Object System.Drawing.Point(10, 60)
$labelEndDate.AutoSize = $true
$form.Controls.Add($labelEndDate)

# Date & Time picker for end date
$dateTimePickerEndDate = New-Object System.Windows.Forms.DateTimePicker
$dateTimePickerEndDate.Location = New-Object System.Drawing.Point(10, 80)
$dateTimePickerEndDate.Size = New-Object System.Drawing.Size(350, 20)
$dateTimePickerEndDate.Format = "Custom"
$dateTimePickerEndDate.CustomFormat = "dd.MM.yyyy HH:mm:ss"
$dateTimePickerEndDate.ShowUpDown = $false
$dateTimePickerEndDate.ShowCheckBox = $true
$dateTimePickerEndDate.Checked = $false
$form.Controls.Add($dateTimePickerEndDate)

# Button and field for input folder selection
$labelInputFolder = New-Object System.Windows.Forms.Label
$labelInputFolder.Text = "Input Folder:"
$labelInputFolder.Location = New-Object System.Drawing.Point(10, 110)
$labelInputFolder.AutoSize = $true
$form.Controls.Add($labelInputFolder)

$textBoxInputFolder = New-Object System.Windows.Forms.TextBox
$textBoxInputFolder.Location = New-Object System.Drawing.Point(10, 130)
$textBoxInputFolder.Size = New-Object System.Drawing.Size(260, 20)
$textBoxInputFolder.Enabled = $false
$form.Controls.Add($textBoxInputFolder)

$buttonInputFolder = New-Object System.Windows.Forms.Button
$buttonInputFolder.Text = "Select Folder"
$buttonInputFolder.Location = New-Object System.Drawing.Point(280, 128)
$buttonInputFolder.Size = New-Object System.Drawing.Size(80, 25)
$buttonInputFolder.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $textBoxInputFolder.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($buttonInputFolder)

# Button and field for output file location
$labelOutputFile = New-Object System.Windows.Forms.Label
$labelOutputFile.Text = "Save As:"
$labelOutputFile.Location = New-Object System.Drawing.Point(10, 160)
$labelOutputFile.AutoSize = $true
$form.Controls.Add($labelOutputFile)

$textBoxOutputFile = New-Object System.Windows.Forms.TextBox
$textBoxOutputFile.Location = New-Object System.Drawing.Point(10, 180)
$textBoxOutputFile.Size = New-Object System.Drawing.Size(260, 20)
$textBoxOutputFile.Enabled = $false
$form.Controls.Add($textBoxOutputFile)

$buttonOutputFile = New-Object System.Windows.Forms.Button
$buttonOutputFile.Text = "Select File"
$buttonOutputFile.Location = New-Object System.Drawing.Point(280, 178)
$buttonOutputFile.Size = New-Object System.Drawing.Size(80, 25)
$buttonOutputFile.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Excel Report with VBA (*.xlsm)|*.xlsm|Excel Report (*.xlsx)|*.xlsx"
    $saveFileDialog.Title = "Save the generated report"
    $saveFileDialog.FileName = "report_parking.xlsm"
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $textBoxOutputFile.Text = $saveFileDialog.FileName
    }
})
$form.Controls.Add($buttonOutputFile)

# Run button
$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.Text = "Generate Report"
$buttonRun.Location = New-Object System.Drawing.Point(10, 220)
$buttonRun.Size = New-Object System.Drawing.Size(350, 40)
$buttonRun.Add_Click({
    # Get values from GUI
    $inputFolder = $textBoxInputFolder.Text
    $outputFile = $textBoxOutputFile.Text
    $startDate = $null
    $endDate = $null

    # Check if a start date is selected; if not, use the minimum possible date
    if ($dateTimePickerStartDate.Checked) {
        $startDate = $dateTimePickerStartDate.Value
    } else {
        $startDate = [datetime]::MinValue
    }

    # Check if an end date is selected; if not, use the maximum possible date
    if ($dateTimePickerEndDate.Checked) {
        # Ensure the end date includes the entire day
        $endDate = $dateTimePickerEndDate.Value.Date.AddDays(1).AddSeconds(-1)
    } else {
        $endDate = [datetime]::MaxValue
    }
    
    # Validate inputs
    if (-not $inputFolder) {
        [System.Windows.Forms.MessageBox]::Show("Please select an input folder.", "Error", "OK", "Error")
        return
    }
    if (-not $outputFile) {
        [System.Windows.Forms.MessageBox]::Show("Please select where to save the output file.", "Error", "OK", "Error")
        return
    }

    $form.Enabled = $false
    try {
        Write-Host "Script has been started..."

        # Check for EVT and INV files
        $eventFiles = Get-ChildItem -Path $inputFolder -Filter "*_EVT_*.txt"
        $deviceFiles = Get-ChildItem -Path $inputFolder -Filter "*_INV_*.txt"

        if ($eventFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No EVT files found in the selected folder.", "Error", "OK", "Error")
            return
        }

        if ($deviceFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No INV files found in the selected folder.", "Error", "OK", "Error")
            return
        }

        Write-Host "Found $($eventFiles.Count) EVT files and $($deviceFiles.Count) INV files"

        # Check if the output file is already open
        if (Test-Path $outputFile) {
            try {
                $fileHandle = [System.IO.File]::OpenWrite($outputFile)
                $fileHandle.Dispose()
            } catch {
                # This is the specific error for a locked file
                [System.Windows.Forms.MessageBox]::Show("The output file is currently open in Excel. Please close it before running the report.", "Error: File Locked", "OK", "Error")
                return
            }
        }

        Get-ParkingReport -inputFolder $inputFolder -outputFile $outputFile -startDate $startDate -endDate $endDate
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred during report generation: `n$($_.Exception.Message)", "Error", "OK", "Error")
    }
    finally {
        # Close the form only on success, otherwise keep it open to show the error
        if ($_.Exception -eq $null) {
            $form.Close()
        } else {
            $form.Enabled = $true
            Write-Host "GUI remains open due to an error."
        }
    }
})
$form.Controls.Add($buttonRun)

# Show the form
$form.ShowDialog() | Out-Null
