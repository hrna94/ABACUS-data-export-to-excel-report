# --- AUTOMATICKÁ KONTROLA A INSTALACE MODULU ---
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
        exit
    }
}
Write-Host "Module '$requiredModule' is available. Continuing with the script." -ForegroundColor Green

# --- HLAVNÍ FUNKCE SKRIPTU ---
# Tato funkce obsahuje veškerou logiku pro zpracování dat a export.
function Get-ParkingReport {
    param(
        [string]$eventFile,
        [string]$deviceFile,
        [string]$outputFile,
        [datetime]$startDate,
        [datetime]$endDate
    )
    # --- CORE SCRIPT LOGIC STARTS HERE ---
    Write-Host "--- Starting script ---"
    
    if (-not (Test-Path $eventFile)) {
        Write-Host "ERROR: Input file not found at path: $eventFile" -ForegroundColor Red
        return
    }
    Write-Host "Input file found: $eventFile" -ForegroundColor Green

    # --- PROCESSING ---
    $ticketTypes = @{
        1="Short term parking-Ticket"; 2="Advanced sale-Ticket"; 3="Congress-Ticket";
        4="Value card"; 5="Season parker card"; 6="Lost ticket"; 8="Hotel ticket";
        12="One use Ticket"; 13="Replacement Ticket"
    }
    $openEntries = @{}
    $allTrips = [System.Collections.Generic.List[object]]::new()

    # Load and process device data from INV file
    Write-Host "--- Loading device data from $deviceFile ---"
    $deviceLookup = @{}
    Get-Content $deviceFile | Where-Object { $_.StartsWith("P;") } | ForEach-Object {
        $fields = $_.Split(';')
        if ($fields.Count -ge 5 -and -not [string]::IsNullOrEmpty($fields[4])) {
            $deviceNumber = $fields[4]
            $carParkName = $fields[2]
            $deviceName = $fields[5]

            $deviceLookup[$deviceNumber] = [PSCustomObject]@{
                CarParkName = $carParkName
                DeviceName = $deviceName
            }
        }
    }
    Write-Host "Found $($deviceLookup.Count) devices to map."

    # Process ALL event data from EVT file without initial filtering
    Write-Host "--- Loading all relevant event (186) data from $eventFile ---"
    $events = Get-Content $eventFile | Where-Object { $_.StartsWith("P;") } | ForEach-Object {
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
    Write-Host "Found a total of $($events.Count) relevant events (type 186)."
    
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
                'Device/Car Park' = $deviceName;
                'Total Entries' = [int]$entries;
                'Total Exits' = [int]$exits
            }
        }
        
        # Calculate total entries and exits across all devices
        $totalEntries = ($joinedSummary | Measure-Object -Sum 'Total Entries').Sum
        $totalExits = ($joinedSummary | Measure-Object -Sum 'Total Exits').Sum

        # Add the totals row to the summary table
        $joinedSummary += [PSCustomObject]@{
            'Device/Car Park' = "SUMA celkem";
            'Total Entries' = [int]$totalEntries;
            'Total Exits' = [int]$totalExits
        }

        # Export the summary table to the first sheet
        $joinedSummary | Export-Excel -Path $outputFile -WorksheetName "Summary" -AutoSize -TableName "Summary" -TableStyle None

        # Export main table to the second sheet (using -Append)
        $finalReport = $completedTrips | Sort-Object 'Entry Time', 'ISO-Card number'
        $finalReport | Export-Excel -Path $outputFile -WorksheetName "Parking Report" -Append -AutoSize -TableName "ParkingData" -TableStyle None
        
        Write-Host "Done! Report successfully generated to: $outputFile" -ForegroundColor Green
    } else {
        Write-Host "WARNING: No records found for export. Output file was not created." -ForegroundColor Yellow
    }
    Write-Host "--- Script finished ---"
}

# --- GUI SECTION OF THE SCRIPT ---
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
$dateTimePickerStartDate.Value = (Get-Date).AddDays(-30) # Default value
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
$dateTimePickerEndDate.Value = (Get-Date) # Default value
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
    $saveFileDialog.Filter = "Excel Report (*.xlsx)|*.xlsx"
    $saveFileDialog.Title = "Save the generated report"
    $saveFileDialog.FileName = "report_parking.xlsx"
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
    $startDate = $dateTimePickerStartDate.Value # Get the DateTime value directly
    $endDate = $dateTimePickerEndDate.Value   # Get the DateTime value directly

    # Ensure the end date includes the entire day
    # We take the date part, add one day, and subtract one second.
    # This ensures it's 23:59:59 on the selected end date, regardless of the time in the GUI.
    $endDate = $endDate.Date.AddDays(1).AddSeconds(-1)
    
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
        $eventFile = Get-ChildItem -Path $inputFolder -Filter "*_EVT_*.txt" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Select-Object -ExpandProperty FullName
        $deviceFile = Get-ChildItem -Path $inputFolder -Filter "*_INV_*.txt" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Select-Object -ExpandProperty FullName

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

        Get-ParkingReport -eventFile $eventFile -deviceFile $deviceFile -outputFile $outputFile -startDate $startDate -endDate $endDate
        [System.Windows.Forms.MessageBox]::Show("Report was successfully generated!", "Success", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred during report generation: `n$($_.Exception.Message)", "Error", "OK", "Error")
    }
    finally {
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