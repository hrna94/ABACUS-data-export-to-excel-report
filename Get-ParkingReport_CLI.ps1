# --- AUTOMATIC MODULE CHECK AND INSTALLATION ---
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

# --- COMMAND LINE PARAMETERS ---
# Defines the parameters for running the script from the command line.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Path to the folder containing the raw EVT and INV files.")]
    [string]$InputFolder,

    [Parameter(Mandatory=$true, HelpMessage="Full path where the output Excel file will be saved.")]
    [string]$OutputFile,

    [Parameter(Mandatory=$false, HelpMessage="Start date and time for the report (e.g., '10.01.2023 08:00:00').")]
    [datetime]$StartDate = [datetime]::MinValue,

    [Parameter(Mandatory=$false, HelpMessage="End date and time for the report (e.g., '10.01.2023 18:00:00').")]
    [datetime]$EndDate = [datetime]::MaxValue
)

# --- MAIN SCRIPT FUNCTION ---
# This function contains all the logic for data processing and export.
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
            'Device/Car Park' = "TOTAL";
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

# --- SCRIPT EXECUTION AND FILE VALIDATION ---
# Process all EVT and INV files in the folder
Write-Host "Checking for data files in '$InputFolder'..." -ForegroundColor Yellow

$eventFiles = Get-ChildItem -Path $InputFolder -Filter "*_EVT_*.txt"
$deviceFiles = Get-ChildItem -Path $InputFolder -Filter "*_INV_*.txt"

if ($eventFiles.Count -eq 0) {
    Write-Host "ERROR: No EVT files found in '$InputFolder'" -ForegroundColor Red
    return
}

if ($deviceFiles.Count -eq 0) {
    Write-Host "ERROR: No INV files found in '$InputFolder'" -ForegroundColor Red
    return
}

Write-Host "Found $($eventFiles.Count) EVT files and $($deviceFiles.Count) INV files" -ForegroundColor Green

# Check if the output file is already open
if (Test-Path $OutputFile) {
    try {
        $fileHandle = [System.IO.File]::OpenWrite($OutputFile)
        $fileHandle.Dispose()
    } catch {
        # This is the specific error for a locked file
        Write-Host "ERROR: The output file '$OutputFile' is currently open. Please close it before running the report." -ForegroundColor Red
        return
    }
}

Get-ParkingReport -inputFolder $InputFolder -outputFile $OutputFile -startDate $StartDate -endDate $EndDate
