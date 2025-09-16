function Get-ParkingReport {
    param(
        [string]$eventFile = "01_EVT_250801_250831.txt",
        [string]$deviceFile = "01_INV_250801_250831.txt",
        [string]$outputFile = "report_parkovani33.xlsx",
        [string]$startDate = (Get-Date "01.08.2025 00:00:00").ToString("dd.MM.yyyy HH:mm:ss"),
        [string]$endDate = (Get-Date "31.08.2025 23:59:59").ToString("dd.MM.yyyy HH:mm:ss")
    )

    # --- SCRIPT BODY STARTS HERE ---
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
            [PSCustomObject]@{
                DeviceNumber  = $fields[3]
                EventType     = $fields[7]
                EventTime     = [datetime]::ParseExact($fields[5], 'dd.MM.yyyy HH:mm:ss', $null)
                CardNumber    = $fields[9]
                TicketTypeRaw = $fields[13]
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
    $startDT = [datetime]::ParseExact($startDate, 'dd.MM.yyyy HH:mm:ss', $null)
    $endDT = [datetime]::ParseExact($endDate, 'dd.MM.yyyy HH:mm:ss', $null)
    $completedTrips = $allTrips | Where-Object { $_.'Entry Time' -ge $startDT -and $_.'Entry Time' -le $endDT }

    # --- EXPORT TO XLSX ---
    Write-Host "Found $($completedTrips.Count) complete records to export."
    if ($completedTrips.Count -gt 0) {
        $finalReport = $completedTrips | Sort-Object 'Entry Time', 'ISO-Card number'
        $finalReport | Export-Excel -Path $outputFile -AutoSize -TableName "ParkingData" -WorksheetName "Parking Report" -TableStyle None
        Write-Host "Done! Report successfully generated to: $outputFile" -ForegroundColor Green
    } else {
        Write-Host "WARNING: No records found for export. Output file was not created." -ForegroundColor Yellow
    }

    Write-Host "--- Script finished ---"
}

# --- HOW TO RUN THE SCRIPT ---
# Save the code as a .ps1 file and run it from PowerShell.
# Example 1: Use default timeframe (August 2025)
# Get-ParkingReport
#
# Example 2: Specify a custom timeframe
# Get-ParkingReport -startDate "01.08.2025 10:00:00" -endDate "02.08.2025 12:00:00"