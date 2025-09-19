# Function to add VBA code to Excel file and convert to XLSM
function Add-VBAToExcel {
    param(
        [string]$ExcelFilePath,
        [string]$VBAFilePath
    )

    Write-Host "Adding VBA analysis tool to Excel file..." -ForegroundColor Yellow

    try {
        # Verify input file exists and is .xlsx
        if (-not (Test-Path $ExcelFilePath)) {
            throw "Input Excel file not found: $ExcelFilePath"
        }

        if ($ExcelFilePath -notmatch '\.xlsx$') {
            throw "Input file must be .xlsx format, got: $ExcelFilePath"
        }

        Write-Host "Opening Excel file: $ExcelFilePath" -ForegroundColor Gray

        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        # Open the Excel file
        $workbook = $excel.Workbooks.Open($ExcelFilePath)

        # Read VBA code from file
        if (Test-Path $VBAFilePath) {
            $vbaCode = Get-Content $VBAFilePath -Raw -Encoding UTF8
        } else {
            throw "VBA file not found: $VBAFilePath"
        }

        # Add VBA module
        $vbaModule = $workbook.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        $vbaModule.Name = "ParkingAnalysisModule"
        $vbaModule.CodeModule.AddFromString($vbaCode)

        # Create XLSM file path
        $xlsmFilePath = $ExcelFilePath -replace '\.xlsx$', '.xlsm'

        # If input file already has .xlsm extension, use it directly
        if ($ExcelFilePath -match '\.xlsm$') {
            $xlsmFilePath = $ExcelFilePath
        }

        # Save as XLSM (Excel Macro-Enabled Workbook)
        $workbook.SaveAs($xlsmFilePath, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled

        # Run the setup function to create UI (without showing message boxes)
        try {
            # Temporarily disable alerts to prevent popup dialogs
            $excel.DisplayAlerts = $false
            $excel.Application.Run("SetupSummaryUI")
            Write-Host "VBA UI setup completed successfully!" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not automatically run UI setup. Please run 'SetupSummaryUI' manually in Excel." -ForegroundColor Yellow
        }

        # Close and clean up
        $workbook.Close($true)
        $excel.Quit()

        # Verify XLSM file was created
        if (Test-Path $xlsmFilePath) {
            Write-Host "Excel file successfully converted to XLSM with VBA analysis tool: $xlsmFilePath" -ForegroundColor Green

            # Delete XLSX file immediately (no background job)
            if ($ExcelFilePath -ne $xlsmFilePath) {
                try {
                    Remove-Item $ExcelFilePath -Force -ErrorAction SilentlyContinue
                } catch { }
            }

            return $xlsmFilePath
        } else {
            throw "Failed to create XLSM file at: $xlsmFilePath"
        }

    } catch {
        Write-Host "Error adding VBA to Excel: $($_.Exception.Message)" -ForegroundColor Red

        # Cleanup on error
        if ($workbook) {
            try { $workbook.Close($false) } catch { }
        }
        if ($excel) {
            try { $excel.Quit() } catch { }
        }

        return $ExcelFilePath  # Return original file if VBA addition fails
    } finally {
        # Release COM objects immediately (no background processing)
        try {
            if ($vbaModule) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbaModule) | Out-Null }
            if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
            if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
        } catch { }
        # Skip garbage collection to avoid delay
    }
}