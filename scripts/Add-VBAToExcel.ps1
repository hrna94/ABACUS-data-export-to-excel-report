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
        try {
            $excel = New-Object -ComObject Excel.Application
            if ($excel -eq $null) {
                throw "Failed to create Excel COM object"
            }
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            Write-Host "Excel COM object created successfully" -ForegroundColor Green
        } catch {
            throw "Excel is not installed or COM interface is not available: $($_.Exception.Message)"
        }

        # Open the Excel file
        try {
            $workbook = $excel.Workbooks.Open($ExcelFilePath)
            if ($workbook -eq $null) {
                throw "Failed to open workbook"
            }
            Write-Host "Workbook opened successfully" -ForegroundColor Green
        } catch {
            throw "Failed to open Excel file: $($_.Exception.Message)"
        }

        # Read VBA code from file
        if (Test-Path $VBAFilePath) {
            $vbaCode = Get-Content $VBAFilePath -Raw -Encoding UTF8
        } else {
            throw "VBA file not found: $VBAFilePath"
        }

        # Add VBA module
        try {
            Write-Host "Accessing VBA project..." -ForegroundColor Gray
            $vbaProject = $workbook.VBProject
            if ($vbaProject -eq $null) {
                Write-Host "" -ForegroundColor Red
                Write-Host "ERROR: VBA project access is blocked!" -ForegroundColor Red
                Write-Host "This usually means Excel Trust Center settings are blocking VBA access." -ForegroundColor Yellow
                Write-Host "" -ForegroundColor Yellow
                Write-Host "SOLUTION:" -ForegroundColor Green
                Write-Host "1. Open Excel manually" -ForegroundColor White
                Write-Host "2. Go to File → Options → Trust Center → Trust Center Settings" -ForegroundColor White
                Write-Host "3. Select 'Macro Settings'" -ForegroundColor White
                Write-Host "4. Check 'Trust access to the VBA project object model'" -ForegroundColor White
                Write-Host "5. Restart Excel and try again" -ForegroundColor White
                Write-Host "" -ForegroundColor Yellow
                Write-Host "See README.md for detailed instructions." -ForegroundColor Cyan
                throw "VBA project access is blocked - Excel Trust Center settings need to be adjusted"
            }

            Write-Host "Adding VBA module..." -ForegroundColor Gray
            $vbaModule = $vbaProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            if ($vbaModule -eq $null) {
                throw "Failed to add VBA module"
            }
            $vbaModule.Name = "ParkingAnalysisModule"
            Write-Host "VBA module added successfully" -ForegroundColor Green
        } catch {
            if ($_.Exception.Message -like "*VBA project access is blocked*") {
                throw $_.Exception.Message
            } else {
                Write-Host "" -ForegroundColor Red
                Write-Host "VBA Integration Error Details:" -ForegroundColor Red
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor White
                Write-Host "" -ForegroundColor Yellow
                Write-Host "Common causes:" -ForegroundColor Yellow
                Write-Host "- VBA project object model access is disabled" -ForegroundColor White
                Write-Host "- Excel macro security settings are too restrictive" -ForegroundColor White
                Write-Host "- Excel COM interface issues" -ForegroundColor White
                Write-Host "" -ForegroundColor Cyan
                Write-Host "See README.md for solutions." -ForegroundColor Cyan
                throw "Failed to add VBA module: $($_.Exception.Message)"
            }
        }
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
        Write-Host "" -ForegroundColor Red
        Write-Host "VBA Integration Failed: $($_.Exception.Message)" -ForegroundColor Red

        if ($_.Exception.Message -like "*VBA project access is blocked*") {
            Write-Host "" -ForegroundColor Yellow
            Write-Host "The script will continue and generate a standard XLSX file." -ForegroundColor Yellow
            Write-Host "You can manually add VBA later using the troubleshooting guide." -ForegroundColor Yellow
        } else {
            Write-Host "" -ForegroundColor Yellow
            Write-Host "Possible solutions:" -ForegroundColor Yellow
            Write-Host "1. Enable VBA project access in Excel Trust Center" -ForegroundColor White
            Write-Host "2. Check if Excel is properly installed" -ForegroundColor White
            Write-Host "3. Run PowerShell as Administrator" -ForegroundColor White
            Write-Host "4. Check README.md" -ForegroundColor Cyan
        }

        # Cleanup on error
        if ($workbook) {
            try { $workbook.Close($false) } catch { }
        }
        if ($excel) {
            try { $excel.Quit() } catch { }
        }

        Write-Host "" -ForegroundColor Gray
        Write-Host "Continuing with standard XLSX export..." -ForegroundColor Gray
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