# Script to help configure Excel for VBA project access
# This script provides guidance and optional registry modifications

Write-Host "Excel VBA Access Setup Tool" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Check if Excel is installed
try {
    $excel = New-Object -ComObject Excel.Application
    $excelVersion = $excel.Version
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Excel $excelVersion detected" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Excel is not installed or COM interface is not available" -ForegroundColor Red
    Write-Host "Please install Microsoft Excel before using VBA features." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "VBA Integration requires specific Excel settings to be enabled." -ForegroundColor Yellow
Write-Host ""
Write-Host "MANUAL SETUP (Recommended):" -ForegroundColor Green
Write-Host "1. Open Excel" -ForegroundColor White
Write-Host "2. Go to File → Options → Trust Center → Trust Center Settings" -ForegroundColor White
Write-Host "3. Select 'Macro Settings'" -ForegroundColor White
Write-Host "4. Check 'Trust access to the VBA project object model'" -ForegroundColor White
Write-Host "5. Select 'Disable all macros with notification' (recommended)" -ForegroundColor White
Write-Host "6. Click OK and restart Excel" -ForegroundColor White
Write-Host ""

# Ask user if they want automated registry modification
Write-Host "AUTOMATED SETUP (Advanced):" -ForegroundColor Cyan
Write-Host "This script can modify registry settings automatically." -ForegroundColor Yellow
Write-Host "WARNING: Registry modifications can affect system behavior." -ForegroundColor Red
Write-Host ""

$choice = Read-Host "Do you want to automatically configure registry settings? (y/N)"

if ($choice -eq "y" -or $choice -eq "Y") {
    Write-Host ""
    Write-Host "Attempting to modify Excel registry settings..." -ForegroundColor Yellow

    try {
        # Determine Excel version for registry path
        $officeVersions = @("16.0", "15.0", "14.0", "12.0")  # 2016/2019/365, 2013, 2010, 2007
        $regPath = $null

        foreach ($version in $officeVersions) {
            $testPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Security"
            if (Test-Path $testPath) {
                $regPath = $testPath
                Write-Host "Found Excel registry path: $regPath" -ForegroundColor Green
                break
            }
        }

        if (-not $regPath) {
            throw "Could not find Excel registry path"
        }

        # Enable VBA project access
        Write-Host "Setting VBA project access..." -ForegroundColor Gray
        Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWORD -Force

        # Set macro security to notification level
        Write-Host "Setting macro security level..." -ForegroundColor Gray
        Set-ItemProperty -Path $regPath -Name "VBAWarnings" -Value 2 -Type DWORD -Force

        Write-Host ""
        Write-Host "Registry settings updated successfully!" -ForegroundColor Green
        Write-Host "Please restart Excel for changes to take effect." -ForegroundColor Yellow

    } catch {
        Write-Host ""
        Write-Host "Failed to modify registry settings: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please use manual setup method instead." -ForegroundColor Yellow
    }
} else {
    Write-Host "Skipping automated setup. Please use manual method." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "TESTING VBA ACCESS:" -ForegroundColor Cyan
Write-Host "After configuration, run your parking report script again." -ForegroundColor White
Write-Host "If VBA integration works, you will get an XLSM file with analysis tools." -ForegroundColor White
Write-Host ""
Write-Host "For troubleshooting, see: docs\VBA_Integration_Troubleshooting.md" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")