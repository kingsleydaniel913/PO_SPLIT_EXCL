# Kill any existing Excel processes to unlock files
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force

# --- CONFIGURATION ---
$myPrinter = "Microsoft Print to PDF" # Change this to your exact printer name
$rawSheet  = "RAW"
$pivotSheet = "INVENTORY"

try {
    Write-Host "--- SCRIPT STARTING ---" -ForegroundColor Cyan
    
    # 1. Check OneDrive Environment Variable
    Write-Host "Step 1: Checking OneDrive path..." -NoNewline
    $oneDrivePath = $env:OneDriveCommercial
    if (!$oneDrivePath) { $oneDrivePath = $env:OneDrive }
    
    if (!$oneDrivePath) {
        Write-Host " FAILED" -ForegroundColor Red
        throw "Could not find OneDrive environment variable. Are you logged into OneDrive?"
    }
    Write-Host " SUCCESS ($oneDrivePath)" -ForegroundColor Green

    # 2. Define and Check Files
    $sourcePath = Join-Path $oneDrivePath "123.xlsx"
    $templatePath = Join-Path $oneDrivePath "LimeInventory.xlsx"

    Write-Host "Step 2: Checking if 123.xlsx exists..." -NoNewline
    if (!(Test-Path $sourcePath)) { 
        Write-Host " FAILED" -ForegroundColor Red
        throw "File NOT FOUND: $sourcePath" 
    }
    Write-Host " FOUND" -ForegroundColor Green

    Write-Host "Step 3: Checking if LimeInventory.xlsx exists..." -NoNewline
    if (!(Test-Path $templatePath)) { 
        Write-Host " FAILED" -ForegroundColor Red
        throw "File NOT FOUND: $templatePath" 
    }
    Write-Host " FOUND" -ForegroundColor Green

    # --- NEW: PRINTER SELECTION ---
    Write-Host "Checking Printers..."
    Get-Printer | ForEach-Object {
        if ($_.Name -eq $myPrinter) {
            Write-Host " [SELECTED] $($_.Name)" -ForegroundColor Green
        } else {
            Write-Host "            $($_.Name)" -ForegroundColor Gray
        }
    }
    # ------------------------------



    # 3. Try to launch Excel
    Write-Host "Step 4: Launching Excel Background Process..." -NoNewline
    $xl = New-Object -ComObject Excel.Application
    $xlPid = (Get-Process -Name "Excel" | Sort-Object StartTime -Descending | Select-Object -First 1).Id
    if (!$xl) { 
        Write-Host " FAILED" -ForegroundColor Red
        throw "Could not start Excel. Is Microsoft Office installed on this RDS?" 
    }
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    Write-Host " SUCCESS" -ForegroundColor Green

    # 4. Open Workbooks
    Write-Host "Step 5: Opening Workbooks..." -NoNewline
    $wbSource = $xl.Workbooks.Open($sourcePath)
    $wbTemplate = $xl.Workbooks.Open($templatePath)
    Write-Host " SUCCESS" -ForegroundColor Green

    # 5. Check Sheets
    $rawSheetName = "RAW"
    $pivotSheetName = "INVENTORY" # <--- DOUBLE CHECK THIS NAME

    Write-Host "Step 6: Accessing Sheets..." -NoNewline
    try {
        $wsSource = $wbSource.Worksheets.Item(1) 
        $wsTemplate = $wbTemplate.Worksheets.Item($rawSheetName)
        $wsPivot = $wbTemplate.Worksheets.Item($pivotSheetName)
        Write-Host " SUCCESS" -ForegroundColor Green
    } catch {
        Write-Host " FAILED" -ForegroundColor Red
        throw "Could not find one of the tabs. Check if '$rawSheetName' or '$pivotSheetName' is spelled correctly."
    }

    # 6. Data Transfer
    Write-Host "Step 7: Transferring Data..." -NoNewline
    # Clear the ENTIRE sheet (not just used range) to kill ghost data
    $wsTemplate.Cells.Clear() 
    $wsSource.UsedRange.Copy($wsTemplate.Range("A1"))
    Write-Host " SUCCESS" -ForegroundColor Green

    # 7. Refresh Pivot
    Write-Host "Step 8: Updating Table1 & Refreshing..." -NoNewline
    try {
        # Remove old Table definition if it exists (keeps data, just removes the 'Table' name)
        foreach ($list in $wsTemplate.ListObjects) { if ($list.Name -eq "Table1") { $list.Unlist() } }

        # Create a brand new 'Table1' covering the new data
        $newTable = $wsTemplate.ListObjects.Add(1, $wsTemplate.UsedRange, $null, 1) # 1 = xlSrcRange, 1 = xlYes (headers)
        $newTable.Name = "Table1"

        foreach ($pivot in $wsPivot.PivotTables()) {
            $pivot.PivotCache().Refresh()
            $pivot.Update()
        }
        Write-Host " SUCCESS" -ForegroundColor Green
    } catch {
        Write-Host " FAILED" -ForegroundColor Red
        Write-Warning "Pivot Error: $($_.Exception.Message)"
    }

    # 7. Apply Page Settings (Matching your Screenshot)
    Write-Host "Step 8.5: Applying Page Layout..." -NoNewline
    $setup = $wsPivot.PageSetup
    
    # You MUST set Zoom to false before FitToPages properties will work
    $setup.Zoom = $false
    
    $setup.Orientation = 1 # xlPortrait
    $setup.PaperSize   = 1 # xlPaperLetter
    
    # Narrow Margins (0.75 top/bottom, 0.25 sides)
    $setup.LeftMargin   = $xl.InchesToPoints(0.25)
    $setup.RightMargin  = $xl.InchesToPoints(0.25)
    $setup.TopMargin    = $xl.InchesToPoints(0.75)
    $setup.BottomMargin = $xl.InchesToPoints(0.75)

    # Fit All Columns on One Page
    $setup.FitToPagesWide = 1
    $setup.FitToPagesTall = $false # $false = "Automatic" (allows multiple pages tall)
    
    Write-Host " SUCCESS" -ForegroundColor Green

    Write-Host "Step 9: Sending to $myPrinter..." -NoNewline
    
    # If printing to PDF, we make Excel visible for a second so the "Save As" dialog pops up
    if ($myPrinter -eq "Microsoft Print to PDF") { $xl.Visible = $true }

    # Parameters: From, To, Copies, Preview, ActivePrinter
    $wsPivot.PrintOut([Type]::Missing, [Type]::Missing, 1, $false, $myPrinter)
    
    # Hide it again immediately
    $xl.Visible = $false
    
    Write-Host " SUCCESS" -ForegroundColor Green

    # 9. Save and Cleanup
    Write-Host "Step 10: Saving and Closing..." -NoNewline
    $wbTemplate.Save()
    $wbTemplate.Close()
    $wbSource.Close($false)
    $xl.Quit()
    Write-Host " SUCCESS" -ForegroundColor Green

} catch {
    Write-Host "`n`n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
    Write-Host "ERROR DETECTED:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
} finally {
    Write-Host "`n--- CLEANUP ---" -ForegroundColor Cyan
    
    if ($xl) {
        Write-Host "Closing Excel Workbooks..."
        try {
            $xl.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        } catch { 
            Write-Host "Excel was already closed or unresponsive." 
        }
    }

    # The "Nuclear" Option: Force kill the specific process ID
    if ($xlPid) {
        Write-Host "Force-terminating Excel process (PID: $xlPid) to unlock files..."
        Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue
    }

    Write-Host "Cleanup Complete. Files are now unlocked." -ForegroundColor Green
    Read-Host -Prompt "Press ENTER to close this window"
}
