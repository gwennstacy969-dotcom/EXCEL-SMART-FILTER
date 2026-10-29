# Test_Excel_Filter.ps1
# Automates and validates EXCEL SMART FILTER core logic

Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "  EXCEL SMART FILTER - AUTOMATED VERIFICATION SUITE  " -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan

$sourcePath = (Join-Path (Get-Location) "Sample_Data_Bulanan.xlsx")
$sourceFileInfo = Get-Item $sourcePath
$sourceOriginalHash = (Get-FileHash $sourcePath).Hash
Write-Host "[1] Source file hash before testing: $sourceOriginalHash"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

# 1. Test Working Copy Creation (Original File Protection)
$tempDir = [System.IO.Path]::GetTempPath()
$workCopyPath = Join-Path $tempDir ("SMART_FILTER_TEST_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".xlsx")
Copy-Item -Path $sourcePath -Destination $workCopyPath
Write-Host "[2] Working copy created at: $workCopyPath" -ForegroundColor Green

$wb = $xl.Workbooks.Open($workCopyPath)
$ws = $wb.Worksheets.Item("Transaksi_Bulanan")

# 2. Test Header Detection Logic
# Finds the row with maximum text headers
$detectedHeaderRow = 1
for ($r = 1; $r -le 10; $r++) {
    $nonEmptyCols = 0
    for ($c = 1; $c -le 12; $c++) {
        $txt = [string]$ws.Cells.Item($r, $c).Text
        if ($txt.Trim().Length -gt 0) { $nonEmptyCols++ }
    }
    if ($nonEmptyCols -ge 5) {
        $detectedHeaderRow = $r
        break
    }
}
Write-Host "[3] Header Row Detection: Detected Row $detectedHeaderRow (Expected: 2)" -ForegroundColor Green
if ($detectedHeaderRow -ne 2) { throw "Header row detection failed!" }

# 3. Test Date Column Detection
$dateCol = 0
for ($c = 1; $c -le 12; $c++) {
    $hText = [string]$ws.Cells.Item($detectedHeaderRow, $c).Text
    if ($hText -match "tanggal|tgl|date|periode") {
        $dateCol = $c
        break
    }
}
Write-Host "[4] Date Column Detection: Detected Col $dateCol ($($ws.Cells.Item(2, $dateCol).Text))" -ForegroundColor Green
if ($dateCol -ne 2) { throw "Date column detection failed!" }

$lastRow = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
$totalDataRows = $lastRow - $detectedHeaderRow
Write-Host "[5] Total Data Rows: $totalDataRows"

$filterRange = $ws.Range("A2:L$lastRow")

# Helper function to count visible rows
function Get-VisibleDataRowCount($worksheet, $hRow, $lRow) {
    $count = 0
    for ($row = $hRow + 1; $row -le $lRow; $row++) {
        if (-not $worksheet.Rows.Item($row).Hidden) {
            $count++
        }
    }
    return $count
}

# 4. Test Date Period: Per Hari (2026-10-15)
# Boundary: >= 2026-10-15 00:00:00 and < 2026-10-16 00:00:00
$d1 = [DateTime]::Parse("2026-10-15").ToOADate()
$d2 = [DateTime]::Parse("2026-10-16").ToOADate()
$filterRange.AutoFilter($dateCol, ">=$d1", 1, "<$d2")
$visDaily = Get-VisibleDataRowCount $ws $detectedHeaderRow $lastRow
Write-Host "[6] Date Period [Per Hari: 2026-10-15] Visible Rows: $visDaily (Expected: 2)" -ForegroundColor Yellow
if ($visDaily -ne 2) { throw "Per Hari filter failed! Expected 2, got $visDaily" }

# 5. Test Date Period: Per Minggu (Reference: 2026-10-14 Wednesday)
# Monday is Oct 12, Sunday is Oct 18 (End boundary: < Oct 19)
$refDate = [DateTime]::Parse("2026-10-14")
# DayOfWeek in .NET: Sunday=0, Monday=1, ...
$diff = [int]$refDate.DayOfWeek - 1
if ($diff -lt 0) { $diff = 6 } # If Sunday, diff is 6
$monday = $refDate.AddDays(-$diff)
$sunday = $monday.AddDays(6)
$weekEndBoundary = $sunday.AddDays(1)
Write-Host "    Calculated Week: $($monday.ToString('yyyy-MM-dd')) s/d $($sunday.ToString('yyyy-MM-dd'))"

$ws.ShowAllData()
$filterRange.AutoFilter($dateCol, ">=$($monday.ToOADate())", 1, "<$($weekEndBoundary.ToOADate())")
$visWeekly = Get-VisibleDataRowCount $ws $detectedHeaderRow $lastRow
Write-Host "[7] Date Period [Per Minggu: 12-18 Okt] Visible Rows: $visWeekly (Expected: 9)" -ForegroundColor Yellow
if ($visWeekly -ne 9) { throw "Per Minggu filter failed! Expected 9, got $visWeekly" }

# 6. Test Weekly Cross-Month Boundary (Reference: 2026-11-01 Sunday)
# Monday is Oct 26, Sunday is Nov 1
$refDateNov = [DateTime]::Parse("2026-11-01")
$diffNov = [int]$refDateNov.DayOfWeek - 1
if ($diffNov -lt 0) { $diffNov = 6 }
$mondayNov = $refDateNov.AddDays(-$diffNov)
$sundayNov = $mondayNov.AddDays(6)
$weekEndNov = $sundayNov.AddDays(1)
Write-Host "    Calculated Cross-Month Week: $($mondayNov.ToString('yyyy-MM-dd')) s/d $($sundayNov.ToString('yyyy-MM-dd'))"

$ws.ShowAllData()
$filterRange.AutoFilter($dateCol, ">=$($mondayNov.ToOADate())", 1, "<$($weekEndNov.ToOADate())")
$visCrossMonth = Get-VisibleDataRowCount $ws $detectedHeaderRow $lastRow
Write-Host "[8] Date Period [Cross-Month Week: 26 Okt - 01 Nov] Visible Rows: $visCrossMonth (Expected: 10)" -ForegroundColor Yellow
if ($visCrossMonth -ne 10) { throw "Cross-Month filter failed! Expected 10, got $visCrossMonth" }

# 7. Test Column Filter Multi-Value (Kategori Layanan = "Sewa Kendaraan" or "Logistik")
$ws.ShowAllData()
$catCol = 4 # Kategori Layanan
$categories = [string[]]@("Sewa Kendaraan", "Logistik")
$filterRange.AutoFilter($catCol, $categories, 7) # 7 = xlFilterValues
$visMultiCat = Get-VisibleDataRowCount $ws $detectedHeaderRow $lastRow
Write-Host "[9] Column Filter [Kategori: Sewa + Logistik] Visible Rows: $visMultiCat (Expected: 31)" -ForegroundColor Yellow
if ($visMultiCat -ne 31) { throw "Multi-category filter failed! Expected 31, got $visMultiCat" }

# 8. Test Blank Cell Handling (Driver / PIC = Blanks)
$ws.ShowAllData()
$driverCol = 6 # Driver / PIC
$filterRange.AutoFilter($driverCol, "=")
$visBlanks = Get-VisibleDataRowCount $ws $detectedHeaderRow $lastRow
Write-Host "[10] Blank Cell Filter [Driver = (Kosong)] Visible Rows: $visBlanks (Expected: 8)" -ForegroundColor Yellow
if ($visBlanks -ne 8) { throw "Blank filter failed! Expected 8, got $visBlanks" }

# 9. Test Combined Multi-Criteria Filter:
# Date Range: 2026-10-01 to 2026-10-15
# AND Kategori Layanan = "Sewa Kendaraan"
# AND Status Pembayaran = "Lunas"
$ws.ShowAllData()
$dStart = [DateTime]::Parse("2026-10-01").ToOADate()
$dEnd = [DateTime]::Parse("2026-10-16").ToOADate() # Up to 2026-10-15 23:59:59
$filterRange.AutoFilter($dateCol, ">=$dStart", 1, "<$dEnd")
$filterRange.AutoFilter(4, [string[]]@("Sewa Kendaraan"), 7)
$filterRange.AutoFilter(11, [string[]]@("Lunas"), 7)

$visCombined = Get-VisibleDataRowCount $ws $detectedHeaderRow $lastRow
Write-Host "[11] Combined Filter [Date <= 15 Okt AND Sewa Kendaraan AND Lunas] Visible Rows: $visCombined" -ForegroundColor Green

# 10. Verify Data & Format Preservation:
# Row 3 formula must remain =H3*I3, values must NOT be deleted
$fRow3 = $ws.Range("J3").Formula
Write-Host "[12] Preserved Formula in J3: $fRow3 (Must be =H3*I3)" -ForegroundColor Green
if ($fRow3 -ne "=H3*I3") { throw "Formula was altered!" }

# Total rows in sheet must still be 43 (headers + 41 data rows)
$checkLastRow = $ws.UsedRange.Rows.Count
Write-Host "[13] Total Sheet Rows: $checkLastRow (Must be 43, no rows deleted!)" -ForegroundColor Green
if ($checkLastRow -ne 43) { throw "Rows were deleted instead of hidden!" }

# 11. Test Save Result to Copy
$savedResultPath = Join-Path (Get-Location) "Sample_Data_Bulanan_Filtered_20261029.xlsx"
$wb.SaveCopyAs($savedResultPath)
Write-Host "[14] Saved filtered copy to: $savedResultPath" -ForegroundColor Green

$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Remove-Item $workCopyPath -Force

# 12. Verify Source File Intact
$sourceFinalHash = (Get-FileHash $sourcePath).Hash
Write-Host "[15] Source file hash after testing: $sourceFinalHash"
if ($sourceOriginalHash -ne $sourceFinalHash) {
    throw "CRITICAL VIOLATION: Source file was modified during processing!"
} else {
    Write-Host "[16] SOURCE FILE 100% UNTOUCHED AND UNALTERED! Verification PASSED!" -ForegroundColor Green
}

Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "  ALL TESTS PASSED SUCCESSFULLY! 100% COMPLIANT!    " -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan

