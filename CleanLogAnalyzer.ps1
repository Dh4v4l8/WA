# ======================================================================
# WINDOWS LOG ANALYZER - CLEAN VERSION (ERROR FIXED)
# No HTML parsing errors, Simple and Working
# ======================================================================

# Clear screen
Clear-Host

# Banner
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "    WINDOWS LOG ANALYZER - SIMPLE TOOL" -ForegroundColor Yellow
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: Run PowerShell as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as administrator'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

# Bypass execution policy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Get date range from user
Write-Host "ENTER DATE RANGE (DD-MM-YYYY Format)" -ForegroundColor Green
Write-Host "--------------------------------------" -ForegroundColor Gray

# Start Date
do {
    $startDateInput = Read-Host "Start Date (Example: 1-12-2025)"
    try {
        $startDate = Get-Date $startDateInput
        Write-Host "✓ Start: $($startDate.ToString('dd MMMM yyyy'))" -ForegroundColor Green
        $validStart = $true
    }
    catch {
        Write-Host "✗ Invalid date. Try: 1-12-2025 or 01-12-2025" -ForegroundColor Red
        $validStart = $false
    }
} while (-not $validStart)

# Start Time
$startTimeInput = Read-Host "Start Time (HH:MM) [Press Enter for 00:00]"
if ([string]::IsNullOrWhiteSpace($startTimeInput)) {
    $startTimeInput = "00:00"
}

# End Date
do {
    $endDateInput = Read-Host "End Date (Example: 3-2-2026)"
    try {
        $endDate = Get-Date $endDateInput
        Write-Host "✓ End: $($endDate.ToString('dd MMMM yyyy'))" -ForegroundColor Green
        $validEnd = $true
    }
    catch {
        Write-Host "✗ Invalid date. Try: 3-2-2026 or 03-02-2026" -ForegroundColor Red
        $validEnd = $false
    }
} while (-not $validEnd)

# End Time
$endTimeInput = Read-Host "End Time (HH:MM) [Press Enter for 23:59]"
if ([string]::IsNullOrWhiteSpace($endTimeInput)) {
    $endTimeInput = "23:59"
}

# Create DateTime objects
$StartTime = Get-Date "$($startDate.ToString('yyyy-MM-dd')) $startTimeInput"
$EndTime = Get-Date "$($endDate.ToString('yyyy-MM-dd')) $endTimeInput"

# Show summary
Write-Host "`nANALYSIS RANGE:" -ForegroundColor Cyan
Write-Host "From: $($StartTime.ToString('dd-MM-yyyy HH:mm'))" -ForegroundColor White
Write-Host "To:   $($EndTime.ToString('dd-MM-yyyy HH:mm'))" -ForegroundColor White
$days = [math]::Round(($EndTime - $StartTime).TotalDays, 1)
Write-Host "Days: $days" -ForegroundColor Yellow

# Create output folder on Desktop
$desktopPath = [Environment]::GetFolderPath("Desktop")
$folderName = "LogReport_$($StartTime.ToString('yyyyMMdd'))_$($EndTime.ToString('yyyyMMdd'))"
$outputPath = "$desktopPath\$folderName"

if (Test-Path $outputPath) {
    Remove-Item -Path $outputPath -Recurse -Force
}
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

# Event IDs to check
$eventsToCheck = @(
    @{ID=4663; Name="File Access"},
    @{ID=4656; Name="File Creation"},
    @{ID=4660; Name="File Deletion"},
    @{ID=4624; Name="User Logon"},
    @{ID=4634; Name="User Logoff"},
    @{ID=6416; Name="USB Connection"},
    @{ID=4688; Name="Process Start"},
    @{ID=4798; Name="Group Check"}
)

# Start analysis
Write-Host "`nSEARCHING LOGS..." -ForegroundColor Cyan

$allResults = @()

foreach ($event in $eventsToCheck) {
    $id = $event.ID
    $name = $event.Name
    
    Write-Host "  Checking $name (ID: $id)..." -NoNewline -ForegroundColor Gray
    
    try {
        $filteredEvents = Get-WinEvent -FilterHashtable @{LogName="Security"; ID=$id} -ErrorAction SilentlyContinue | 
            Where-Object { $_.TimeCreated -ge $StartTime -and $_.TimeCreated -le $EndTime }
        
        Write-Host " Found: $($filteredEvents.Count)" -ForegroundColor Green
        
        # Add to results
        foreach ($evt in $filteredEvents) {
            # Parse user
            $user = "N/A"
            if ($evt.Message -match "Account Name:\s+(\S+)") {
                $user = $matches[1]
            }
            
            # Parse file path
            $filePath = "N/A"
            if ($evt.Message -match "Object Name:\s+(.+)") {
                $filePath = $matches[1]
            }
            
            # Parse process
            $process = "N/A"
            if ($evt.Message -match "Process Name:\s+(.+)") {
                $process = $matches[1]
            }
            
            # Parse device (for USB)
            $device = "N/A"
            if ($evt.Message -match "Device Description:\s+(.+)") {
                $device = $matches[1]
            }
            
            $allResults += [PSCustomObject]@{
                Time = $evt.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                EventID = $id
                EventName = $name
                User = $user
                FilePath = $filePath
                Process = $process
                Device = $device
                Computer = $evt.MachineName
                Message = $evt.Message.Substring(0, [Math]::Min(200, $evt.Message.Length))
            }
        }
    }
    catch {
        Write-Host " Error" -ForegroundColor Red
    }
}

# Generate reports
Write-Host "`nGENERATING REPORTS..." -ForegroundColor Cyan

# 1. CSV Report (Main Report)
if ($allResults.Count -gt 0) {
    $csvFile = "$outputPath\LogAnalysis_Detailed.csv"
    $allResults | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
    Write-Host "✓ CSV Report: $csvFile" -ForegroundColor Green
    
    # 2. Create Simple HTML Report (FIXED VERSION)
    $htmlFile = "$outputPath\LogAnalysis_Report.html"
    
    # Build HTML content piece by piece (to avoid parsing errors)
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Windows Log Analysis Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 30px; }
        h1 { color: #2c3e50; }
        h2 { color: #34495e; border-bottom: 2px solid #3498db; padding-bottom: 5px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th { background: #3498db; color: white; padding: 10px; text-align: left; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        .summary { background: #ecf0f1; padding: 15px; border-radius: 5px; }
        .count { color: #e74c3c; font-weight: bold; }
        .footer { margin-top: 30px; color: #7f8c8d; font-size: 12px; }
    </style>
</head>
<body>
    <h1>Windows Security Log Analysis Report</h1>
    
    <div class="summary">
        <h3>Analysis Summary</h3>
"@

    # Add summary details
    $htmlContent += "<p><strong>Date Range:</strong> " + $StartTime.ToString('dd-MM-yyyy HH:mm') + " to " + $EndTime.ToString('dd-MM-yyyy HH:mm') + "</p>"
    $htmlContent += "<p><strong>Computer:</strong> $env:COMPUTERNAME</p>"
    $htmlContent += "<p><strong>Analyst:</strong> $env:USERNAME</p>"
    $htmlContent += "<p><strong>Generated:</strong> $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')</p>"
    $htmlContent += "<p><strong>Total Events Found:</strong> <span class='count'>$($allResults.Count)</span></p>"
    
    $htmlContent += @"
    </div>
    
    <h2>Event Summary</h2>
    <table>
        <tr><th>Event Type</th><th>Event ID</th><th>Count</th></tr>
"@

    # Add summary rows
    $summary = $allResults | Group-Object EventName
    foreach ($group in $summary) {
        $htmlContent += "<tr><td>$($group.Name)</td><td>$($group.Group[0].EventID)</td><td class='count'>$($group.Count)</td></tr>"
    }

    $htmlContent += @"
    </table>
    
    <h2>Detailed Events (First 50)</h2>
    <table>
        <tr>
            <th>Time</th>
            <th>Event</th>
            <th>User</th>
            <th>File/Device</th>
            <th>Computer</th>
        </tr>
"@

    # Add detailed events
    $displayCount = [Math]::Min($allResults.Count, 50)
    for ($i = 0; $i -lt $displayCount; $i++) {
        $row = $allResults[$i]
        
        $details = $row.FilePath
        if ($details -eq "N/A") { $details = $row.Device }
        if ($details -eq "N/A") { $details = $row.Process }
        
        $htmlContent += "<tr>"
        $htmlContent += "<td>$($row.Time)</td>"
        $htmlContent += "<td>$($row.EventName)</td>"
        $htmlContent += "<td>$($row.User)</td>"
        $htmlContent += "<td>$details</td>"
        $htmlContent += "<td>$($row.Computer)</td>"
        $htmlContent += "</tr>"
    }

    if ($allResults.Count -gt 50) {
        $remaining = $allResults.Count - 50
        $htmlContent += "<tr><td colspan='5'>... and $remaining more events in CSV file</td></tr>"
    }

    $htmlContent += @"
    </table>
    
    <div class="footer">
        <hr>
        <p>Report generated by Windows Log Analyzer Tool</p>
        <p>Location: $outputPath</p>
        <p>Files: LogAnalysis_Detailed.csv (Excel), LogAnalysis_Report.html (Browser)</p>
    </div>
</body>
</html>
"@

    # Save HTML file
    $htmlContent | Out-File -FilePath $htmlFile -Encoding UTF8
    Write-Host "✓ HTML Report: $htmlFile" -ForegroundColor Green
    
    # 3. Text Summary
    $textFile = "$outputPath\Analysis_Summary.txt"
    
    $textContent = @"
================================================
WINDOWS LOG ANALYSIS - SUMMARY REPORT
================================================
Analysis Date: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')
Computer: $env:COMPUTERNAME
Analyst: $env:USERNAME

ANALYSIS PERIOD:
================================================
Start: $($StartTime.ToString('dd-MM-yyyy HH:mm'))
End:   $($EndTime.ToString('dd-MM-yyyy HH:mm'))
Duration: $days days

EVENT SUMMARY:
================================================
"@

    foreach ($group in $summary) {
        $textContent += "$($group.Name): $($group.Count) events`n"
    }

    $textContent += @"

TOP 10 FILE ACCESS EVENTS:
================================================
"@

    $fileAccess = $allResults | Where-Object { $_.EventID -eq 4663 } | Sort-Object Time -Descending | Select-Object -First 10
    foreach ($event in $fileAccess) {
        $textContent += "$($event.Time) | $($event.User) | $($event.FilePath)`n"
    }

    $textContent += @"

TOP 10 USB DEVICE CONNECTIONS:
================================================
"@

    $usbEvents = $allResults | Where-Object { $_.EventID -eq 6416 } | Sort-Object Time -Descending | Select-Object -First 10
    foreach ($event in $usbEvents) {
        $textContent += "$($event.Time) | $($event.User) | $($event.Device)`n"
    }

    $textContent += @"

REPORT FILES:
================================================
1. LogAnalysis_Detailed.csv - Open in Excel
2. LogAnalysis_Report.html - Open in browser
3. Analysis_Summary.txt - This file

LOCATION: $outputPath

INVESTIGATION NOTES:
================================================
1. File Access (4663): Track file open/copy operations
2. USB Connections (6416): Check external device usage
3. User Logons (4624): Monitor login activities
4. Process Creation (4688): Check program executions

================================================
END OF REPORT
================================================
"@

    $textContent | Out-File -FilePath $textFile -Encoding UTF8
    Write-Host "✓ Text Summary: $textFile" -ForegroundColor Green
}
else {
    Write-Host "No events found in the specified date range." -ForegroundColor Yellow
    $noEventsFile = "$outputPath\No_Events_Found.txt"
    "No security events found between $($StartTime.ToString('dd-MM-yyyy HH:mm')) and $($EndTime.ToString('dd-MM-yyyy HH:mm'))" | Out-File -FilePath $noEventsFile
}

# Open the folder
Start-Process $outputPath

# Final message
Write-Host "`n" + ("=" * 50) -ForegroundColor Green
Write-Host "ANALYSIS COMPLETE!" -ForegroundColor Green
Write-Host "=" * 50 -ForegroundColor Green
Write-Host "`nReports saved to: $outputPath" -ForegroundColor Cyan

if ($allResults.Count -gt 0) {
    Write-Host "Events found: $($allResults.Count)" -ForegroundColor White
    Write-Host "`nFiles created:" -ForegroundColor Yellow
    Write-Host "  1. LogAnalysis_Report.html (Open in browser)" -ForegroundColor White
    Write-Host "  2. LogAnalysis_Detailed.csv (Open in Excel)" -ForegroundColor White
    Write-Host "  3. Analysis_Summary.txt (Quick summary)" -ForegroundColor White
    
    $openNow = Read-Host "`nOpen HTML report in browser now? (Y/N)"
    if ($openNow -eq 'Y' -or $openNow -eq 'y') {
        Start-Process "$outputPath\LogAnalysis_Report.html"
    }
}
else {
    Write-Host "No security events found in the specified date range." -ForegroundColor Yellow
}

Write-Host "`nInvestigation completed!" -ForegroundColor Cyan
Read-Host "Press Enter to exit"