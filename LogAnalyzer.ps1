# ======================================================================
# CLEAN WINDOWS LOG ANALYZER - SIMPLE VERSION
# No syntax errors, No special characters
# ======================================================================

# Clear screen
Clear-Host

# Check Admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: Run as Administrator!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

# Bypass execution policy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Banner
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "    WINDOWS LOG ANALYZER - SIMPLE TOOL" -ForegroundColor Yellow
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

# Get date range
Write-Host "ENTER DATE RANGE (DD-MM-YYYY Format)" -ForegroundColor Green
Write-Host "--------------------------------------" -ForegroundColor Gray

# Start Date
do {
    $startInput = Read-Host "Start Date (Example: 1-12-2025)"
    try {
        $startDate = Get-Date $startInput
        Write-Host "✓ Start: $($startDate.ToString('dddd, dd MMMM yyyy'))" -ForegroundColor Green
        $validStart = $true
    }
    catch {
        Write-Host "✗ Invalid date. Try: 1-12-2025 or 01-12-2025" -ForegroundColor Red
        $validStart = $false
    }
} while (-not $validStart)

# Start Time
$startTime = Read-Host "Start Time (HH:MM) [Default: 00:00]"
if ([string]::IsNullOrWhiteSpace($startTime)) { $startTime = "00:00" }
$StartTime = Get-Date "$($startDate.ToString('yyyy-MM-dd')) $startTime"

# End Date
do {
    $endInput = Read-Host "End Date (Example: 3-2-2026)"
    try {
        $endDate = Get-Date $endInput
        Write-Host "✓ End: $($endDate.ToString('dddd, dd MMMM yyyy'))" -ForegroundColor Green
        $validEnd = $true
    }
    catch {
        Write-Host "✗ Invalid date. Try: 3-2-2026 or 03-02-2026" -ForegroundColor Red
        $validEnd = $false
    }
} while (-not $validEnd)

# End Time
$endTime = Read-Host "End Time (HH:MM) [Default: 23:59]"
if ([string]::IsNullOrWhiteSpace($endTime)) { $endTime = "23:59" }
$EndTime = Get-Date "$($endDate.ToString('yyyy-MM-dd')) $endTime"

# Show summary
Write-Host "`nANALYSIS RANGE:" -ForegroundColor Cyan
Write-Host "From: $($StartTime.ToString('dd-MM-yyyy HH:mm'))" -ForegroundColor White
Write-Host "To:   $($EndTime.ToString('dd-MM-yyyy HH:mm'))" -ForegroundColor White
Write-Host "Days: $([math]::Round(($EndTime - $StartTime).TotalDays, 1))" -ForegroundColor Yellow

# Create output folder on Desktop
$desktopPath = [Environment]::GetFolderPath("Desktop")
$folderName = "LogReport_$($StartTime.ToString('yyyyMMdd'))_$($EndTime.ToString('yyyyMMdd'))"
$outputPath = "$desktopPath\$folderName"

New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

# Event IDs to check
$eventsToCheck = @(
    @{ID=4663; Name="File Access"},
    @{ID=4656; Name="File Create"},
    @{ID=4660; Name="File Delete"},
    @{ID=4624; Name="User Logon"},
    @{ID=4634; Name="User Logoff"},
    @{ID=6416; Name="USB Device"},
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
            } elseif ($evt.Message -match "Subject:.+?Account Name:\s+(\S+)") {
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
            }
        }
    }
    catch {
        Write-Host " Error" -ForegroundColor Red
    }
}

# Generate reports
Write-Host "`nGENERATING REPORTS..." -ForegroundColor Cyan

# 1. CSV Report
$csvFile = "$outputPath\LogAnalysis_Detailed.csv"
$allResults | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
Write-Host "✓ CSV Report: $csvFile" -ForegroundColor Green

# 2. HTML Report
$htmlFile = "$outputPath\LogAnalysis_Report.html"
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
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
        <p><strong>Date Range:</strong> $($StartTime.ToString('dd-MM-yyyy HH:mm')) to $($EndTime.ToString('dd-MM-yyyy HH:mm'))</p>
        <p><strong>Computer:</strong> $env:COMPUTERNAME</p>
        <p><strong>Analyst:</strong> $env:USERNAME</p>
        <p><strong>Generated:</strong> $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')</p>
        <p><strong>Total Events Found:</strong> <span class="count">$($allResults.Count)</span></p>
    </div>
    
    <h2>Event Summary</h2>
    <table>
        <tr><th>Event Type</th><th>Event ID</th><th>Count</th></tr>
"@

# Add summary by event type
$summary = $allResults | Group-Object EventName
foreach ($group in $summary) {
    $htmlContent += "<tr><td>$($group.Name)</td><td>$($group.Group[0].EventID)</td><td class='count'>$($group.Count)</td></tr>"
}

$htmlContent += @"
    </table>
    
    <h2>Detailed Events</h2>
    <table>
        <tr>
            <th>Time</th>
            <th>Event</th>
            <th>User</th>
            <th>Details</th>
            <th>Computer</th>
        </tr>
"@

# Add detailed events (max 100)
$displayCount = [Math]::Min($allResults.Count, 100)
for ($i = 0; $i -lt $displayCount; $i++) {
    $row = $allResults[$i]
    
    $details = $row.FilePath
    if ($details -eq "N/A") { $details = $row.Device }
    if ($details -eq "N/A") { $details = $row.Process }
    
    $htmlContent += "<tr>"
    $htmlContent += "<td>$($row.Time)</td>"
    $htmlContent += "<td>$($row.EventName) ($($row.EventID))</td>"
    $htmlContent += "<td>$($row.User)</td>"
    $htmlContent += "<td>$details</td>"
    $htmlContent += "<td>$($row.Computer)</td>"
    $htmlContent += "</tr>"
}

if ($allResults.Count -gt 100) {
    $htmlContent += "<tr><td colspan='5'>... and $($allResults.Count - 100) more events in CSV file</td></tr>"
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
Duration: $([math]::Round(($EndTime - $StartTime).TotalDays, 1)) days

EVENT SUMMARY:
================================================
"@

foreach ($group in $summary) {
    $textContent += "$($group.Name): $($group.Count) events`n"
}

$textContent += @"

TOP 10 RECENT EVENTS:
================================================
"@

$allResults | Sort-Object Time -Descending | Select-Object -First 10 | ForEach-Object {
    $textContent += "$($_.Time) | $($_.EventName) | $($_.User) | $($_.FilePath)`n"
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

# Open the folder
Start-Process $outputPath

# Final message
Write-Host "`n" + ("=" * 50) -ForegroundColor Green
Write-Host "✅ ANALYSIS COMPLETE!" -ForegroundColor Green
Write-Host "=" * 50 -ForegroundColor Green
Write-Host "`n📁 Reports saved to: $outputPath" -ForegroundColor Cyan
Write-Host "📄 Files created:" -ForegroundColor White
Write-Host "   1. LogAnalysis_Report.html (Open in browser)" -ForegroundColor Yellow
Write-Host "   2. LogAnalysis_Detailed.csv (Open in Excel)" -ForegroundColor Yellow
Write-Host "   3. Analysis_Summary.txt (Quick summary)" -ForegroundColor Yellow

Write-Host "`n🔍 To analyze specific file transfers:" -ForegroundColor White
Write-Host "   Open CSV file in Excel and filter:" -ForegroundColor Gray
Write-Host "   - EventID = 4663 for file access" -ForegroundColor Gray
Write-Host "   - EventID = 6416 for USB activity" -ForegroundColor Gray

$openNow = Read-Host "`nOpen HTML report in browser now? (Y/N)"
if ($openNow -eq 'Y' -or $openNow -eq 'y') {
    Start-Process $htmlFile
}

Write-Host "`n🎯 Investigation completed successfully!" -ForegroundColor Cyan
Read-Host "Press Enter to exit"