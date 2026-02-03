# ======================================================================
# SIMPLE LOG CHECKER - CSV ONLY (NO ERRORS)
# ======================================================================

Clear-Host
Write-Host "Windows Log Checker - CSV Export" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor White

# Bypass execution policy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# Get date range
Write-Host "`nEnter Date Range (DD-MM-YYYY Format)" -ForegroundColor Green
$startDate = Read-Host "Start Date (Example: 1-12-2025)"
$endDate = Read-Host "End Date (Example: 3-2-2026)"

# Convert dates
try {
    $StartTime = Get-Date $startDate
    $EndTime = Get-Date $endDate
}
catch {
    Write-Host "Invalid date format! Use DD-MM-YYYY" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

Write-Host "`nAnalyzing logs from $($StartTime.ToString('dd-MM-yyyy')) to $($EndTime.ToString('dd-MM-yyyy'))..." -ForegroundColor Yellow

# Get File Access Events
Write-Host "Checking File Access logs..." -ForegroundColor Gray
$fileEvents = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4663} -ErrorAction SilentlyContinue | 
    Where-Object { $_.TimeCreated -ge $StartTime -and $_.TimeCreated -le $EndTime }

# Get USB Events
Write-Host "Checking USB Connection logs..." -ForegroundColor Gray
$usbEvents = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=6416} -ErrorAction SilentlyContinue | 
    Where-Object { $_.TimeCreated -ge $StartTime -and $_.TimeCreated -le $EndTime }

# Create Desktop folder
$desktop = [Environment]::GetFolderPath("Desktop")
$folder = "$desktop\LogAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmm')"
New-Item -ItemType Directory -Path $folder -Force | Out-Null

# Export File Access to CSV
if ($fileEvents.Count -gt 0) {
    $fileData = $fileEvents | ForEach-Object {
        [PSCustomObject]@{
            Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $_.Id
            EventName = "File Access"
            User = if ($_.Message -match "Account Name:\s+(\S+)") { $matches[1] } else { "N/A" }
            FilePath = if ($_.Message -match "Object Name:\s+(.+)") { $matches[1] } else { "N/A" }
            Process = if ($_.Message -match "Process Name:\s+(.+)") { $matches[1] } else { "N/A" }
            Computer = $_.MachineName
        }
    }
    
    $filePath = "$folder\File_Access_Logs.csv"
    $fileData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
    Write-Host "✓ File Access: $($fileEvents.Count) events saved to CSV" -ForegroundColor Green
}

# Export USB Events to CSV
if ($usbEvents.Count -gt 0) {
    $usbData = $usbEvents | ForEach-Object {
        [PSCustomObject]@{
            Time = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $_.Id
            EventName = "USB Connection"
            User = if ($_.Message -match "Account Name:\s+(\S+)") { $matches[1] } else { "N/A" }
            Device = if ($_.Message -match "Device Description:\s+(.+)") { $matches[1] } else { "N/A" }
            Computer = $_.MachineName
        }
    }
    
    $usbPath = "$folder\USB_Connection_Logs.csv"
    $usbData | Export-Csv -Path $usbPath -NoTypeInformation -Encoding UTF8
    Write-Host "✓ USB Connections: $($usbEvents.Count) events saved to CSV" -ForegroundColor Green
}

# Create Summary
$summary = @"
Windows Log Analysis Summary
============================
Date: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')
Analyst: $env:USERNAME
Computer: $env:COMPUTERNAME

Date Range: $($StartTime.ToString('dd-MM-yyyy')) to $($EndTime.ToString('dd-MM-yyyy'))

Results:
--------
File Access Events: $($fileEvents.Count)
USB Connection Events: $($usbEvents.Count)

Location: $folder

Files:
------
File_Access_Logs.csv - Contains all file access/transfer logs
USB_Connection_Logs.csv - Contains all USB device connections
"@

$summary | Out-File -Path "$folder\Analysis_Summary.txt" -Encoding UTF8

# Open folder
Start-Process $folder

Write-Host "`n" + ("=" * 50) -ForegroundColor Green
Write-Host "COMPLETE! Reports saved to:" -ForegroundColor Cyan
Write-Host "$folder" -ForegroundColor White
Write-Host "=" * 50 -ForegroundColor Green

Read-Host "`nPress Enter to exit"