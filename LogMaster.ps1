# ======================================================================
# LOGMASTER PRO - ADVANCED WINDOWS FORENSIC ANALYSIS TOOL
# Version: 2.0
# Author: Forensic Investigator
# Features: Complete Log Analysis, Timeline, USB Tracking, HTML Reports
# ======================================================================

#region Initial Setup
# ====================

# Clear screen and show banner
Clear-Host

# ASCII Art Banner
Write-Host @"

▓█████▄  ▒█████   ▄████▄   ██▀███   ██▓ ▄████▄   █    ██  ██▓    
▒██▀ ██▌▒██▒  ██▒▒██▀ ▀█  ▓██ ▒ ██▒▓██▒▒██▀ ▀█   ██  ▓██▒▓██▒    
░██   █▌▒██░  ██▒▒▓█    ▄ ▓██ ░▄█ ▒▒██▒▒▓█    ▄ ▓██  ▒██░▒██░    
░▓█▄   ▌▒██   ██░▒▓▓▄ ▄██▒▒██▀▀█▄  ░██░▒▓▓▄ ▄██▒▓▓█  ░██░▒██░    
░▒████▓ ░ ████▓▒░▒ ▓███▀ ░░██▓ ▒██▒░██░▒ ▓███▀ ░▒▒█████▓ ░██████▒
 ▒▒▓  ▒ ░ ▒░▒░▒░ ░ ░▒ ▒  ░░ ▒▓ ░▒▓░░▓  ░ ░▒ ▒  ░░▒▓▒ ▒ ▒ ░ ▒░▓  ▴
 ░ ▒  ▒   ░ ▒ ▒░   ░  ▒     ░▒ ░ ▒░ ▒ ░  ░  ▒   ░░▒░ ░ ░ ░ ░ ▒  ▄
 ░ ░  ░ ░ ░ ░ ▒  ░          ░░   ░  ▒ ░░         ░░░ ░ ░   ░ ░ ░▒
   ░        ░ ░  ░ ░         ░      ░  ░ ░         ░         ░  ░
 ░              ░                     ░                          
                                                                 
          WINDOWS FORENSIC ANALYSIS SUITE v2.0
"@ -ForegroundColor Cyan

Write-Host "`n          🛡️  Advanced Log Analysis & Reporting Tool" -ForegroundColor Yellow
Write-Host "          ===========================================" -ForegroundColor White

# Check Administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "`n❌ ERROR: This script requires Administrator privileges!" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    Read-Host "`nPress Enter to exit"
    exit
}

# Set execution policy for this session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Create required directories
$BasePath = "C:\ForensicAnalysis"
$ReportPath = "$BasePath\Reports"
$ExportPath = "$env:USERPROFILE\Desktop\Forensic_Reports"

if (-not (Test-Path $BasePath)) { New-Item -ItemType Directory -Path $BasePath -Force | Out-Null }
if (-not (Test-Path $ReportPath)) { New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null }
if (-not (Test-Path $ExportPath)) { New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null }

# Log file setup
$LogFile = "$ReportPath\Analysis_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
Start-Transcript -Path $LogFile -Append

#endregion

#region Configuration
# ===================

# Color scheme
$ColorHeader = "Cyan"
$ColorSuccess = "Green"
$ColorWarning = "Yellow"
$ColorError = "Red"
$ColorInfo = "White"
$ColorHighlight = "Magenta"

# Event IDs to monitor
$EventConfig = @{
    FileAccess = @(4663)
    FileCreation = @(4656)
    FileDeletion = @(4660)
    ProcessCreation = @(4688)
    UserLogon = @(4624, 4625)
    UserLogoff = @(4634)
    USBConnection = @(6416)
    AccountChange = @(4720, 4722, 4725, 4726, 4738, 4798, 4799)
    PolicyChange = @(4719, 4907)
    PrivilegeUse = @(4672, 4673, 4674)
}

# Time periods for analysis
$TimePeriods = @{
    "Last 1 Hour" = (Get-Date).AddHours(-1)
    "Last 4 Hours" = (Get-Date).AddHours(-4)
    "Last 24 Hours" = (Get-Date).AddDays(-1)
    "Last 7 Days" = (Get-Date).AddDays(-7)
    "Last 30 Days" = (Get-Date).AddDays(-30)
}

#endregion

#region Functions
# ===============

function Show-Menu {
    param([string]$Title = "MAIN MENU")
    
    Clear-Host
    Write-Host "`n" + ("=" * 70) -ForegroundColor $ColorHeader
    Write-Host "   $Title" -ForegroundColor $ColorHeader -NoNewline
    Write-Host "   [$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]" -ForegroundColor $ColorInfo
    Write-Host ("=" * 70) -ForegroundColor $ColorHeader
}

function Get-TimeRange {
    Show-Menu -Title "TIME RANGE SELECTION"
    
    Write-Host "`n📅 SELECT TIME RANGE FOR ANALYSIS:" -ForegroundColor $ColorHighlight
    Write-Host "──────────────────────────────────────────────────────" -ForegroundColor $ColorInfo
    
    $i = 1
    foreach ($period in $TimePeriods.Keys) {
        Write-Host "  $i. $period" -ForegroundColor $ColorInfo
        $i++
    }
    
    Write-Host "  $i. Custom Date Range" -ForegroundColor $ColorHighlight
    
    Write-Host "`n  X. Exit Program" -ForegroundColor $ColorError
    Write-Host "──────────────────────────────────────────────────────" -ForegroundColor $ColorInfo
    
    $choice = Read-Host "`nEnter your choice (1-$i)"
    
    if ($choice -eq 'X' -or $choice -eq 'x') {
        Write-Host "Exiting program..." -ForegroundColor $ColorWarning
        Stop-Transcript
        exit
    }
    
    if ($choice -eq $i.ToString()) {
        # Custom date range
        Write-Host "`n📝 ENTER CUSTOM DATE RANGE:" -ForegroundColor $ColorHighlight
        
        $startDate = Read-Host "Start Date (YYYY-MM-DD)"
        $startTime = Read-Host "Start Time (HH:MM:SS)"
        $endDate = Read-Host "End Date (YYYY-MM-DD)"
        $endTime = Read-Host "End Time (HH:MM:SS)"
        
        try {
            $StartTime = Get-Date "$startDate $startTime"
            $EndTime = Get-Date "$endDate $endTime"
        }
        catch {
            Write-Host "Invalid date format! Using default (Last 24 hours)" -ForegroundColor $ColorError
            $StartTime = (Get-Date).AddDays(-1)
            $EndTime = Get-Date
        }
    }
    else {
        $periodNames = @($TimePeriods.Keys)
        $selectedPeriod = $periodNames[$choice - 1]
        $StartTime = $TimePeriods[$selectedPeriod]
        $EndTime = Get-Date
        
        Write-Host "Selected: $selectedPeriod" -ForegroundColor $ColorSuccess
    }
    
    return @{
        Start = $StartTime
        End = $EndTime
        RangeText = if ($selectedPeriod) { $selectedPeriod } else { "Custom: $StartTime to $EndTime" }
    }
}

function Get-EventDetails {
    param($Event)
    
    $details = @{}
    $message = $Event.Message
    
    # Parse common patterns
    $patterns = @{
        User = @("Account Name:\s+(\S+)", "Subject:.+?Account Name:\s+(\S+)")
        Computer = @("Computer:\s+(\S+)", "Workstation Name:\s+(\S+)")
        FilePath = @("Object Name:\s+(.+)", "File Name:\s+(.+)")
        Process = @("Process Name:\s+(.+)", "New Process Name:\s+(.+)")
        IPAddress = @("Source Network Address:\s+(\S+)", "IP Address:\s+(\S+)")
        Device = @("Device Description:\s+(.+)", "Device Name:\s+(.+)")
        AccessMask = @("Accesses:\s+(.+)", "Access Request Information:\s+(.+)")
        Status = @("Status:\s+(\S+)", "Result Status:\s+(\S+)")
    }
    
    foreach ($key in $patterns.Keys) {
        foreach ($pattern in $patterns[$key]) {
            if ($message -match $pattern) {
                $details[$key] = $matches[1].Trim()
                break
            }
        }
        if (-not $details.ContainsKey($key)) {
            $details[$key] = "N/A"
        }
    }
    
    return $details
}

function Export-ToHTML {
    param($Data, $Title, $FileName)
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>$Title</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; background: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        h2 { color: #34495e; margin-top: 30px; }
        .summary { background: #ecf0f1; padding: 20px; border-radius: 5px; margin: 20px 0; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th { background: #3498db; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background: #f9f9f9; }
        .critical { background: #ffebee !important; }
        .warning { background: #fff3e0 !important; }
        .success { background: #e8f5e9 !important; }
        .timestamp { font-family: 'Consolas', monospace; }
        .badge { display: inline-block; padding: 3px 8px; border-radius: 12px; font-size: 12px; font-weight: bold; }
        .badge-success { background: #2ecc71; color: white; }
        .badge-warning { background: #f39c12; color: white; }
        .badge-danger { background: #e74c3c; color: white; }
        .badge-info { background: #3498db; color: white; }
        .logo { text-align: center; margin-bottom: 30px; }
        .footer { margin-top: 40px; text-align: center; color: #7f8c8d; font-size: 12px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">
            <h1 style="color: #2980b9;">🔍 Forensic Analysis Report</h1>
            <p style="color: #7f8c8d;">Generated on $(Get-Date -Format 'dddd, MMMM dd, yyyy HH:mm:ss')</p>
        </div>
        
        <div class="summary">
            <h2>📊 Executive Summary</h2>
            <p><strong>Analysis Period:</strong> $($TimeRange.RangeText)</p>
            <p><strong>System:</strong> $env:COMPUTERNAME</p>
            <p><strong>Generated By:</strong> $env:USERNAME</p>
            <p><strong>Total Events Analyzed:</strong> $($Data.Count)</p>
        </div>
"@

    # Add each section
    foreach ($section in $Data.Keys) {
        $sectionData = $Data[$section]
        
        $html += @"
        <h2>$section</h2>
        <table>
            <thead>
                <tr>
"@
        # Add headers
        if ($sectionData.Count -gt 0) {
            $headers = $sectionData[0].PSObject.Properties.Name
            foreach ($header in $headers) {
                $html += "<th>$header</th>"
            }
        }
        
        $html += @"
                </tr>
            </thead>
            <tbody>
"@
        # Add rows
        foreach ($row in $sectionData) {
            $html += "<tr>"
            foreach ($property in $row.PSObject.Properties) {
                $value = $property.Value
                $cellClass = ""
                
                # Add color coding based on content
                if ($property.Name -eq "EventID") {
                    if ($value -in @(4625, 4660, 4673)) {
                        $cellClass = "critical"
                    } elseif ($value -in @(4672, 4720)) {
                        $cellClass = "warning"
                    }
                }
                
                # Format timestamp
                if ($property.Name -like "*Time*" -or $property.Name -eq "TimeCreated") {
                    $value = "<span class='timestamp'>$value</span>"
                }
                
                # Format EventID with badge
                if ($property.Name -eq "EventID") {
                    $badgeClass = "badge-info"
                    if ($value -in @(4625, 4660)) { $badgeClass = "badge-danger" }
                    elseif ($value -in @(4672, 4720)) { $badgeClass = "badge-warning" }
                    $value = "<span class='badge $badgeClass'>$value</span>"
                }
                
                $html += "<td class='$cellClass'>$value</td>"
            }
            $html += "</tr>"
        }
        
        $html += @"
            </tbody>
        </table>
"@
    }
    
    $html += @"
        <div class="footer">
            <hr>
            <p>Report generated by LogMaster Pro Forensic Tool</p>
            <p>Confidential - For Investigative Purposes Only</p>
            <p>© $(Get-Date -Format 'yyyy') Digital Forensics Suite</p>
        </div>
    </div>
</body>
</html>
"@
    
    $filePath = "$ExportPath\$FileName.html"
    $html | Out-File -FilePath $filePath -Encoding UTF8
    Write-Host "✅ HTML Report saved: $filePath" -ForegroundColor $ColorSuccess
    
    # Also export as CSV
    foreach ($section in $Data.Keys) {
        $csvPath = "$ExportPath\${FileName}_$section.csv"
        $Data[$section] | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    }
    
    return $filePath
}

#endregion

#region Main Analysis Modules
# ===========================

function Invoke-FileAccessAnalysis {
    param($TimeRange)
    
    Write-Host "`n📂 ANALYZING FILE ACCESS ACTIVITY..." -ForegroundColor $ColorHighlight
    
    $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4663} -ErrorAction SilentlyContinue | 
        Where-Object { $_.TimeCreated -ge $TimeRange.Start -and $_.TimeCreated -le $TimeRange.End }
    
    $results = @()
    foreach ($event in $events) {
        $details = Get-EventDetails -Event $event
        
        $results += [PSCustomObject]@{
            TimeCreated = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            User = $details.User
            Computer = $details.Computer
            FilePath = $details.FilePath
            Process = $details.Process
            AccessMask = $details.AccessMask
            IPAddress = $details.IPAddress
        }
    }
    
    return $results
}

function Invoke-UserActivityAnalysis {
    param($TimeRange)
    
    Write-Host "`n👤 ANALYZING USER ACTIVITY..." -ForegroundColor $ColorHighlight
    
    $logonEvents = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624} -ErrorAction SilentlyContinue | 
        Where-Object { $_.TimeCreated -ge $TimeRange.Start -and $_.TimeCreated -le $TimeRange.End } |
        Select-Object -First 50
    
    $results = @()
    foreach ($event in $logonEvents) {
        $details = Get-EventDetails -Event $event
        
        $results += [PSCustomObject]@{
            TimeCreated = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            User = $details.User
            LogonType = if ($event.Message -match "Logon Type:\s+(\d+)") { $matches[1] } else { "N/A" }
            IPAddress = $details.IPAddress
            Computer = $details.Computer
            Status = $details.Status
        }
    }
    
    return $results
}

function Invoke-USBAnalysis {
    param($TimeRange)
    
    Write-Host "`n💾 ANALYZING USB DEVICE ACTIVITY..." -ForegroundColor $ColorHighlight
    
    $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=6416} -ErrorAction SilentlyContinue | 
        Where-Object { $_.TimeCreated -ge $TimeRange.Start -and $_.TimeCreated -le $TimeRange.End }
    
    $results = @()
    foreach ($event in $events) {
        $details = Get-EventDetails -Event $event
        
        $results += [PSCustomObject]@{
            TimeCreated = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            User = $details.User
            Device = $details.Device
            Computer = $details.Computer
            Action = if ($event.Message -match "was (mounted|dismounted)") { $matches[1] } else { "Connected" }
            Status = $details.Status
        }
    }
    
    return $results
}

function Invoke-ProcessAnalysis {
    param($TimeRange)
    
    Write-Host "`n⚙️ ANALYZING PROCESS CREATION..." -ForegroundColor $ColorHighlight
    
    $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4688} -ErrorAction SilentlyContinue | 
        Where-Object { $_.TimeCreated -ge $TimeRange.Start -and $_.TimeCreated -le $TimeRange.End } |
        Select-Object -First 50
    
    $results = @()
    foreach ($event in $events) {
        $details = Get-EventDetails -Event $event
        
        $results += [PSCustomObject]@{
            TimeCreated = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            User = $details.User
            Process = $details.Process
            ParentProcess = if ($event.Message -match "Parent Process Name:\s+(.+)") { $matches[1] } else { "N/A" }
            CommandLine = if ($event.Message -match "Command Line:\s+(.+)") { $matches[1] } else { "N/A" }
            Computer = $details.Computer
        }
    }
    
    return $results
}

function Invoke-AccountManagementAnalysis {
    param($TimeRange)
    
    Write-Host "`n🔐 ANALYZING ACCOUNT MANAGEMENT..." -ForegroundColor $ColorHighlight
    
    $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4720, 4722, 4725, 4726, 4738, 4798, 4799} -ErrorAction SilentlyContinue | 
        Where-Object { $_.TimeCreated -ge $TimeRange.Start -and $_.TimeCreated -le $TimeRange.End }
    
    $results = @()
    foreach ($event in $events) {
        $details = Get-EventDetails -Event $event
        
        # Determine action based on Event ID
        $action = switch ($event.Id) {
            4720 { "User Account Created" }
            4722 { "User Account Enabled" }
            4725 { "User Account Disabled" }
            4726 { "User Account Deleted" }
            4738 { "User Account Changed" }
            4798 { "User Group Membership Enumerated" }
            4799 { "Failed Group Enumeration" }
            default { "Account Activity" }
        }
        
        $results += [PSCustomObject]@{
            TimeCreated = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            Action = $action
            TargetUser = if ($event.Message -match "Target Account Name:\s+(\S+)") { $matches[1] } else { $details.User }
            PerformedBy = $details.User
            Computer = $details.Computer
            Status = $details.Status
        }
    }
    
    return $results
}

#endregion

#region Main Program
# ==================

try {
    # Get time range
    $TimeRange = Get-TimeRange
    
    Write-Host "`n🔍 STARTING COMPREHENSIVE ANALYSIS..." -ForegroundColor $ColorHeader
    Write-Host "   Time Range: $($TimeRange.RangeText)" -ForegroundColor $ColorInfo
    Write-Host "   From: $($TimeRange.Start)" -ForegroundColor $ColorInfo
    Write-Host "   To: $($TimeRange.End)" -ForegroundColor $ColorInfo
    
    # Initialize results collection
    $AllResults = @{}
    
    # Run all analysis modules
    $AllResults["File_Access_Logs"] = Invoke-FileAccessAnalysis -TimeRange $TimeRange
    $AllResults["User_Activity"] = Invoke-UserActivityAnalysis -TimeRange $TimeRange
    $AllResults["USB_Activity"] = Invoke-USBAnalysis -TimeRange $TimeRange
    $AllResults["Process_Creation"] = Invoke-ProcessAnalysis -TimeRange $TimeRange
    $AllResults["Account_Management"] = Invoke-AccountManagementAnalysis -TimeRange $TimeRange
    
    # Show summary
    Show-Menu -Title "ANALYSIS SUMMARY"
    
    Write-Host "`n📊 ANALYSIS COMPLETED SUCCESSFULLY!" -ForegroundColor $ColorSuccess
    Write-Host "──────────────────────────────────────────────────────" -ForegroundColor $ColorInfo
    
    foreach ($section in $AllResults.Keys) {
        $count = $AllResults[$section].Count
        $color = if ($count -gt 0) { $ColorWarning } else { $ColorInfo }
        Write-Host "   $section".PadRight(25) -NoNewline -ForegroundColor $ColorInfo
        Write-Host ": $count events" -ForegroundColor $color
    }
    
    # Generate reports
    Write-Host "`n📁 GENERATING REPORTS..." -ForegroundColor $ColorHighlight
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportFile = Export-ToHTML -Data $AllResults -Title "Forensic Analysis Report" -FileName "Forensic_Report_$timestamp"
    
    # Generate summary text file
    $summaryText = @"
=============================================
FORENSIC ANALYSIS REPORT - SUMMARY
=============================================
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
System: $env:COMPUTERNAME
Investigator: $env:USERNAME
Analysis Period: $($TimeRange.RangeText)
From: $($TimeRange.Start)
To: $($TimeRange.End)

EVENT SUMMARY:
=============================================
File Access Events: $($AllResults['File_Access_Logs'].Count)
User Activity Events: $($AllResults['User_Activity'].Count)
USB Activity Events: $($AllResults['USB_Activity'].Count)
Process Creation Events: $($AllResults['Process_Creation'].Count)
Account Management Events: $($AllResults['Account_Management'].Count)

CRITICAL FINDINGS:
=============================================
"@
    
    # Add critical findings
    $criticalEvents = $AllResults['File_Access_Logs'] | Where-Object { $_.FilePath -match "secret|confidential|password|sensitive" }
    if ($criticalEvents.Count -gt 0) {
        $summaryText += "`nSensitive File Accesses:`n"
        $criticalEvents | ForEach-Object {
            $summaryText += "  - $($_.TimeCreated): $($_.User) accessed $($_.FilePath)`n"
        }
    }
    
    $failedLogons = $AllResults['User_Activity'] | Where-Object { $_.Status -match "Failure|0xC" }
    if ($failedLogons.Count -gt 0) {
        $summaryText += "`nFailed Logon Attempts:`n"
        $failedLogons | ForEach-Object {
            $summaryText += "  - $($_.TimeCreated): $($_.User) from $($_.IPAddress)`n"
        }
    }
    
    $summaryText += @"

REPORT LOCATIONS:
=============================================
HTML Report: $reportFile
CSV Reports: $ExportPath\Forensic_Report_$timestamp_*.csv
Log File: $LogFile
Desktop Folder: $ExportPath

RECOMMENDATIONS:
=============================================
1. Review File Access logs for unauthorized access
2. Monitor failed login attempts
3. Check USB device connections for unauthorized data transfer
4. Review account management activities
5. Export and preserve logs for evidence

=============================================
INVESTIGATION COMPLETE
=============================================
"@
    
    $summaryFile = "$ExportPath\Investigation_Summary_$timestamp.txt"
    $summaryText | Out-File -FilePath $summaryFile -Encoding UTF8
    
    # Show completion message
    Write-Host "`n" + ("═" * 70) -ForegroundColor $ColorSuccess
    Write-Host "   ✅ ANALYSIS COMPLETE!" -ForegroundColor $ColorSuccess
    Write-Host ("═" * 70) -ForegroundColor $ColorSuccess
    
    Write-Host "`n📂 REPORTS GENERATED:" -ForegroundColor $ColorHeader
    Write-Host "──────────────────────────────────────────────────────" -ForegroundColor $ColorInfo
    Write-Host "HTML Report:    " -NoNewline -ForegroundColor $ColorInfo
    Write-Host "$reportFile" -ForegroundColor $ColorHighlight
    Write-Host "Summary File:   " -NoNewline -ForegroundColor $ColorInfo
    Write-Host "$summaryFile" -ForegroundColor $ColorHighlight
    Write-Host "Desktop Folder: " -NoNewline -ForegroundColor $ColorInfo
    Write-Host "$ExportPath" -ForegroundColor $ColorHighlight
    Write-Host "Log File:       " -NoNewline -ForegroundColor $ColorInfo
    Write-Host "$LogFile" -ForegroundColor $ColorHighlight
    
    # Open the HTML report
    $openReport = Read-Host "`nDo you want to open the HTML report now? (Y/N)"
    if ($openReport -eq 'Y' -or $openReport -eq 'y') {
        Start-Process $reportFile
    }
    
    # Open the desktop folder
    $openFolder = Read-Host "Do you want to open the reports folder? (Y/N)"
    if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
        Start-Process $ExportPath
    }
    
    Write-Host "`n🔐 Log file saved for audit trail: $LogFile" -ForegroundColor $ColorInfo
    
}
catch {
    Write-Host "`n❌ ERROR OCCURRED: $_" -ForegroundColor $ColorError
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor $ColorError
}
finally {
    Stop-Transcript
    Write-Host "`n👋 Analysis session ended. Reports are saved on your Desktop." -ForegroundColor $ColorInfo
    Read-Host "`nPress Enter to exit"
}

#endregion