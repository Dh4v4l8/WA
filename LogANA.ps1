# ======================================================================
# LOGMASTER PRO - ADVANCED WINDOWS FORENSIC ANALYSIS TOOL
# Version: 2.0 (Fixed Version)
# ======================================================================

# Clear screen and show banner
Clear-Host

# ASCII Art Banner (FIXED)
Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "    WINDOWS FORENSIC ANALYSIS SUITE v2.0" -ForegroundColor Yellow
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "          Advanced Log Analysis & Reporting" -ForegroundColor White
Write-Host ""

# Check Administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script requires Administrator privileges!" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
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
}

# Time periods for analysis
$TimePeriods = @{
    "Last 1 Hour" = (Get-Date).AddHours(-1)
    "Last 4 Hours" = (Get-Date).AddHours(-4)
    "Last 24 Hours" = (Get-Date).AddDays(-1)
    "Last 7 Days" = (Get-Date).AddDays(-7)
    "Custom Date Range" = $null
}

# Function to show menu
function Show-Menu {
    param([string]$Title = "MAIN MENU")
    
    Clear-Host
    Write-Host "`n=====================================================" -ForegroundColor $ColorHeader
    Write-Host "   $Title" -ForegroundColor $ColorHeader
    Write-Host "   [$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]" -ForegroundColor $ColorInfo
    Write-Host "=====================================================" -ForegroundColor $ColorHeader
}

# Function to get time range
function Get-TimeRange {
    Show-Menu -Title "TIME RANGE SELECTION"
    
    Write-Host "`nSELECT TIME RANGE FOR ANALYSIS:" -ForegroundColor $ColorHighlight
    Write-Host "-----------------------------------------------------" -ForegroundColor $ColorInfo
    
    $i = 1
    foreach ($period in $TimePeriods.Keys) {
        Write-Host "  $i. $period" -ForegroundColor $ColorInfo
        $i++
    }
    
    Write-Host "`n  X. Exit Program" -ForegroundColor $ColorError
    Write-Host "-----------------------------------------------------" -ForegroundColor $ColorInfo
    
    $choice = Read-Host "`nEnter your choice (1-$($i-1))"
    
    if ($choice -eq 'X' -or $choice -eq 'x') {
        Write-Host "Exiting program..." -ForegroundColor $ColorWarning
        Stop-Transcript
        exit
    }
    
    if ($choice -eq "5") {
        # Custom date range
        Write-Host "`nENTER CUSTOM DATE RANGE:" -ForegroundColor $ColorHighlight
        
        # Start Date
        do {
            $startDate = Read-Host "Start Date (DD-MM-YYYY)"
            try {
                $StartDate = Get-Date $startDate
                Write-Host "✓ Start Date: $($StartDate.ToString('dd MMMM yyyy'))" -ForegroundColor Green
                $validStart = $true
            }
            catch {
                Write-Host "Invalid date format! Example: 1-12-2025 or 01-12-2025" -ForegroundColor Red
                $validStart = $false
            }
        } while (-not $validStart)
        
        # Start Time
        $startTime = Read-Host "Start Time (HH:MM) [Default: 00:00]"
        if ([string]::IsNullOrWhiteSpace($startTime)) {
            $startTime = "00:00"
        }
        
        # End Date
        do {
            $endDate = Read-Host "End Date (DD-MM-YYYY)"
            try {
                $EndDate = Get-Date $endDate
                Write-Host "✓ End Date: $($EndDate.ToString('dd MMMM yyyy'))" -ForegroundColor Green
                $validEnd = $true
            }
            catch {
                Write-Host "Invalid date format! Example: 3-2-2026 or 03-02-2026" -ForegroundColor Red
                $validEnd = $false
            }
        } while (-not $validEnd)
        
        # End Time
        $endTime = Read-Host "End Time (HH:MM) [Default: 23:59]"
        if ([string]::IsNullOrWhiteSpace($endTime)) {
            $endTime = "23:59"
        }
        
        $StartTime = Get-Date "$($StartDate.ToString('yyyy-MM-dd')) $startTime"
        $EndTime = Get-Date "$($EndDate.ToString('yyyy-MM-dd')) $endTime"
        $selectedPeriod = "Custom: $($StartTime.ToString('dd-MM-yyyy HH:mm')) to $($EndTime.ToString('dd-MM-yyyy HH:mm'))"
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
        RangeText = $selectedPeriod
    }
}

# Function to get event details
function Get-EventDetails {
    param($Event)
    
    $details = @{}
    $message = $Event.Message
    
    # Parse patterns
    if ($message -match "Account Name:\s+(\S+)") {
        $details["User"] = $matches[1]
    } elseif ($message -match "Subject:.+?Account Name:\s+(\S+)") {
        $details["User"] = $matches[1]
    } else {
        $details["User"] = "N/A"
    }
    
    if ($message -match "Object Name:\s+(.+)") {
        $details["FilePath"] = $matches[1]
    } else {
        $details["FilePath"] = "N/A"
    }
    
    if ($message -match "Process Name:\s+(.+)") {
        $details["Process"] = $matches[1]
    } else {
        $details["Process"] = "N/A"
    }
    
    if ($message -match "Device Description:\s+(.+)") {
        $details["Device"] = $matches[1]
    } else {
        $details["Device"] = "N/A"
    }
    
    if ($message -match "Computer:\s+(\S+)") {
        $details["Computer"] = $matches[1]
    } else {
        $details["Computer"] = $env:COMPUTERNAME
    }
    
    return $details
}

# Function to export to HTML (FIXED VERSION)
function Export-ToHTML {
    param($Data, $Title, $FileName)
    
    # Build HTML step by step to avoid parsing errors
    $html = "<!DOCTYPE html>`n"
    $html += "<html>`n"
    $html += "<head>`n"
    $html += "    <meta charset=`"UTF-8`">`n"
    $html += "    <title>$Title</title>`n"
    $html += "    <style>`n"
    $html += "        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; background: #f5f5f5; }`n"
    $html += "        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }`n"
    $html += "        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }`n"
    $html += "        h2 { color: #34495e; margin-top: 30px; }`n"
    $html += "        .summary { background: #ecf0f1; padding: 20px; border-radius: 5px; margin: 20px 0; }`n"
    $html += "        table { width: 100%; border-collapse: collapse; margin: 20px 0; }`n"
    $html += "        th { background: #3498db; color: white; padding: 12px; text-align: left; }`n"
    $html += "        td { padding: 10px; border-bottom: 1px solid #ddd; }`n"
    $html += "        tr:hover { background: #f9f9f9; }`n"
    $html += "        .timestamp { font-family: 'Consolas', monospace; }`n"
    $html += "        .footer { margin-top: 40px; text-align: center; color: #7f8c8d; font-size: 12px; }`n"
    $html += "    </style>`n"
    $html += "</head>`n"
    $html += "<body>`n"
    $html += "    <div class=`"container`">`n"
    $html += "        <div class=`"logo`">`n"
    $html += "            <h1 style=`"color: #2980b9;`">Forensic Analysis Report</h1>`n"
    $html += "            <p style=`"color: #7f8c8d;`">Generated on $(Get-Date -Format 'dddd, MMMM dd, yyyy HH:mm:ss')</p>`n"
    $html += "        </div>`n"
    $html += "        `n"
    $html += "        <div class=`"summary`">`n"
    $html += "            <h2>Executive Summary</h2>`n"
    $html += "            <p><strong>Analysis Period:</strong> $($TimeRange.RangeText)</p>`n"
    $html += "            <p><strong>System:</strong> $env:COMPUTERNAME</p>`n"
    $html += "            <p><strong>Generated By:</strong> $env:USERNAME</p>`n"
    $html += "            <p><strong>Total Events Analyzed:</strong> $(($Data.Values | Measure-Object -Property Count -Sum).Sum)</p>`n"
    $html += "        </div>`n"
    
    # Add each section
    foreach ($section in $Data.Keys) {
        $sectionData = $Data[$section]
        
        if ($sectionData.Count -gt 0) {
            $html += "        <h2>$section</h2>`n"
            $html += "        <table>`n"
            $html += "            <thead>`n"
            $html += "                <tr>`n"
            
            # Add headers
            $headers = $sectionData[0].PSObject.Properties.Name
            foreach ($header in $headers) {
                $html += "                    <th>$header</th>`n"
            }
            
            $html += "                </tr>`n"
            $html += "            </thead>`n"
            $html += "            <tbody>`n"
            
            # Add rows
            foreach ($row in $sectionData) {
                $html += "                <tr>`n"
                foreach ($property in $row.PSObject.Properties) {
                    $value = $property.Value
                    $html += "                    <td>$value</td>`n"
                }
                $html += "                </tr>`n"
            }
            
            $html += "            </tbody>`n"
            $html += "        </table>`n"
        }
    }
    
    $html += "        <div class=`"footer`">`n"
    $html += "            <hr>`n"
    $html += "            <p>Report generated by LogMaster Pro Forensic Tool</p>`n"
    $html += "            <p>Confidential - For Investigative Purposes Only</p>`n"
    $html += "            <p>© $(Get-Date -Format 'yyyy') Digital Forensics Suite</p>`n"
    $html += "        </div>`n"
    $html += "    </div>`n"
    $html += "</body>`n"
    $html += "</html>"
    
    $filePath = "$ExportPath\$FileName.html"
    $html | Out-File -FilePath $filePath -Encoding UTF8
    Write-Host "HTML Report saved: $filePath" -ForegroundColor Green
    
    # Also export as CSV
    foreach ($section in $Data.Keys) {
        $csvPath = "$ExportPath\${FileName}_$section.csv"
        $Data[$section] | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    }
    
    return $filePath
}

# Analysis modules
function Invoke-FileAccessAnalysis {
    param($TimeRange)
    
    Write-Host "`nANALYZING FILE ACCESS ACTIVITY..." -ForegroundColor $ColorHighlight
    
    $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4663} -ErrorAction SilentlyContinue | 
        Where-Object { $_.TimeCreated -ge $TimeRange.Start -and $_.TimeCreated -le $TimeRange.End }
    
    $results = @()
    foreach ($event in $events) {
        $details = Get-EventDetails -Event $event
        
        $results += [PSCustomObject]@{
            Time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            User = $details.User
            FilePath = $details.FilePath
            Process = $details.Process
            Computer = $details.Computer
        }
    }
    
    return $results
}

function Invoke-UserActivityAnalysis {
    param($TimeRange)
    
    Write-Host "`nANALYZING USER ACTIVITY..." -ForegroundColor $ColorHighlight
    
    $logonEvents = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624} -ErrorAction SilentlyContinue | 
        Where-Object { $_.TimeCreated -ge $TimeRange.Start -and $_.TimeCreated -le $TimeRange.End } |
        Select-Object -First 50
    
    $results = @()
    foreach ($event in $logonEvents) {
        $details = Get-EventDetails -Event $event
        
        $logonType = "N/A"
        if ($event.Message -match "Logon Type:\s+(\d+)") {
            $logonType = $matches[1]
        }
        
        $results += [PSCustomObject]@{
            Time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            User = $details.User
            LogonType = $logonType
            Computer = $details.Computer
        }
    }
    
    return $results
}

function Invoke-USBAnalysis {
    param($TimeRange)
    
    Write-Host "`nANALYZING USB DEVICE ACTIVITY..." -ForegroundColor $ColorHighlight
    
    $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=6416} -ErrorAction SilentlyContinue | 
        Where-Object { $_.TimeCreated -ge $TimeRange.Start -and $_.TimeCreated -le $TimeRange.End }
    
    $results = @()
    foreach ($event in $events) {
        $details = Get-EventDetails -Event $event
        
        $results += [PSCustomObject]@{
            Time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            User = $details.User
            Device = $details.Device
            Computer = $details.Computer
        }
    }
    
    return $results
}

function Invoke-ProcessAnalysis {
    param($TimeRange)
    
    Write-Host "`nANALYZING PROCESS CREATION..." -ForegroundColor $ColorHighlight
    
    $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4688} -ErrorAction SilentlyContinue | 
        Where-Object { $_.TimeCreated -ge $TimeRange.Start -and $_.TimeCreated -le $TimeRange.End } |
        Select-Object -First 50
    
    $results = @()
    foreach ($event in $events) {
        $details = Get-EventDetails -Event $event
        
        $parentProcess = "N/A"
        if ($event.Message -match "Parent Process Name:\s+(.+)") {
            $parentProcess = $matches[1]
        }
        
        $results += [PSCustomObject]@{
            Time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            User = $details.User
            Process = $details.Process
            ParentProcess = $parentProcess
            Computer = $details.Computer
        }
    }
    
    return $results
}

function Invoke-AccountManagementAnalysis {
    param($TimeRange)
    
    Write-Host "`nANALYZING ACCOUNT MANAGEMENT..." -ForegroundColor $ColorHighlight
    
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
        
        $targetUser = "N/A"
        if ($event.Message -match "Target Account Name:\s+(\S+)") {
            $targetUser = $matches[1]
        }
        
        $results += [PSCustomObject]@{
            Time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            EventID = $event.Id
            Action = $action
            TargetUser = $targetUser
            PerformedBy = $details.User
            Computer = $details.Computer
        }
    }
    
    return $results
}

# Main Program
try {
    # Get time range
    $TimeRange = Get-TimeRange
    
    Write-Host "`nSTARTING COMPREHENSIVE ANALYSIS..." -ForegroundColor $ColorHeader
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
    
    Write-Host "`nANALYSIS COMPLETED SUCCESSFULLY!" -ForegroundColor $ColorSuccess
    Write-Host "-----------------------------------------------------" -ForegroundColor $ColorInfo
    
    $totalEvents = 0
    foreach ($section in $AllResults.Keys) {
        $count = $AllResults[$section].Count
        $totalEvents += $count
        $color = if ($count -gt 0) { $ColorWarning } else { $ColorInfo }
        Write-Host "   $section".PadRight(25) -NoNewline -ForegroundColor $ColorInfo
        Write-Host ": $count events" -ForegroundColor $color
    }
    
    Write-Host "`n   Total Events: $totalEvents" -ForegroundColor $ColorHighlight
    
    # Generate reports
    Write-Host "`nGENERATING REPORTS..." -ForegroundColor $ColorHighlight
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportFile = Export-ToHTML -Data $AllResults -Title "Forensic Analysis Report" -FileName "Forensic_Report_$timestamp"
    
    # Generate summary text file
    $summaryText = "=============================================`n"
    $summaryText += "FORENSIC ANALYSIS REPORT - SUMMARY`n"
    $summaryText += "=============================================`n"
    $summaryText += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $summaryText += "System: $env:COMPUTERNAME`n"
    $summaryText += "Investigator: $env:USERNAME`n"
    $summaryText += "Analysis Period: $($TimeRange.RangeText)`n"
    $summaryText += "From: $($TimeRange.Start)`n"
    $summaryText += "To: $($TimeRange.End)`n"
    $summaryText += "`n"
    $summaryText += "EVENT SUMMARY:`n"
    $summaryText += "=============================================`n"
    $summaryText += "File Access Events: $($AllResults['File_Access_Logs'].Count)`n"
    $summaryText += "User Activity Events: $($AllResults['User_Activity'].Count)`n"
    $summaryText += "USB Activity Events: $($AllResults['USB_Activity'].Count)`n"
    $summaryText += "Process Creation Events: $($AllResults['Process_Creation'].Count)`n"
    $summaryText += "Account Management Events: $($AllResults['Account_Management'].Count)`n"
    $summaryText += "Total Events: $totalEvents`n"
    $summaryText += "`n"
    $summaryText += "REPORT LOCATIONS:`n"
    $summaryText += "=============================================`n"
    $summaryText += "HTML Report: $reportFile`n"
    $summaryText += "CSV Reports: $ExportPath\Forensic_Report_$timestamp_*.csv`n"
    $summaryText += "Log File: $LogFile`n"
    $summaryText += "Desktop Folder: $ExportPath`n"
    $summaryText += "`n"
    $summaryText += "RECOMMENDATIONS:`n"
    $summaryText += "=============================================`n"
    $summaryText += "1. Review File Access logs for unauthorized access`n"
    $summaryText += "2. Monitor failed login attempts`n"
    $summaryText += "3. Check USB device connections for unauthorized data transfer`n"
    $summaryText += "4. Review account management activities`n"
    $summaryText += "5. Export and preserve logs for evidence`n"
    $summaryText += "`n"
    $summaryText += "=============================================`n"
    $summaryText += "INVESTIGATION COMPLETE`n"
    $summaryText += "=============================================`n"
    
    $summaryFile = "$ExportPath\Investigation_Summary_$timestamp.txt"
    $summaryText | Out-File -FilePath $summaryFile -Encoding UTF8
    
    # Show completion message
    Write-Host "`n=====================================================" -ForegroundColor $ColorSuccess
    Write-Host "   ANALYSIS COMPLETE!" -ForegroundColor $ColorSuccess
    Write-Host "=====================================================" -ForegroundColor $ColorSuccess
    
    Write-Host "`nREPORTS GENERATED:" -ForegroundColor $ColorHeader
    Write-Host "-----------------------------------------------------" -ForegroundColor $ColorInfo
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
    
    Write-Host "`nLog file saved for audit trail: $LogFile" -ForegroundColor $ColorInfo
}
catch {
    Write-Host "`nERROR OCCURRED: $_" -ForegroundColor $ColorError
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor $ColorError
}
finally {
    Stop-Transcript
    Write-Host "`nAnalysis session ended. Reports are saved on your Desktop." -ForegroundColor $ColorInfo
    Read-Host "Press Enter to exit"
}
