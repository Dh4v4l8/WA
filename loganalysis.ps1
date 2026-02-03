# =============================================
# SPECIFIC DATE TIME RANGE LOG ANALYSIS SCRIPT
# =============================================

# Step 1: Define your date-time range
$StartTime = Get-Date "2024-01-15 09:00:00"  # Start date-time
$EndTime = Get-Date "2024-01-15 17:00:00"    # End date-time

# Step 2: Filter events by time range
$FilteredEvents = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4663} | 
    Where-Object { 
        $_.TimeCreated -ge $StartTime -and 
        $_.TimeCreated -le $EndTime 
    }

# Step 3: Display results in readable format
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "FILE ACCESS LOGS ANALYSIS" -ForegroundColor Yellow
Write-Host "Time Range: $StartTime to $EndTime" -ForegroundColor Green
Write-Host "Total Events Found: $($FilteredEvents.Count)" -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Cyan

# Display each event
$FilteredEvents | ForEach-Object {
    Write-Host "`n[EVENT] Time: $($_.TimeCreated)" -ForegroundColor White
    Write-Host "------------------------------------------------" -ForegroundColor Gray
    
    # Parse and display important details from Message
    $message = $_.Message
    
    # Extract Subject (User who accessed)
    if ($message -match "Subject:.+?Account Name:\s+(\S+)") {
        Write-Host "User: $($matches[1])" -ForegroundColor Cyan
    }
    
    # Extract File Path
    if ($message -match "Object Name:.+?:\s+(.+)") {
        Write-Host "File: $($matches[1])" -ForegroundColor Yellow
    }
    
    # Extract Process
    if ($message -match "Process Name:.+?:\s+(.+)") {
        Write-Host "Process: $($matches[1])" -ForegroundColor Magenta
    }
    
    # Extract Access Type
    if ($message -match "Accesses:.+?:\s+(.+)") {
        Write-Host "Access Type: $($matches[1])" -ForegroundColor Green
    }
}
