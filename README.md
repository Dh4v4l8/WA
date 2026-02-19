```
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
```
```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
```
# Run PowerShell as Administrator
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```
```
powershell -ExecutionPolicy Bypass -File "C:\LogAnalyzer.ps1"
```
```
$start = (Get-Date).AddDays(-3)
Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4663; StartTime=$start} |
  Where-Object { $_.Message -like '*E:\*' } |
  Select-Object TimeCreated, Id, Message |
  Out-File .\usb_file_events.txt -Encoding utf8
```

### Windows Registry Artifacts (FOR USB)
```
SYSTEM\CurrentControlSet\Enum\USBSTOR

SYSTEM\MountedDevices

NTUSER.DAT → Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2
```
