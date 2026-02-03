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