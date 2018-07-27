### Please note that You will need to enable the PSRemoting Service by running Enable-PSRemoting on machine in Naverisk - Powershell
### Get computers ###

$Computer = Read-Host "Enter Computer Name"

$Computers = (Get-ADComputer $Computer).Name

set /p id=$Computers

### Copy Delprofv2 to machines ###

Get-Service -Name RemoteRegistry -ComputerName $Computers | Start-service

robocopy C:\Temp\Delprof2 \\$Computers\c$\Temp\Delprof2\ /r:0 /w:0

### Start WinRM Service on remote machine ###

Get-Service -Name WinRM -ComputerName $Computers | Start-service

### Launch Delprofv2 to delete profileas onder than 180 days using NTUSER.DAT###
### When determining profile age for /d, use the file NTUSER.INI
#### instead of NTUSER.DAT for age calculation 
 
Invoke-Command -ComputerName $Computers -ScriptBlock { c:\Temp\Delprof2\DelProf2.exe /q /ed:admin* /ed:deep* /d:180 /i}

### When determining profile age for /d, use the file NTUSER.INI
### instead of NTUSER.DAT for age calculation

Invoke-Command -ComputerName $Computers -ScriptBlock { c:\Temp\DelProf2.exe /l /ed:admin* /ed:deep* /ntuserini}