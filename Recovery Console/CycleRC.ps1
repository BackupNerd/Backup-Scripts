clear-host
Get-Process -Name 'RecoveryConsole' -ErrorAction SilentlyContinue
Stop-Process -Name 'RecoveryConsole' -Force -ErrorAction SilentlyContinue


Get-Service 'Recovery Console Service' -ErrorAction SilentlyContinue
Stop-Service 'Recovery Console Service' -Force -ErrorAction SilentlyContinue
Stop-Process -Name BackupFP -Force -ErrorAction SilentlyContinue
Start-Service 'Recovery Console Service' -ErrorAction SilentlyContinue

Start-Process -FilePath "C:\Program Files\RecoveryConsole\RecoveryConsole.exe" -ArgumentList "service-managed" -ErrorAction SilentlyContinue

#Get-Process -ComputerName savage -Name RecoveryConsole
#Get-Service -ComputerName Savage -name 'Recovery Console Service'
