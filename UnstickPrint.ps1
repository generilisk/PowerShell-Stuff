Stop-Service Spooler
Get-Process -Name printfilterpipelinesvc | Stop-Process
Remove-Item %systemroot%\System32\spool\printers\* -Recurse -Force
Start-Service Spooler
