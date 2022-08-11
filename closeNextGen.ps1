$User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$NGParts = ("NGAppLauncher","NextGenICS","qsi_cps","NextGenEMR","NGFileMaint","NGLicenseManager","NextGenEPM","NextGenEPMRptSvr","RosettaHoldingTank","NGSetDB","NGSystemAdmin","TemplateEditor","TemplateImportExport")
Get-Process -Name $NGParts -IncludeUserName | Where-Object {$User -eq $_.username} | Stop-Process