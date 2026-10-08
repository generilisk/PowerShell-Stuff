<#
.Synopsis
    Returns the most recent error-level events from a Windows event log as an HTML report.
.DESCRIPTION
    Queries the specified event log (default: System) on the specified computer
    (default: local machine) for the most recent Error-level entries, groups them
    by source, and writes the results out as an HTML report to the given path.
.EXAMPLE
    PS C:\> .\Get-SysErrorEvents.ps1 -Path C:\Reports\SysErrors.html

    Pulls the last 500 System log errors from the local machine and writes the
    report to C:\Reports\SysErrors.html using the default title.
.EXAMPLE
    PS C:\> .\Get-SysErrorEvents.ps1 -Log Application -ComputerName rdbroker1 -Newest 100 -ReportTitle "RD Broker App Errors" -Path C:\Reports\AppErrors.html

    Pulls the last 100 Application log errors from rdbroker1 and writes a
    custom-titled report to the given path.
#>

#Script Name: Get-SysErrorEvents.ps1
#Date: 2024-02-16
#Patch Notes:
#2026-10-08 - Filled in placeholder doc comments. Fixed -Newest parameter,
#             which was declared but never used -- Get-EventLog was always
#             hardcoded to -Newest 500 regardless of what was passed in.
#             Renamed -computerName to -ComputerName for PascalCase
#             consistency (case-insensitive binding, so this doesn't break
#             any existing caller).

Param(
    [string]$Log = "System",
    [string]$ComputerName = $env:COMPUTERNAME,
    [int32]$Newest = 500,
    [string]$ReportTitle = "Event Log Report",
    [Parameter(Mandatory, HelpMessage = "Enter the path for the HTML file.")]
    [string]$Path
)

$data = Get-EventLog -LogName $Log -EntryType Error -Newest $Newest -ComputerName $ComputerName |
Group-Object -Property Source -NoElement


$footer = "<h5><i>report run $(Get-Date)</i></h5>"
$css = "https://jdhitsolutions.com/sample.css"
$precontent = "<H1>$ComputerName</H1><H2>Last $Newest error sources from $Log</H2>"

$data | Sort-Object -Property Count, Name -Descending |
Select-Object Count, Name |
ConvertTo-Html -Title $ReportTitle -PreContent $precontent -PostContent $footer -CssUri $css |
Out-File -FilePath $Path