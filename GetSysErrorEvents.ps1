<#
.Synopsis
    Return the last errors in the System Event Log
.DESCRIPTION
    Long Description
.EXAMPLE
    Example of how to use this cmdlet
.EXAMPLE
    Another example of how to use this cmdlet
#>

#Script Name: GetSysErrorEvents
#Created by Will Hughes
#Date: 2024-02-16
#Patch Notes:

Param(
    [string]$Log = "System",
    [string]$computerName = $env:COMPUTERNAME,
    [int32]$Newest = 500,
    [string]$ReportTitle = "Event Log Report",
    [Parameter(Mandatory, HelpMessage = "Enter the path for the HTML file.")]
    [string]$Path
)

$data = Get-EventLog -LogName $Log -EntryType Error -Newest 500 -ComputerName $computerName |
Group-Object -Property Source -NoElement
    

$footer = "<h5><i>report run $(Get-Date)</i></h5>"
$css = "https://jdhitsolutions.com/sample.css"
$precontent = "<H1>$computerName</H1><H2>Last $newest error sources from $Log</H2>"

$data | Sort-Object -Property Count, Name -Descending |
Select-Object Count, Name |
ConvertTo-Html -Title $ReportTitle -PreContent $precontent -PostContent $footer -CssUri $css |
Out-File -FilePath $Path