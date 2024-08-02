<#
.Synopsis
    Script to add printers for a user
.DESCRIPTION
    This script prompts the user to enter the names of the full-page and label printers they wish to add. 
    It also asks if the full-page printer should be set as the default printer. 
    The script then adds the specified printers to the user's system and sets the default printer if requested.
.EXAMPLE
    .\Add-Printers.ps1
    Prompts the user for the names of the full-page and label printers, adds them, and optionally sets the full-page printer as default.
.EXAMPLE
    .\Add-Printers.ps1
    User enters "OfficePrinter" for the full-page printer, chooses to set it as default, and enters "LabelPrinter" for the label printer. 
    The script adds both printers and sets "OfficePrinter" as the default printer.
#>

#Script Name: Add-Printers.ps1
#Created by Will Hughes
#Date: 2024-07-17
#Patch Notes:

# Function to add a printer
function Add-PrinterConnection {
    param (
        [string]$printerName
    )
    Add-Printer -ConnectionName "\\prtsvr2\$printerName"
}

# Prompt for the full page printer name
$printerName = Read-Host "Please Type Printer Name (Full Page Printer)"
Add-PrinterConnection -printerName $printerName

# Ask if the full page printer should be set as default
$defaultPrinterResponse = Read-Host "Do you want to set the full page printer as default? (y/n)"

# Check the response and set the printer as default if the response is 'y' or 'Y'
if ($defaultPrinterResponse.Substring(0,1).ToLower() -eq 'y') {
    Set-DefaultPrinter -printerName $printerName
    Write-Output "The printer '$printerName' has been set as the default printer."
} elseif ($defaultPrinterResponse.Substring(0,1).ToLower() -eq 'n') {
    Write-Output "The printer '$printerName' has not been set as the default printer."
} else {
    Write-Output "Invalid input. The printer '$printerName' has not been set as the default printer."
}

# Prompt for the label printer name
$labelPrinterName = Read-Host "Please Type Label Printer Name"
Add-PrinterConnection -printerName $labelPrinterName