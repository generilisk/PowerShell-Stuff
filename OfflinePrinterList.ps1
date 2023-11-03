$printServer = "prtsvr2"

# Connect to the remote print server
$printServerConnection = Get-CimInstance -ClassName Win32_Printer -ComputerName $printServer -ErrorAction SilentlyContinue

if ($printServerConnection) {
    # Filter printers in offline status or with an error status
    $offlinePrinters = $printServerConnection | Where-Object { $_.PrinterStatus -eq 4 -or $_.PrinterStatus -eq 8 }

    if ($offlinePrinters) {
        Write-Host "Offline Printers on $($printServer):"
        $offlinePrinters | Select-Object Name
    }
    else {
        Write-Host "No offline printers found on $($printServer)."
    }
}
else {
    Write-Host "Failed to connect to the print server: $($printServer)"
}
