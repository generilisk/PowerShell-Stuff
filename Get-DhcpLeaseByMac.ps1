<#
.SYNOPSIS
    Searches all DHCP scopes on the specified DHCP server for an IP address assigned to a specified MAC address.
.DESCRIPTION
    This script connects to the DHCP server and retrieves the IP address associated with a specified MAC address. 
    It searches across all DHCP scopes on the server until it finds a match.
.EXAMPLE
    .\Get-DhcpLeaseByMac.ps1 -MacAddress "00-11-22-33-44-55"
    Searches the DHCP server for the IP address assigned to the MAC address 00-11-22-33-44-55.
.EXAMPLE
    .\Get-DhcpLeaseByMac.ps1 -MacAddress "AA-BB-CC-DD-EE-FF"
    Searches the DHCP server for the IP address assigned to the MAC address AA-BB-CC-DD-EE-FF.
#>

# Script Name: Get-DhcpLeaseByMac.ps1
# Created by: Will Hughes
# Date: 2024-11-14
# Patch Notes: Initial script creation.

param (
    [string]$MacAddress  # Target MAC address
)

$DhcpServer = "NTDHCP001"  # DHCP server name
$found = $false

# Retrieve all scopes from the DHCP server
$scopes = Get-DhcpServerv4Scope -ComputerName $DhcpServer

# Loop through each scope to find a lease with the specified MAC address
foreach ($scope in $scopes) {
    $lease = Get-DhcpServerv4Lease -ComputerName $DhcpServer -ScopeId $scope.ScopeId -ClientId $MacAddress -ErrorAction SilentlyContinue
    if ($lease) {
        Write-Output "IP Address for MAC $MacAddress in Scope $($scope.ScopeId) is: $($lease.IPAddress)"
        $found = $true
        break
    }
}

if (-not $found) {
    Write-Output "No lease found for MAC address $MacAddress on server $DhcpServer."
}