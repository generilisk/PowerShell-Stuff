<#
.SYNOPSIS
    Lists domain devices with names matching a specified regex pattern.

.DESCRIPTION
    This script queries Active Directory to find computer objects with names that match a given regex pattern.
    It uses Get-ADComputer to retrieve all devices and filters them based on the regex pattern provided.

.EXAMPLE
    .\Find-DomainDevices.ps1 -RegexPattern "PC[0-9]{3}"
    This example searches for devices with names like 'PC001', 'PC123', etc., where the name starts with 'PC' followed by exactly three digits.

.EXAMPLE
    .\Find-DomainDevices.ps1 -RegexPattern "^SERVER"
    This example searches for devices with names that start with 'SERVER'.
#>

# Script Name: Find-DomainDevices
# Created by: Will Hughes
# Date: 2024-11-14
# Patch Notes: Initial script version

param (
    [string]$RegexPattern = ".*" # Default pattern matches all devices
)

# Import Active Directory module (unnecessary if already imported in your environment)
Import-Module ActiveDirectory

try {
    # Search for devices matching the regex pattern
    $devices = Get-ADComputer -Filter { Name -like "*" } | Where-Object { $_.Name -match $RegexPattern } | Sort-Object Name

    # Check if any devices were found
    if ($devices.Count -eq 0) {
        Write-Output "No devices found matching the pattern '$RegexPattern'."
    }
    else {
        # Display matching devices
        Write-Output "Devices matching the pattern '$RegexPattern':"
        $devices | ForEach-Object { Write-Output $_.Name }
    }
}
catch {
    Write-Error "An error occurred: $_"
}
