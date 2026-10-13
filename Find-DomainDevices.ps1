<#
.SYNOPSIS
    Lists domain devices with names matching a specified regex pattern.

.DESCRIPTION
    This script queries Active Directory to find computer objects with names that match a given regex pattern.
    It uses Get-ADComputer to retrieve all devices and filters them on the regex pattern provided.
    Matching is case-insensitive and is not anchored, so use ^ and $ if you want to match the whole name.
    The matching computer names are written to the pipeline, so the output can be piped into other commands.
    If the Active Directory tools (RSAT) are missing they are installed automatically (needs an elevated session).

.PARAMETER RegexPattern
    The regex pattern computer names are matched against. Default is ".*", which matches all devices.

.EXAMPLE
    .\Find-DomainDevices.ps1 -RegexPattern "^PC[0-9]{3}$"
    Finds devices named 'PC' followed by exactly three digits, such as 'PC001' or 'PC123'.

.EXAMPLE
    .\Find-DomainDevices.ps1 -RegexPattern "^SERVER"
    Finds devices with names that start with 'SERVER'.

.NOTES
    Script Name: Find-DomainDevices.ps1
    Created by: Will Hughes
    Date: 2024-11-14
    Patch Notes:
        2024-11-14: Initial script version
        2026-10-13: Added RSAT install, error handling for AD queries and bad patterns, fixed example pattern,
                    and moved results to the pipeline (status messages no longer mix in with the names).
#>

[CmdletBinding()]
param (
    [ValidateNotNullOrEmpty()]
    [string]$RegexPattern = ".*" # Default pattern matches all devices
)

# Active Directory PowerShell module (RSAT). This is a Windows feature, so adding it
# needs an elevated session; if that's missing we stop with clear instructions.
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Error ("The Active Directory PowerShell module (RSAT) is not installed. Run this script once from an " +
            "elevated PowerShell window (Run as administrator) and it will install it, or install it yourself with: " +
            "Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'")
        return
    }
    Write-Host "Installing RSAT Active Directory tools (this can take a few minutes)..." -ForegroundColor Cyan
    try {
        Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -WhatIf:$false -Confirm:$false -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Error "Failed to install RSAT Active Directory tools: $_"
        return
    }
}

# Make sure the pattern is a valid regex before querying AD
try {
    $null = [regex]::new($RegexPattern)
}
catch {
    Write-Warning "'$RegexPattern' is not a valid regex pattern: $($_.Exception.Message)"
    return
}

try {
    # Search for devices matching the regex pattern
    $devices = @(Get-ADComputer -Filter "Name -like '*'" -ErrorAction Stop |
        Where-Object { $_.Name -match $RegexPattern } |
        Sort-Object Name)
}
catch {
    Write-Warning "Could not query Active Directory: $($_.Exception.Message)"
    return
}

# Report what was found (status messages go to the console, names go to the pipeline)
if ($devices.Count -eq 0) {
    Write-Host "No devices found matching the pattern '$RegexPattern'."
}
else {
    Write-Host "$($devices.Count) devices matching the pattern '$RegexPattern':"
    $devices | Select-Object -ExpandProperty Name
}