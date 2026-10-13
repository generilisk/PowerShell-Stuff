#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Finds the next available computer name following a naming convention in Active Directory and renames this computer to it.

.DESCRIPTION
    Computer names are a letter prefix followed by exactly 3 digits (for example AMEDW001). Digits at the end of
    the -Prefix count toward those 3, which lets you pick the number range:
        -Prefix "AMEDW0"  ->  AMEDW001 to AMEDW099  (for example hallway stations)
        -Prefix "AMEDW1"  ->  AMEDW101 to AMEDW199  (for example patient rooms)
        -Prefix "AMEDW"   ->  AMEDW001 to AMEDW999  (any number)
    The script queries Active Directory for existing computers in that range, finds the lowest unused number,
    and builds the next computer name from it. If no existing computers match, it starts at the first name in the range.
    Before renaming and restarting, the script asks for confirmation (or use -WhatIf to preview without changing anything).
    Needs an elevated PowerShell session. If the Active Directory tools (RSAT) are missing they are installed automatically.

.PARAMETER Prefix
    The letters, plus any leading digits, that the computer name should start with. The name always ends up with
    3 digits in total, so "AMEDW0" fills in 2 more digits and "AMEDW1" does the same. Default is "AMEDW0".

.EXAMPLE
    .\Rename-ComputerWithLookup.ps1 -Prefix "AMEDW0"
    Finds the next available AMEDW0xx name (AMEDW001 to AMEDW099), asks for confirmation, then renames this computer and restarts it.

.EXAMPLE
    .\Rename-ComputerWithLookup.ps1 -Prefix "AMEDW1" -WhatIf
    Shows the next available AMEDW1xx name without renaming or restarting this computer.

.NOTES
    Script Name: Rename-ComputerWithLookup.ps1
    Created by: Will Hughes
    Date: 2024-05-07
    Patch Notes:
        2026-10-13: Fixed param block (Clear-Host before it broke -Prefix), fixed next-number logic and the
                    3-digit name format, added AD error handling, RSAT install, -WhatIf/confirmation,
                    and enabled the rename.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [ValidateNotNullOrEmpty()]
    [string]$Prefix = "AMEDW0"
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

# Names always end in 3 digits. Digits already at the end of the prefix count toward those 3,
# so "AMEDW0" leaves 2 digits to fill in (01-99) and "AMEDW1" does the same (AMEDW101-AMEDW199).
$totalDigits  = 3
$prefixDigits = if ($Prefix -match '\d+$') { $Matches[0].Length } else { 0 }
$numberWidth  = $totalDigits - $prefixDigits
if ($numberWidth -lt 1) {
    Write-Warning "Prefix '$Prefix' already ends in $prefixDigits digits, leaving no room for the 3-digit number. Use a shorter prefix such as 'AMEDW0'."
    return
}
$maxNumber = [int]([math]::Pow(10, $numberWidth) - 1)

# Find existing computers named exactly <Prefix> followed by the remaining digits
try {
    $existingNames = @(Get-ADComputer -Filter "Name -like '$Prefix*'" -ErrorAction Stop |
        Select-Object -ExpandProperty Name)
}
catch {
    Write-Warning "Could not query Active Directory: $($_.Exception.Message)"
    return
}

$pattern = '^' + [regex]::Escape($Prefix) + '(\d{' + $numberWidth + '})$'
$numbers = @(foreach ($name in $existingNames) {
    if ($name -match $pattern) { [int]$Matches[1] }
})
Write-Verbose "Found $($numbers.Count) existing computers matching '$Prefix'"

# Find the lowest unused number (starting at 1)
$nextNumber = 1
while ($numbers -contains $nextNumber) {
    $nextNumber++
}
if ($nextNumber -gt $maxNumber) {
    Write-Warning "All $maxNumber names for prefix '$Prefix' are already in use."
    return
}

# Format the next computer name
$nextComputerName = $Prefix + $nextNumber.ToString("D$numberWidth")
Write-Host "Next computer name: $nextComputerName"

# Rename the computer and restart (asks for confirmation; -WhatIf previews only)
if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Rename to $nextComputerName and restart")) {
    try {
        Rename-Computer -NewName $nextComputerName -Force -Restart -ErrorAction Stop
    }
    catch {
        Write-Warning "Rename failed: $($_.Exception.Message)"
    }
}