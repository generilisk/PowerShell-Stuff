<#
.SYNOPSIS
    Lists the user accounts that are members of an Active Directory group.

.DESCRIPTION
    Looks up an Active Directory group by name and lists every user in it, including users who are
    members through nested groups. Computer accounts and the nested groups themselves are not listed.
    The group can be given by its name or its SamAccountName. If no group name is provided as a
    parameter, the script prompts for one.
    The users are written to the pipeline (Name and SamAccountName), so the output can be piped into
    other commands; status messages go to the console.
    If the Active Directory tools (RSAT) are missing they are installed automatically
    (needs an elevated session).

.PARAMETER GroupName
    The name or SamAccountName of the AD group. If omitted, the script prompts for it.

.EXAMPLE
    .\Get-GroupUsers.ps1 -GroupName "Justice League"
    Lists all users in the "Justice League" group, including users in groups nested inside it.

.EXAMPLE
    .\Get-GroupUsers.ps1 -GroupName "Justice League" | Export-Csv .\JusticeLeague.csv -NoTypeInformation
    Saves the list of users to a CSV file.

.EXAMPLE
    .\Get-GroupUsers.ps1
    Prompts for the group name, then lists its users.

.NOTES
    Script Name: Get-GroupUsers.ps1
    Created by: Will Hughes
    Patch Notes:
        2026-10-15: Added help, RSAT install, blank-input guard, and group lookup by name or SamAccountName.
                    Now lists users only (computers are left out) with SamAccountName, and reports
                    "group not found" separately from other AD errors.
#>

[CmdletBinding()]
param (
    [string]$GroupName
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

# Prompt for the AD group name if it wasn't provided
if ([string]::IsNullOrWhiteSpace($GroupName)) {
    $GroupName = Read-Host "Enter the Active Directory group name"
}
if ([string]::IsNullOrWhiteSpace($GroupName)) {
    Write-Warning "No group name entered."
    return
}

try {
    # Find the group by name or SamAccountName (single quotes doubled so names like O'Brien Staff work)
    $safeName = $GroupName -replace "'", "''"
    $groups = @(Get-ADGroup -Filter "Name -eq '$safeName' -or SamAccountName -eq '$safeName'" -ErrorAction Stop)

    if ($groups.Count -eq 0) {
        Write-Host "Group '$GroupName' not found." -ForegroundColor Red
        return
    }
    if ($groups.Count -gt 1) {
        Write-Warning "More than one group matches '$GroupName'. Use the SamAccountName instead: $(($groups | ForEach-Object SamAccountName) -join ', ')"
        return
    }

    # Get users in the group, including members of nested groups (computers are left out)
    $groupMembers = @(Get-ADGroupMember -Identity $groups[0] -Recursive -ErrorAction Stop |
        Where-Object { $_.objectClass -eq 'user' } |
        Select-Object Name, SamAccountName |
        Sort-Object Name)
}
catch {
    Write-Host "Could not read group '$GroupName': $($_.Exception.Message)" -ForegroundColor Red
    return
}

if ($groupMembers.Count -eq 0) {
    Write-Host "No users found in the group '$GroupName'."
}
else {
    Write-Host "$($groupMembers.Count) users in the group '$GroupName':"
    $groupMembers
}