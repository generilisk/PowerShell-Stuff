<#
.SYNOPSIS
    Retrieves the badge number for a specified user from Active Directory.
.DESCRIPTION
    Looks up a user in Active Directory and returns the badge number stored in
    their extensionAttribute3 attribute. Can be run interactively (prompts for a
    username if none is supplied) or called from another script via -Username,
    with the result captured or piped onward. Returns nothing (with a warning)
    if the user can't be found or has no badge number set.
.PARAMETER Username
    The SamAccountName of the user to look up. If omitted, the script prompts for one.
.EXAMPLE
    PS C:\> .\Get-BadgeByUser.ps1 -Username clarkent
    48213

    Returns the badge number for user 'clarkent'.
.EXAMPLE
    PS C:\> $badge = .\Get-BadgeByUser.ps1 -Username clarkent

    Captures the returned badge number for use in another script.
.EXAMPLE
    PS C:\> .\Get-BadgeByUser.ps1
    Enter username to find matching badge number: clarkent
    48213

    Prompts interactively when no -Username is supplied.
.EXAMPLE
    PS C:\> .\Get-BadgeByUser.ps1 -Username tonstark
    WARNING: Unable to look up user tonstark, please check spelling or AD connectivity

    Shows the warning displayed when the user can't be found; nothing is returned.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Username
)

# Script Name: Get-BadgeByUser.ps1
#Created by Will Hughes
#Date: 2024-08-07
#Patch Notes:
<#
    2024-11-20: Removed self-congratulatory message disguised as information added by AI-generated description.
    2026-10-08: Added -Username parameter (prompts if omitted) and switched output
    from Write-Host to returning the badge number through the pipeline, so this can
    be called from another script like Get-UserByBadge.ps1. Errors now go to
    Write-Warning, added a blank-input guard, and added a warning when the user
    exists but has no badge number set (previously printed a blank line). The
    "Looking up..." message moved to Write-Verbose.
#>

if ([string]::IsNullOrWhiteSpace($Username)) {
    $Username = Read-Host "Enter username to find matching badge number"
}

if ([string]::IsNullOrWhiteSpace($Username)) {
    Write-Warning "No username entered."
    return
}

Write-Verbose "Looking up badge number for $Username..."

try {
    $badgeNumber = (Get-ADUser -Identity $Username -Properties extensionAttribute3).extensionAttribute3
}
catch {
    Write-Warning "Unable to look up user $Username, please check spelling or AD connectivity"
    return
}

if ([string]::IsNullOrWhiteSpace($badgeNumber)) {
    Write-Warning "User $Username was found, but has no badge number set."
    return
}

$badgeNumber