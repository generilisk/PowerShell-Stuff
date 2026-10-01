<#
.SYNOPSIS
    Finds the Active Directory user(s) associated with a given badge number.
.DESCRIPTION
    Searches Active Directory for user(s) whose extensionAttribute3 matches
    the supplied badge number and returns the matching ADUser object(s).
    Can be run interactively (prompts for a badge number if none is
    supplied) or called from another script via -BadgeNumber, with the
    result captured or piped onward.
.PARAMETER BadgeNumber
    The badge number to search for. If omitted, the script prompts for one.
.EXAMPLE
    PS C:\> .\check-BadgeNumber.ps1 -BadgeNumber 48213

    Returns the ADUser object whose extensionAttribute3 is 48213.
.EXAMPLE
    PS C:\> $user = .\check-BadgeNumber.ps1 -BadgeNumber 48213
    PS C:\> $user.Name
    Clark Kent

    Captures the returned user object for use in another script.
.EXAMPLE
    PS C:\> .\check-BadgeNumber.ps1
    Enter the badge number you want to search for: 48213

    Prompts interactively when no -BadgeNumber is supplied.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$BadgeNumber
)

# Script Name: check-BadgeNumber.ps1
# Patch Notes:
<#
    2026-09-26: Added -BadgeNumber parameter, switched output from
    Write-Host text to returning ADUser object(s) through the pipeline,
    and dropped the trailing Read-Host pause -- so this can be called
    from another script instead of only run interactively. Status and
    not-found/error messages moved to Write-Warning so they don't mix
    into the object output. Added a blank-input guard and wrapped the
    AD lookup in try/catch to match get-BadgeNumber.ps1's error handling.
#>

if ([string]::IsNullOrWhiteSpace($BadgeNumber)) {
    $BadgeNumber = Read-Host "Enter the badge number you want to search for"
}

if ([string]::IsNullOrWhiteSpace($BadgeNumber)) {
    Write-Warning "No badge number entered."
    return
}

try {
    $user = Get-ADUser -Filter { extensionAttribute3 -eq $BadgeNumber }

    if ($user) {
        $user
    }
    else {
        Write-Warning "No user found with the badge number: $BadgeNumber"
    }
}
catch {
    Write-Warning "Unable to look up badge number $BadgeNumber, please check spelling or AD connectivity"
}