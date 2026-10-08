<#
.SYNOPSIS
    Retrieves the last password change date and expiration date for an Active Directory user.
.DESCRIPTION
    Fetches the pwdLastSet attribute and the computed password expiration date
    (msDS-UserPasswordExpiryTimeComputed) for a given Active Directory user.
    Handles accounts with "password never expires" set, since that attribute
    returns a sentinel value in that case rather than a usable date.
    If no username is provided as a parameter, the script prompts for one.
.PARAMETER Username
    The SamAccountName of the user to look up. If omitted, the script prompts for it.
.EXAMPLE
    PS C:\> .\Get-PasswordExpirationDate.ps1 -Username bruwayne

    ===================================
    User: bruwayne
    Password Last Set: 01/15/2026 08:12:03
    Password Expires On: 04/15/2026 08:12:03
    ===================================

    Retrieves password last set and expiration date for user 'bruwayne'.
.EXAMPLE
    PS C:\> .\Get-PasswordExpirationDate.ps1 -Username diprince

    User 'diprince' has "password never expires" set.
    Password Last Set: 03/02/2025 14:40:11

    Shows the output when the account's password is set to never expire.
.EXAMPLE
    PS C:\> .\Get-PasswordExpirationDate.ps1

    Enter the username:

    Prompts for a username if none is provided, then retrieves password details.
#>

# Get-PasswordExpirationDate.ps1
# Created by Will Hughes
# Date: 2025.01.30
# Patch Notes:
# 2025.01.30 - Completely re-wrote as I thought I'd lost this. I like this version better.
# 2026.10.02 - Corrected header/example script name (was still "Get-PasswordExpiration.ps1").
#              Wrapped Get-ADUser in try/catch -- previously an invalid username threw a raw
#              terminating error before the "not found" check could ever run. Added explicit
#              handling for the "password never expires" sentinel value
#              (9223372036854775807), which previously crashed FromFileTime() with an
#              ArgumentOutOfRangeException instead of displaying a clean message. Replaced
#              exit with return so this can be called from another script without killing
#              the whole session.

param (
    [string]$Username
)

# msDS-UserPasswordExpiryTimeComputed returns this sentinel value when
# "password never expires" is set on the account, instead of 0 or $null.
$NeverExpiresValue = [int64]::MaxValue

if (-not $Username) {
    $Username = Read-Host "Enter the username"
}

try {
    $User = Get-ADUser -Identity $Username -Properties "pwdLastSet", "msDS-UserPasswordExpiryTimeComputed"
}
catch {
    Write-Host "User '$Username' not found." -ForegroundColor Red
    return
}

$PasswordLastSet = if ($User.pwdLastSet) { [datetime]::FromFileTime($User.pwdLastSet) } else { "N/A" }
$ExpiryRaw = $User."msDS-UserPasswordExpiryTimeComputed"

if ($ExpiryRaw -eq $NeverExpiresValue) {
    Write-Host "`nUser '$Username' has `"password never expires`" set." -ForegroundColor Yellow
    Write-Host "Password Last Set: $PasswordLastSet"
}
elseif ($ExpiryRaw) {
    $ExpirationDate = [datetime]::FromFileTime($ExpiryRaw)
    Write-Host "`n===================================" -ForegroundColor Cyan
    Write-Host "User: $Username" -ForegroundColor Green
    Write-Host "Password Last Set: $PasswordLastSet" -ForegroundColor Yellow
    Write-Host "Password Expires On: $ExpirationDate" -ForegroundColor Magenta
    Write-Host "===================================" -ForegroundColor Cyan
}
else {
    Write-Host "`nUnable to retrieve password expiration date for user '$Username'." -ForegroundColor Red
}