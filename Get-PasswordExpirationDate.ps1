<#
.SYNOPSIS
    Retrieves the last password change date and expiration date for an Active Directory user.
.DESCRIPTION
    Fetches the pwdLastSet attribute and the computed password expiration date
    (msDS-UserPasswordExpiryTimeComputed) for a given Active Directory user.
    Handles accounts with "password never expires" set, since that attribute
    returns a sentinel value in that case rather than a usable date. Also handles
    accounts flagged to change their password at next logon (pwdLastSet is 0).
    If no username is provided as a parameter, the script prompts for one.
    If the Active Directory tools (RSAT) are missing they are installed automatically
    (needs an elevated session).
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
# 2026.10.13 - Installs RSAT Active Directory tools if missing (needs elevation). Blank username
#              is now rejected instead of being sent to AD. "Not found" is only reported when the
#              user really doesn't exist; other AD errors (no DC, access denied) now show their
#              actual message. Added a message for accounts set to change password at next logon
#              (previously reported as "unable to retrieve").

param (
    [string]$Username
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

# msDS-UserPasswordExpiryTimeComputed returns this sentinel value when
# "password never expires" is set on the account, instead of 0 or $null.
$NeverExpiresValue = [int64]::MaxValue

if ([string]::IsNullOrWhiteSpace($Username)) {
    $Username = Read-Host "Enter the username"
}
if ([string]::IsNullOrWhiteSpace($Username)) {
    Write-Host "No username entered." -ForegroundColor Red
    return
}

try {
    $User = Get-ADUser -Identity $Username -Properties "pwdLastSet", "msDS-UserPasswordExpiryTimeComputed" -ErrorAction Stop
}
catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    Write-Host "User '$Username' not found." -ForegroundColor Red
    return
}
catch {
    Write-Host "Could not look up user '$Username': $($_.Exception.Message)" -ForegroundColor Red
    return
}

$PasswordLastSet = if ($User.pwdLastSet) { [datetime]::FromFileTime($User.pwdLastSet) } else { "N/A" }
$ExpiryRaw = $User."msDS-UserPasswordExpiryTimeComputed"

if ($ExpiryRaw -eq $NeverExpiresValue) {
    Write-Host "`nUser '$Username' has `"password never expires`" set." -ForegroundColor Yellow
    Write-Host "Password Last Set: $PasswordLastSet"
}
elseif ($User.pwdLastSet -eq 0) {
    # pwdLastSet of 0 means the account is flagged to change its password at next logon
    Write-Host "`nUser '$Username' must change their password at next logon." -ForegroundColor Yellow
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