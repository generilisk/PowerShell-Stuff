<#
.Synopsis
    Retrieves the last password change date and expiration date for an Active Directory user.
.DESCRIPTION
    This script fetches the `pwdLastSet` attribute and the computed password expiration date for a given Active Directory user.
    If no username is provided as a parameter, the script will prompt the user for input.
.EXAMPLE
    .\Get-PasswordExpiration.ps1 -Username johndoe
    Retrieves password last set and expiration date for user 'johndoe'.
.EXAMPLE
    .\Get-PasswordExpiration.ps1
    Prompts for a username if none is provided, then retrieves password details.
#>

# Get-PasswordExpiration.ps1
# Created by Will Hughes
# Date: 2025.01.30
# Patch Notes:
    #2025.01.30 - Completely re-wrote as I thought I'd lost this. I like this version better.

param (
    [string]$Username
)

# If no username is provided, prompt the user
if (-not $Username) {
    $Username = Read-Host "Enter the username"
}

# Get user details including msDS-UserPasswordExpiryTimeComputed for accurate expiration date
$User = Get-ADUser -Identity $Username -Properties "pwdLastSet", "msDS-UserPasswordExpiryTimeComputed"

if (-not $User) {
    Write-Host "User '$Username' not found." -ForegroundColor Red
    exit
}

# Convert pwdLastSet to readable date
$PasswordLastSet = if ($User.pwdLastSet) { [datetime]::FromFileTime($User.pwdLastSet) } else { "N/A" }

# Get password expiration date
if ($User."msDS-UserPasswordExpiryTimeComputed") {
    $ExpirationDate = [datetime]::FromFileTime($User."msDS-UserPasswordExpiryTimeComputed")
    Write-Host "`n===================================" -ForegroundColor Cyan
    Write-Host "User: $Username" -ForegroundColor Green
    Write-Host "Password Last Set: $PasswordLastSet" -ForegroundColor Yellow
    Write-Host "Password Expires On: $ExpirationDate" -ForegroundColor Magenta
    Write-Host "===================================" -ForegroundColor Cyan
} else {
    Write-Host "`nUnable to retrieve password expiration date for user '$Username'." -ForegroundColor Red
}
