<#
.Synopsis
    Retrieve a user's password expiration date.
.DESCRIPTION
    This script retrieves the password expiration date for an Active Directory user. 
    The username can be provided as a parameter. If no username is supplied, the script will prompt for one.
.EXAMPLE
    .\Get-PasswordExpirationDate.ps1 -Username jdoe
    Retrieves the password expiration date for the user 'jdoe'.
.EXAMPLE
    .\Get-PasswordExpirationDate.ps1
    Prompts for a username and retrieves the password expiration date for the entered user.
#>

# Script Name: Get-PasswordExpirationDate.ps1
# Created by Will Hughes
# Date: 2024-10-07
# Patch Notes: Initial script creation

param (
    [string]$Username
)

if (-not $Username) {
    $Username = Read-Host "Please enter the username"
}

try {
    # Retrieve the password expiration date
    $user = Get-ADUser -Identity $Username -Properties "msDS-UserPasswordExpiryTimeComputed"

    if ($user -and $user."msDS-UserPasswordExpiryTimeComputed") {
        $expirationDate = [datetime]::FromFileTime($user."msDS-UserPasswordExpiryTimeComputed")
        Write-Host "Password for user $Username expires on: $expirationDate"
    } else {
        Write-Host "Unable to retrieve password expiration date for user $Username."
    }
} catch {
    Write-Host "Error: $_"
}