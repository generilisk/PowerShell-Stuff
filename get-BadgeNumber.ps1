<#
.SYNOPSIS
    Retrieves the badge number for a specified user from Active Directory.
.DESCRIPTION
    This script prompts the user for a username and looks up the corresponding badge number stored in the `extensionAttribute3` attribute of the user's Active Directory account. 
.EXAMPLE
    PS C:\> .\Get-BadgeNumber.ps1
    Enter username to find matching badge number: jdoe
    Looking up badge number for jdoe...
    12345

    This example prompts the user to enter a username (`jdoe`) and displays the badge number (`12345`) associated with that username.
.EXAMPLE
    PS C:\> .\Get-BadgeNumber.ps1
    Enter username to find matching badge number: nonexistentuser
    Looking up badge number for nonexistentuser...
    Unable to find nonexistentuser, please check spelling

    This example prompts the user to enter a username (`nonexistentuser`) and provides an error message when the user is not found.
#>

# Script Name: Get-BadgeNumber.ps1
#Created by Will Hughes
#Date: 2024-08-07
#Patch Notes:
<#
    2024-11-20: Removed self-congratulatory message disguised as information added by AI-generated description.
    #>

#Prompt for the user name
$userName = Read-Host "Enter username to find matching badge number"

Write-Host "Looking up badge number for $userName..."

try {
    $badgeNumber = (Get-ADUser -Identity $userName -Properties extensionAttribute3).extensionAttribute3
    Write-Host "$badgeNumber"
}
catch {
    Write-Host "Unable to find $userName, please check spelling"
}