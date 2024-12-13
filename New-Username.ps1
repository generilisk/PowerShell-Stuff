<#
.Synopsis
    Generates a username based on the organization's naming convention.
.DESCRIPTION
    This script prompts for a first and last name, then creates a username by combining the first initial and last name.
    If the result is fewer than 8 characters, it appends additional letters from the first name as needed, without trimming 
    the last name.
.EXAMPLE
    Generate-Username
    # Prompts for input and generates a username based on the specified logic.
    # For first name "Bruce" and last name "Wayne", returns "bruwayne".
#>

# Script Name: Generate-Username
# Created by Will Hughes
# Date: 2024-10-31
# Patch Notes: Ensured the last name is not trimmed.

function New-Username {
    # Prompt for first and last names
    $firstName = Read-Host "Enter the first name"
    $lastName = Read-Host "Enter the last name"
    $firstName = $firstName.ToLower()
    $lastName = $lastName.ToLower()

    switch($lastName){
        {$_.Length -gt 6} {
            $username = ($firstName.Substring(0,1) + $lastName)
        }
        {$_.Length + $firstName.Length -le 8} {
            $username = ($firstName + $lastName)
        }
        {($_.Length -lt 7) -and ($_.Length + $firstName.Length -gt 8)} {
            $trimLength = (8 - $lastName.Length)
            $firstNameTrimmed = $firstName.Substring(0,$trimLength)
            $username = $firstNameTrimmed  + $lastName
        }
    }
    Write-Output "Generated username: $username"
}

# Run the function
New-Username