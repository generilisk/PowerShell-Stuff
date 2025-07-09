<#
.Synopsis
    Generates a unique username based on the organization's naming convention with Active Directory validation.
.DESCRIPTION
    This script prompts for a first and last name, then creates a username using the following logic:
    1. If last name > 6 characters: Use first initial + last name
    2. If full name <= 8 characters: Use full first name + last name
    3. If last name <= 6 and full name > 8: Trim first name to make exactly 8 characters
    
    The script then checks Active Directory for username availability and handles collisions by:
    - Adding additional letters from the first name (johsmith -> johnsmith)
    - Only when all letters are used, appending numbers (johnsmith1, johnsmith2, etc.)
    
    Requires Active Directory PowerShell module and appropriate permissions.
.PARAMETER FirstName
    The user's first name (prompted during execution)
.PARAMETER LastName
    The user's last name (prompted during execution)
.EXAMPLE
    New-Username
    # Prompts for input and generates a unique username.
    # For "John Smith": tries "johsmith", if taken tries "johnsmith", if taken tries "johnsmith1", etc.
.EXAMPLE
    New-Username
    # For "Bruce Wayne": returns "bruwayne" (if available)
    # For "Ang Lee": returns "anglee" (if available)
.NOTES
    Requires:
    - Active Directory PowerShell module
    - Appropriate AD read permissions
    - Domain connectivity
#>

# Script Name: New-Username
# Created by Will Hughes
# Date: 2024-10-31
# Updated: 2025-01-09
# Patch Notes: 
# - Added Active Directory integration with Get-ADUser validation
# - Implemented smart collision handling (letters before numbers)
# - Added comprehensive input validation
# - Improved output formatting with original name display
# - Added color-coded status messages for better UX

# Check if Active Directory module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "Active Directory module is not installed. Please install RSAT tools or run this on a domain controller."
    return
}

# Import Active Directory module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "Active Directory module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to import Active Directory module: $($_.Exception.Message)"
    return
}

function Test-ADUsername {
    param(
        [string]$Username
    )
    
    try {
        # Check if the username exists in Active Directory
        $adUser = Get-ADUser -Identity $Username -ErrorAction Stop
        return $true  # Username exists
    }
    catch {
        return $false  # Username doesn't exist
    }
}

function Get-UniqueUsername {
    param(
        [string]$BaseUsername,
        [string]$FirstName,
        [string]$LastName
    )
    
    # First, check if the base username is available
    if (-not (Test-ADUsername -Username $BaseUsername)) {
        Write-Host "Username '$BaseUsername' is available" -ForegroundColor Green
        return $BaseUsername
    }
    
    Write-Host "Username '$BaseUsername' is taken, checking alternatives..." -ForegroundColor Yellow
    
    # Try adding letters from the first name
    $currentUsername = $BaseUsername
    $firstNameIndex = $currentUsername.Length - $LastName.Length  # Current position in first name
    
    # Keep adding letters from first name until we run out or find a unique username
    while ($firstNameIndex -lt $FirstName.Length) {
        $currentUsername = $FirstName.Substring(0, $firstNameIndex + 1) + $LastName
        
        if (-not (Test-ADUsername -Username $currentUsername)) {
            Write-Host "Found available username: '$currentUsername'" -ForegroundColor Green
            return $currentUsername
        }
        
        Write-Host "Username '$currentUsername' is also taken" -ForegroundColor Yellow
        $firstNameIndex++
    }
    
    # If we've used all letters from first name, start appending numbers
    Write-Host "All letter combinations taken, trying numbers..." -ForegroundColor Yellow
    $counter = 1
    
    do {
        $numberedUsername = $currentUsername + $counter
        
        if (-not (Test-ADUsername -Username $numberedUsername)) {
            Write-Host "Found available username: '$numberedUsername'" -ForegroundColor Green
            return $numberedUsername
        }
        
        Write-Host "Username '$numberedUsername' is also taken" -ForegroundColor Yellow
        $counter++
    } while ($counter -le 99)  # Reasonable limit
    
    # If we somehow get here, return the base username with a timestamp
    $timestampUsername = $BaseUsername + (Get-Date -Format "MMdd")
    Write-Warning "Unable to find unique username after 99 attempts, using timestamp: '$timestampUsername'"
    return $timestampUsername
}

function New-Username {
    # Prompt for first and last names
    $firstName = Read-Host "Enter the first name"
    $lastName = Read-Host "Enter the last name"
    
    # Validate input - check for empty or whitespace-only strings
    if ([string]::IsNullOrWhiteSpace($firstName)) {
        Write-Error "First name cannot be empty or contain only whitespace"
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($lastName)) {
        Write-Error "Last name cannot be empty or contain only whitespace"
        return
    }
    
    # Store original names for display purposes
    $originalFirstName = $firstName.Trim()
    $originalLastName = $lastName.Trim()
    
    # Clean and normalize the names for processing
    $firstName = $firstName.ToLower()
    $firstName = $firstName -replace '\s', ''
    $lastName = $lastName.ToLower()
    $lastName = $lastName -replace '\s', ''
    
    # Additional validation after cleaning - ensure they're not empty after removing spaces
    if ([string]::IsNullOrEmpty($firstName)) {
        Write-Error "First name cannot be empty after removing spaces"
        return
    }
    
    if ([string]::IsNullOrEmpty($lastName)) {
        Write-Error "Last name cannot be empty after removing spaces"
        return
    }

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
    
    # Check if username exists in Active Directory and handle collisions
    $finalUsername = Get-UniqueUsername -BaseUsername $username -FirstName $firstName -LastName $lastName
    
    Write-Output "Generated username for ${originalLastName}, ${originalFirstName}: $finalUsername"
}

# Run the function
New-Username