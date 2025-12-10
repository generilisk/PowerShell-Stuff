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
# Updated: 2025-12-05
# Patch Notes: 
# - Added script-level parameters for direct execution support
# - Restored interactive prompting with hybrid automation support
# - Optimized AD checks using Filter instead of Try/Catch
# - Standardized output to console with color coding
# - Improved AD validation and collision handling logic

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0)]
    [string]$FirstName,
    
    [Parameter(Mandatory=$false, Position=1)]
    [string]$LastName
)

#Requires -Modules ActiveDirectory

<# # Check if Active Directory module is available (Redundant with #Requires but good for explicit error message if script is run in ISE/VSCode without analyzing prerequisites)
# However, #Requires is the standard way. #>

# Import Active Directory module not strictly needed if we rely on autoloading or #Requires, 
# but explicitness handles edge cases where autoload fails or errors are needed.
# Since we are optimizing, we will rely on #Requires and autoloading.


function Test-ADUsername {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Username
    )
    
    # Use Filter instead of try/catch for better performance
    $adUser = Get-ADUser -Filter "SamAccountName -eq '$Username'" -ErrorAction SilentlyContinue
    return [bool]$adUser
}

function Get-UniqueUsername {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseUsername,
        [Parameter(Mandatory=$true)]
        [string]$FirstName,
        [Parameter(Mandatory=$true)]
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
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false, Position=0)]
        [string]$FirstName,
        
        [Parameter(Mandatory=$false, Position=1)]
        [string]$LastName
    )
    
    # Interactive fallback: Prompt if parameters are missing
    if ([string]::IsNullOrWhiteSpace($FirstName)) {
        $FirstName = Read-Host "Enter the first name"
    }
    
    if ([string]::IsNullOrWhiteSpace($LastName)) {
        $LastName = Read-Host "Enter the last name"
    }
    
    # Store original names for display purposes
    $originalFirstName = $FirstName.Trim()
    $originalLastName = $LastName.Trim()
    
    # Clean and normalize the names for processing
    $FirstName = $FirstName.ToLower() -replace '\s', ''
    $LastName = $LastName.ToLower() -replace '\s', ''
    
    # Additional validation after cleaning - ensure they're not empty after removing spaces
    if ([string]::IsNullOrEmpty($FirstName)) {
        Write-Error "First name cannot be empty after removing spaces"
        return
    }
    
    if ([string]::IsNullOrEmpty($LastName)) {
        Write-Error "Last name cannot be empty after removing spaces"
        return
    }

    switch($LastName){
        {$_.Length -gt 6} {
            $username = ($FirstName.Substring(0,1) + $LastName)
        }
        {$_.Length + $FirstName.Length -le 8} {
            $username = ($FirstName + $LastName)
        }
        {($_.Length -lt 7) -and ($_.Length + $FirstName.Length -gt 8)} {
            $trimLength = (8 - $LastName.Length)
            $firstNameTrimmed = $FirstName.Substring(0,$trimLength)
            $username = $firstNameTrimmed  + $LastName
        }
    }
    
    # Check if username exists in Active Directory and handle collisions
    $finalUsername = Get-UniqueUsername -BaseUsername $username -FirstName $FirstName -LastName $LastName
    
    # Return the object (best practice) but also print the success message as requested
    
    Write-Host "Generated username for ${originalLastName}, ${originalFirstName}: $finalUsername" -ForegroundColor Cyan
    return $finalUsername
}


# Run the function if script is executed directly (not dot-sourced)
New-Username -FirstName $FirstName -LastName $LastName