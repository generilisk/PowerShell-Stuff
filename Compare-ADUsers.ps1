#Requires -Module ActiveDirectory

<#
.SYNOPSIS
    Compare two Active Directory user accounts
.DESCRIPTION
    This script compares two Active Directory user accounts and provides a detailed comparison of their AD attributes and group memberships.
.PARAMETER User1
    Specifies the first Active Directory user account for comparison.
    Type: String
    Required: No (script will prompt if not provided)
    Accepts: Username, SamAccountName, or Distinguished Name
.PARAMETER User2
    Specifies the second Active Directory user account for comparison.
    Type: String
    Required: No (script will prompt if not provided)
    Accepts: Username, SamAccountName, or Distinguished Name
.EXAMPLE
    .\Compare-ADUsers.ps1
    
    Interactive mode - script prompts for both usernames:
    This example runs the script without parameters, prompting the user to enter two usernames interactively.
    The script will ask for the first username, then the second username, and perform the comparison.
.EXAMPLE
    .\Compare-ADUsers.ps1 -User1 peparker -User2 clarkent
    
    Non-interactive mode - direct parameter usage:
    This example compares the Active Directory accounts for users 'peparker' and 'clarkent' without prompting for input.
    The script will retrieve both user accounts and display their differences in attributes and group memberships.
.AUTHOR
    Will Hughes
.VERSION
    1.2
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$User1,
    
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$User2
)

# Function to get user input with validation
function Get-ValidatedUsername {
    param(
        [string]$username,
        [string]$promptText
    )
    
    do {
        # If username is provided as a parameter, use it. Otherwise, prompt.
        if ([string]::IsNullOrEmpty($username)) {
            $username = Read-Host $promptText
        }
        
        if ([string]::IsNullOrWhiteSpace($username)) {
            Write-Host "Username cannot be empty. Please try again." -ForegroundColor Red
            $username = $null # Reset to prompt again
            continue
        }
        
        try {
            # Try to get the user; if it fails, the catch block will handle it.
            $user = Get-ADUser -Identity $username -ErrorAction Stop
            return $user.SamAccountName
        }
        catch {
            Write-Host "User '$username' not found in Active Directory. Please try again." -ForegroundColor Red
            $username = $null # Reset to prompt again
        }
    } while ([string]::IsNullOrEmpty($username))
}

# Function to format user information for display
function Format-UserInfo {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        [string]$UserLabel
    )
    
    Write-Output "`n=== $UserLabel ==="
    Write-Output "Username: $($User.SamAccountName)"
    Write-Output "Display Name: $($User.DisplayName)"
    Write-Output "Email: $($User.EmailAddress)"
    Write-Output "Enabled: $($User.Enabled)"
    Write-Output "Distinguished Name: $($User.DistinguishedName)"
    Write-Output "Department: $($User.Department)"
    Write-Output "Title: $($User.Title)"
    Write-Output "Manager: $($User.Manager)"
    Write-Output "Office: $($User.Office)"
    Write-Output "Phone: $($User.OfficePhone)"
    Write-Output "Last Logon: $($User.LastLogonDate)"
    Write-Output "Password Last Set: $($User.PasswordLastSet)"
    Write-Output "Password Never Expires: $($User.PasswordNeverExpires)"
    Write-Output "Account Locked: $($User.LockedOut)"
    Write-Output "Created: $($User.WhenCreated)"
    Write-Output "Modified: $($User.WhenChanged)"
}

# Function to compare user attributes
function Compare-UserAttributes {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User1,
        [Microsoft.ActiveDirectory.Management.ADUser]$User2,
        [string]$User1Name,
        [string]$User2Name
    )
    
    Write-Output "`n=== ATTRIBUTE COMPARISON ==="
    
    $attributes = @(
        'DisplayName', 'EmailAddress', 'Enabled', 'Department', 'Title', 'Manager',
        'Office', 'OfficePhone', 'PasswordNeverExpires', 'LockedOut'
    )
    
    $differences = @()
    
    foreach ($attr in $attributes) {
        $value1 = $User1.$attr
        $value2 = $User2.$attr
        
        if ($value1 -ne $value2) {
            $differences += [PSCustomObject]@{
                Attribute = $attr
                $User1Name = $value1
                $User2Name = $value2
            }
        }
    }
    
    if ($differences.Count -eq 0) {
        Write-Output "No differences found in compared attributes."
    } else {
        Write-Output "Differences found:"
        $differences | Format-Table -AutoSize
    }
}

# Function to compare group memberships
function Compare-GroupMemberships {
    param(
        [string]$Username1,
        [string]$Username2,
        [string]$User1Name,
        [string]$User2Name
    )
    
    Write-Output "`n=== GROUP MEMBERSHIP COMPARISON ==="
    
    # Get group memberships
    $groups1 = Get-ADPrincipalGroupMembership -Identity $Username1 | Select-Object -ExpandProperty Name | Sort-Object
    $groups2 = Get-ADPrincipalGroupMembership -Identity $Username2 | Select-Object -ExpandProperty Name | Sort-Object
    
    # Find unique groups
    $onlyInUser1 = $groups1 | Where-Object { $_ -notin $groups2 }
    $onlyInUser2 = $groups2 | Where-Object { $_ -notin $groups1 }
    $common = $groups1 | Where-Object { $_ -in $groups2 }
    
    Write-Output "`nGroups only in $User1Name ($($onlyInUser1.Count)):"
    if ($onlyInUser1.Count -gt 0) {
        $onlyInUser1 | ForEach-Object { Write-Output "  - $_" }
    } else {
        Write-Output "  None"
    }
    
    Write-Output "`nGroups only in $User2Name ($($onlyInUser2.Count)):"
    if ($onlyInUser2.Count -gt 0) {
        $onlyInUser2 | ForEach-Object { Write-Output "  - $_" }
    } else {
        Write-Output "  None"
    }
    
    Write-Output "`nCommon groups ($($common.Count)):"
    if ($common.Count -gt 0) {
        $common | ForEach-Object { Write-Output "  - $_" }
    } else {
        Write-Output "  None"
    }
}

Clear-Host

# Main script execution
try {
    # Check if Active Directory module is available
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "Active Directory PowerShell module is not installed. Please install RSAT tools."
    }
    
    Import-Module ActiveDirectory
    
    Write-Host "=== Active Directory User Comparison Tool ===" -ForegroundColor Green
    Write-Host "This tool will compare two AD user accounts and highlight differences.`n"

    # Get and validate usernames using the single, merged function
    $username1 = Get-ValidatedUsername -username $User1 -prompt "Enter the first username"
    $username2 = Get-ValidatedUsername -username $User2 -prompt "Enter the second username"

    if ($username1 -eq $username2) {
        Write-Host "Both usernames are the same. Please run the script again with different usernames." -ForegroundColor Yellow
        exit
    }
    
    # Get detailed user information with required properties
    Write-Host "`nRetrieving user information..." -ForegroundColor Yellow
    
    $requiredProperties = @(
        'DisplayName', 'EmailAddress', 'Enabled', 'Department', 'Title', 'Manager',
        'Office', 'OfficePhone', 'PasswordNeverExpires', 'LockedOut', 'LastLogonDate',
        'PasswordLastSet', 'WhenCreated', 'WhenChanged'
    )

    $user1 = Get-ADUser -Identity $username1 -Properties $requiredProperties
    $user2 = Get-ADUser -Identity $username2 -Properties $requiredProperties
    
    # Store the output in a variable
    $scriptOutput = New-Object -TypeName System.Text.StringBuilder
    $scriptOutput.Append((Format-UserInfo -User $user1 -UserLabel "USER 1: $($user1.SamAccountName)" | Out-String)) | Out-Null
    $scriptOutput.Append((Format-UserInfo -User $user2 -UserLabel "USER 2: $($user2.SamAccountName)" | Out-String)) | Out-Null
    $scriptOutput.Append((Compare-UserAttributes -User1 $user1 -User2 $user2 -User1Name $($user1.SamAccountName) -User2Name $($user2.SamAccountName) | Out-String)) | Out-Null
    $scriptOutput.Append((Compare-GroupMemberships -Username1 $user1.SamAccountName -Username2 $user2.SamAccountName -User1Name $($user1.SamAccountName) -User2Name $($user2.SamAccountName) | Out-String)) | Out-Null
    
    # Display the final output to the console
    Write-Host $scriptOutput.ToString()
    
    Write-Host "`n=== COMPARISON COMPLETE ===" -ForegroundColor Green
    
    # Option to export results
    $export = Read-Host "`nWould you like to export the comparison results to a file? (y/n)"
    if ($export -eq 'y' -or $export -eq 'Y') {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filename = "ADUserComparison_${user1.SamAccountName}_vs_${user2.SamAccountName}_${timestamp}.txt"
        
        $scriptOutput.ToString() | Out-File -FilePath $filename -Encoding UTF8
        Write-Host "Results exported to: $filename" -ForegroundColor Green
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please ensure you have the necessary permissions to query Active Directory." -ForegroundColor Yellow
}

# Pause to allow user to review results using a more compatible method
Write-Host "`nPress Enter to continue..." -ForegroundColor Gray
$null = Read-Host