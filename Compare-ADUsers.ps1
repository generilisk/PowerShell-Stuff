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
    
    Interactive mode - script prompts for both usernames.
.EXAMPLE
    .\Compare-ADUsers.ps1 -User1 peparker -User2 clarkent
    
    Non-interactive mode - direct parameter usage.
.AUTHOR
    Will Hughes
.VERSION
    1.3
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$User1,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$User2
)

# --- GLOBAL CONFIGURATION (Moved here for use in functions) ---

# Define the properties needed for comparison to optimize Get-ADUser calls
$RequiredADProperties = @(
    'DisplayName', 'EmailAddress', 'Enabled', 'Department', 'Title', 'Manager',
    'Office', 'OfficePhone', 'PasswordNeverExpires', 'LockedOut', 'LastLogonDate',
    'PasswordLastSet', 'WhenCreated', 'WhenChanged'
)

# Attributes to compare in the Compare-UserAttributes function
$ComparisonAttributes = @(
    'DisplayName', 'EmailAddress', 'Enabled', 'Department', 'Title', 'Manager',
    'Office', 'OfficePhone', 'PasswordNeverExpires', 'LockedOut'
)

# --- FUNCTIONS ---

# Function to get user input with validation
function Get-ValidatedUsername {
    param(
        [string]$username,
        [string]$promptText
    )

    # Use the globally defined required properties
    $props = $script:RequiredADProperties

    do {
        if ([string]::IsNullOrWhiteSpace($username)) {
            $username = Read-Host $promptText
        }

        if ([string]::IsNullOrWhiteSpace($username)) {
            Write-Host "Username cannot be empty. Please try again." -ForegroundColor Red
            $username = $null
            continue
        }

        try {
            # FIX: Added -Properties parameter to ensure the AD object is complete
            $user = Get-ADUser -Identity $username -Properties $props -ErrorAction Stop
            if (-not $user) {
                throw "No AD user object returned for '$username'."
            }
            return $user  # return the full AD user object with all needed properties
        }
        catch {
            Write-Host "User '$username' not found in Active Directory. Please try again." -ForegroundColor Red
            # IMPORTANT: Clear the username so Read-Host prompts again
            $username = $null
        }
    } while (-not $user)
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
        [Microsoft.ActiveDirectory.Management.ADUser]$User1Object,
        [Microsoft.ActiveDirectory.Management.ADUser]$User2Object,
        [string]$User1Name,
        [string]$User2Name
    )

    Write-Output "`n=== ATTRIBUTE COMPARISON ==="

    # Use the globally defined comparison attributes
    $attributes = $script:ComparisonAttributes

    $differences = @()

    foreach ($attr in $attributes) {
        $value1 = $User1Object.$attr
        $value2 = $User2Object.$attr

        # Note: -ne works correctly for all data types (string, boolean, date)
        if ($value1 -ne $value2) {
            $differences += [PSCustomObject]@{
                Attribute  = $attr
                $User1Name = $value1
                $User2Name = $value2
            }
        }
    }

    if ($differences.Count -eq 0) {
        Write-Output "No differences found in compared attributes."
    }
    else {
        Write-Output "Differences found:"
        $differences | Format-Table -AutoSize
    }
}

# Function to compare group memberships
function Compare-GroupMemberships {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User1Object,
        [Microsoft.ActiveDirectory.Management.ADUser]$User2Object,
        [string]$User1Name,
        [string]$User2Name
    )

    Write-Output "`n=== GROUP MEMBERSHIP COMPARISON ==="

    $groups1 = Get-ADPrincipalGroupMembership -Identity $User1Object | Select-Object -ExpandProperty Name | Sort-Object
    $groups2 = Get-ADPrincipalGroupMembership -Identity $User2Object | Select-Object -ExpandProperty Name | Sort-Object

    # FIX: Use case-insensitive comparison operators (-notcontains, -contains) 
    # as group names are typically not case-sensitive in AD.
    $onlyInUser1 = $groups1 | Where-Object { $groups2 -notcontains $_ }
    $onlyInUser2 = $groups2 | Where-Object { $groups1 -notcontains $_ }
    $common = $groups1 | Where-Object { $groups2 -contains $_ }

    Write-Output "`nGroups only in $User1Name ($($onlyInUser1.Count)):"
    if ($onlyInUser1.Count -gt 0) {
        $onlyInUser1 | ForEach-Object { Write-Output "   - $_" }
    }
    else {
        Write-Output "   None"
    }

    Write-Output "`nGroups only in $User2Name ($($onlyInUser2.Count)):"
    if ($onlyInUser2.Count -gt 0) {
        $onlyInUser2 | ForEach-Object { Write-Output "   - $_" }
    }
    else {
        Write-Output "   None"
    }

    Write-Output "`nCommon groups ($($common.Count)):"
    if ($common.Count -gt 0) {
        $common | ForEach-Object { Write-Output "   - $_" }
    }
    else {
        Write-Output "   None"
    }
}

Clear-Host

# --- MAIN SCRIPT EXECUTION ---
try {
    # Check if module is loaded, then import if needed
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "Active Directory PowerShell module is not installed. Please install RSAT tools."
    }
    
    # Import-Module will not re-import if already loaded
    Import-Module ActiveDirectory

    Write-Host "=== Active Directory User Comparison Tool ===" -ForegroundColor Green
    Write-Host "This tool will compare two AD user accounts and highlight differences.`n"

    # FIX: Use new variables ($ADUser1, $ADUser2) to store the AD objects 
    # instead of overwriting the $User1/$User2 string parameters.
    $ADUser1 = Get-ValidatedUsername -username $User1 -prompt "Enter the first username"
    $ADUser2 = Get-ValidatedUsername -username $User2 -prompt "Enter the second username"

    if (-not $ADUser1 -or -not $ADUser2) {
        throw "One or both users could not be validated. Exiting."
    }

    if ($ADUser1.SamAccountName -eq $ADUser2.SamAccountName) {
        Write-Host "Both usernames are the same. Please run the script again with different usernames." -ForegroundColor Yellow
        exit
    }

    Write-Host "`nRetrieving user information..." -ForegroundColor Yellow

    # FIX: Removed the redundant Get-ADUser calls here. The users were 
    # already retrieved with the full property set in Get-ValidatedUsername.
    
    $scriptOutput = New-Object -TypeName System.Text.StringBuilder
    $scriptOutput.Append((Format-UserInfo -User $ADUser1 -UserLabel "USER 1: $($ADUser1.SamAccountName)" | Out-String)) | Out-Null
    $scriptOutput.Append((Format-UserInfo -User $ADUser2 -UserLabel "USER 2: $($ADUser2.SamAccountName)" | Out-String)) | Out-Null
    $scriptOutput.Append((Compare-UserAttributes -User1Object $ADUser1 -User2Object $ADUser2 -User1Name $($ADUser1.SamAccountName) -User2Name $($ADUser2.SamAccountName) | Out-String)) | Out-Null
    $scriptOutput.Append((Compare-GroupMemberships -User1Object $ADUser1 -User2Object $ADUser2 -User1Name $($ADUser1.SamAccountName) -User2Name $($ADUser2.SamAccountName) | Out-String)) | Out-Null

    Write-Host $scriptOutput.ToString()

    Write-Host "`n=== COMPARISON COMPLETE ===" -ForegroundColor Green

    $export = Read-Host "`nWould you like to export the comparison results to a file? (y/n)"
    if ($export -eq 'y' -or $export -eq 'Y') {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filename = "ADUserComparison_$($ADUser1.SamAccountName)_vs_$($ADUser2.SamAccountName)_$timestamp.txt"

        $scriptOutput.ToString() | Out-File -FilePath $filename -Encoding UTF8
        Write-Host "Results exported to: $filename" -ForegroundColor Green
    }

}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please ensure you have the necessary permissions to query Active Directory." -ForegroundColor Yellow
}

Write-Host "`nPress Enter to continue..." -ForegroundColor Gray
$null = Read-Host