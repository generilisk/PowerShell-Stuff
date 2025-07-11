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
    1.1
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

# Function to get user input if missing
function Get-UsernameIfMissing {
    param(
        [string]$username,
        [string]$prompt
    )
    
    if ([string]::IsNullOrEmpty($username)) {
        $username = Read-Host -Prompt $prompt
    }
    return $username
}

# Function to get user input with validation
function Get-ValidatedUsername {
    param(
        [string]$PromptText
    )
    
    do {
        $username = Read-Host $PromptText
        if ([string]::IsNullOrWhiteSpace($username)) {
            Write-Host "Username cannot be empty. Please try again." -ForegroundColor Red
            continue
        }
        
        try {
            $user = Get-ADUser -Identity $username -ErrorAction Stop
            return $username
        }
        catch {
            Write-Host "User '$username' not found in Active Directory. Please try again." -ForegroundColor Red
        }
    } while ($true)
}

# Function to format user information for display
function Format-UserInfo {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        [string]$UserLabel
    )
    
    Write-Host "`n=== $UserLabel ===" -ForegroundColor Cyan
    Write-Host "Username: $($User.SamAccountName)"
    Write-Host "Display Name: $($User.DisplayName)"
    Write-Host "Email: $($User.EmailAddress)"
    Write-Host "Enabled: $($User.Enabled)"
    Write-Host "Distinguished Name: $($User.DistinguishedName)"
    Write-Host "Department: $($User.Department)"
    Write-Host "Title: $($User.Title)"
    Write-Host "Manager: $($User.Manager)"
    Write-Host "Office: $($User.Office)"
    Write-Host "Phone: $($User.OfficePhone)"
    Write-Host "Last Logon: $($User.LastLogonDate)"
    Write-Host "Password Last Set: $($User.PasswordLastSet)"
    Write-Host "Password Never Expires: $($User.PasswordNeverExpires)"
    Write-Host "Account Locked: $($User.LockedOut)"
    Write-Host "Created: $($User.WhenCreated)"
    Write-Host "Modified: $($User.WhenChanged)"
}

# Function to compare user attributes
function Compare-UserAttributes {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User1,
        [Microsoft.ActiveDirectory.Management.ADUser]$User2,
        [string]$User1Name,
        [string]$User2Name
    )
    
    Write-Host "`n=== ATTRIBUTE COMPARISON ===" -ForegroundColor Yellow
    
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
        Write-Host "No differences found in compared attributes." -ForegroundColor Green
    } else {
        Write-Host "Differences found:" -ForegroundColor Red
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
    
    Write-Host "`n=== GROUP MEMBERSHIP COMPARISON ===" -ForegroundColor Yellow
    
    # Get group memberships
    $groups1 = Get-ADPrincipalGroupMembership -Identity $Username1 | Select-Object -ExpandProperty Name | Sort-Object
    $groups2 = Get-ADPrincipalGroupMembership -Identity $Username2 | Select-Object -ExpandProperty Name | Sort-Object
    
    # Find unique groups
    $onlyInUser1 = $groups1 | Where-Object { $_ -notin $groups2 }
    $onlyInUser2 = $groups2 | Where-Object { $_ -notin $groups1 }
    $common = $groups1 | Where-Object { $_ -in $groups2 }
    
    Write-Host "`nGroups only in $User1Name ($($onlyInUser1.Count)):" -ForegroundColor Red
    if ($onlyInUser1.Count -gt 0) {
        $onlyInUser1 | ForEach-Object { Write-Host "  - $_" }
    } else {
        Write-Host "  None"
    }
    
    Write-Host "`nGroups only in $User2Name ($($onlyInUser2.Count)):" -ForegroundColor Red
    if ($onlyInUser2.Count -gt 0) {
        $onlyInUser2 | ForEach-Object { Write-Host "  - $_" }
    } else {
        Write-Host "  None"
    }
    
    Write-Host "`nCommon groups ($($common.Count)):" -ForegroundColor Green
    if ($common.Count -gt 0) {
        $common | ForEach-Object { Write-Host "  - $_" }
    } else {
        Write-Host "  None"
    }
}

# Main script execution
try {
    # Check if Active Directory module is available
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "Active Directory PowerShell module is not installed. Please install RSAT tools."
    }
    
    Import-Module ActiveDirectory
    
    Write-Host "=== Active Directory User Comparison Tool ===" -ForegroundColor Green
    Write-Host "This tool will compare two AD user accounts and highlight differences.`n"
    
# Get usernames using helper function
    $User1 = Get-UsernameIfMissing -username $User1 -prompt "Enter the first username"
    $User2 = Get-UsernameIfMissing -username $User2 -prompt "Enter the second username"

    if ($User1 -eq $User2) {
        Write-Host "Both usernames are the same. Please run the script again with different usernames." -ForegroundColor Yellow
        exit
    }
    
    # Get detailed user information
    Write-Host "`nRetrieving user information..." -ForegroundColor Yellow
    
    $user1 = Get-ADUser -Identity $User1 -Properties *
    $user2 = Get-ADUser -Identity $User2 -Properties *
    
    # Display user information
    Format-UserInfo -User $user1 -UserLabel "USER 1: $User1"
    Format-UserInfo -User $user2 -UserLabel "USER 2: $User2"
    
    # Compare attributes
    Compare-UserAttributes -User1 $user1 -User2 $user2 -User1Name $User1 -User2Name $User2
    
    # Compare group memberships
    Compare-GroupMemberships -Username1 $User1 -Username2 $User2 -User1Name $User1 -User2Name $User2
    
    Write-Host "`n=== COMPARISON COMPLETE ===" -ForegroundColor Green
    
    
    # Option to export results
    $export = Read-Host "`nWould you like to export the comparison results to a file? (y/n)"
    if ($export -eq 'y' -or $export -eq 'Y') {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filename = "ADUserComparison_${User1}_vs_${User2}_${timestamp}.txt"
        
        # Redirect output to file
        $scriptBlock = {
            Format-UserInfo -User $user1 -UserLabel "USER 1: $User1"
            Format-UserInfo -User $user2 -UserLabel "USER 2: $User2"
            Compare-UserAttributes -User1 $user1 -User2 $user2 -User1Name $User1 -User2Name $User2
            Compare-GroupMemberships -Username1 $User1 -Username2 $User2 -User1Name $User1 -User2Name $User2
        }
        
        & $scriptBlock | Out-File -FilePath $filename -Encoding UTF8
        Write-Host "Results exported to: $filename" -ForegroundColor Green
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please ensure you have the necessary permissions to query Active Directory." -ForegroundColor Yellow
}

# Pause to allow user to review results
Write-Host "`nPress any key to continue..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
