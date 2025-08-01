#Requires -Module ActiveDirectory

<#
.SYNOPSIS
    Generates offboarding form data for terminated employees from Active Directory.

.DESCRIPTION
    This script retrieves basic user information from Active Directory for terminated employees.
    Outputs formatted data to the console for copy-paste into MS eDiscovery forms.

.PARAMETER Username
    Specifies the Active Directory username of the employee being offboarded.
    Type: String
    Required: No (script will prompt if not provided)
    Accepts: Username, SamAccountName, or Distinguished Name

.EXAMPLE
    .\Generate-OffboardingFormData.ps1
    
    Interactive mode - script prompts for username and displays formatted eDiscovery data.

.EXAMPLE
    .\Generate-OffboardingFormData.ps1 -Username jsmith
    
    Direct parameter usage - generates eDiscovery data for user 'jsmith'.

.AUTHOR
    Will Hughes

.VERSION
    1.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$Username
)

# Helper function to get username if missing
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

# Error handling
try {
    # Get username using helper function
    $Username = Get-UsernameIfMissing -username $Username -prompt "Enter the username"
    
    # Retrieve user from Active Directory
    $user = Get-ADUser -Identity $Username -Properties GivenName, Surname -ErrorAction Stop
    
    # Compose variables
    $first = $user.GivenName
    $last = $user.Surname

    # Generate the four required output lines
    Write-Output $Username
    Write-Output "Off-boarding for $last, $first"
    Write-Output "${Username}_export"
    Write-Output "Off-boarding Export for $last, $first"

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
