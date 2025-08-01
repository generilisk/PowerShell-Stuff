<#
.SYNOPSIS
Populates the 'Username' column in a CSV by resolving names via Active Directory.

.DESCRIPTION
Reads a CSV file with a 'Name' column (formatted as 'First Last') and fills in the 'Username' column
by querying Active Directory. Skips any rows where 'Username' is already populated.
Tries exact match using 'First Last' and 'Last, First' formats. If no match, allows interactive fuzzy selection.

.PARAMETER Path
The path to the CSV file (e.g., H:\offboard\Progress.csv).

.EXAMPLE
.\Update-OffboardList.ps1 -Path "H:\offboard\Progress.csv"

.NOTES
Requires RSAT tools (Active Directory module).
Must be run by a user with permission to query Active Directory.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$Path
)

# Import AD module
Import-Module ActiveDirectory -ErrorAction Stop

# Load CSV
try {
    $users = Import-Csv -Path $Path
} catch {
    Write-Error "Failed to load CSV: $_"
    return
}

foreach ($user in $users) {
    $name = $user.Name.Trim()
    $username = $user.Username

    if (![string]::IsNullOrWhiteSpace($username)) {
        continue
    }

    # Try exact match on DisplayName: First Last
    $adUser = Get-ADUser -Filter {DisplayName -eq $name} -Properties sAMAccountName

    # If no match, try Last, First
    if (-not $adUser -and $name -match '^\S+\s+\S+$') {
        $parts = $name -split '\s+'
        $reversedName = "$($parts[1]), $($parts[0])"
        $adUser = Get-ADUser -Filter {DisplayName -eq $reversedName} -Properties sAMAccountName
    }

    if ($adUser) {
        $user.Username = $adUser.sAMAccountName
        continue
    }

    # Fuzzy match fallback
    $potentialMatches = Get-ADUser -Filter {DisplayName -like "*$name*"} -Properties DisplayName, sAMAccountName

    if ($potentialMatches.Count -eq 0) {
        Write-Host "No match found for '$name'"
        $user.Username = ""
        continue
    }

    Write-Host "`nPossible matches for '$name':"
    $i = 1
    foreach ($match in $potentialMatches) {
        Write-Host "$i. $($match.DisplayName) [$($match.sAMAccountName)]"
        $i++
    }

    $choice = Read-Host "Enter the number of the correct user, or press Enter to skip"

    if ([string]::IsNullOrWhiteSpace($choice) -or -not ($choice -match '^\d+$')) {
        Write-Host "Skipped."
        $user.Username = ""
        continue
    }

    $selectedIndex = [int]$choice - 1

    if ($selectedIndex -ge 0 -and $selectedIndex -lt $potentialMatches.Count) {
        $selectedUser = $potentialMatches[$selectedIndex]
        $user.Username = $selectedUser.sAMAccountName
    } else {
        Write-Host "Invalid choice. Skipped."
        $user.Username = ""
    }
}

# Save updated CSV
$users | Export-Csv -Path $Path -NoTypeInformation
Write-Host "`nDone! Usernames updated in $Path"
