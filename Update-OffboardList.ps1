<#
.SYNOPSIS
Populates the 'Username' column in a CSV by resolving names via Active Directory,
then looks up each user's assigned Microsoft 365 licenses and writes both back to the CSV.

.DESCRIPTION
Reads a CSV file with a 'Name' column (formatted as 'First Last') and fills in the 'Username' column
by querying Active Directory. Skips any rows where 'Username' is already populated.
Tries exact match using 'First Last' and 'Last, First' formats. If no match, allows interactive fuzzy selection.

If a name still has no username after automatic and fuzzy matching, the script prompts for manual
entry of the correct username before moving on to license lookup. Rows left blank at that point are
treated as skipped and are not queried against Microsoft Graph.

After usernames are resolved, the script always queries Microsoft Graph for each user's assigned
licenses and writes a friendly-named license summary into a 'License' column in the same CSV.
If the Graph connection cannot be established, the script saves the CSV with resolved usernames
only and skips license data rather than aborting.

License lookup logic lives in LicenseLookup.ps1 (same folder), which is shared with
Get-OffboardListLicenses.ps1 - update it there to change behavior for both scripts at once.

At the end, a summary table is printed to the console showing every user in the file
(not just unmatched/errored ones), along with their resolved username and license status.

.PARAMETER Path
The path to the CSV file. Defaults to C:\offboard\progress.csv.

.PARAMETER DefaultDomain
Domain appended to resolved usernames (e.g. "shastahealth.org") to form a UPN for the Graph lookup.

.EXAMPLE
.\Update-OffboardList.ps1

.EXAMPLE
.\Update-OffboardList.ps1 -Path "H:\offboard\Progress.csv"

.NOTES
Requires RSAT tools (Active Directory module) and the Microsoft.Graph.Users module.
Must be run by a user with permission to query Active Directory and a Microsoft Graph
admin role (User Administrator / License Administrator / Global Administrator).
No E5 license is required to run this script.

Prompts for manual username entry when AD matching fails; unresolved rows are skipped
for license lookup and marked accordingly in the CSV.

For an on-demand license-only check (no AD updates), use Get-UserLicenses.ps1 instead.
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$Path = "C:\offboard\progress.csv",

    [Parameter()]
    [string]$DefaultDomain = "shastahealth.org"
)

# ============================================================
# PART 1a: Resolve usernames via Active Directory
# ============================================================

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
        Write-Host "No Active Directory match found for '$name'."
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

# ============================================================
# PART 1b: Manual entry for any names that couldn't be resolved
# ============================================================
foreach ($user in $users) {
    if ([string]::IsNullOrWhiteSpace($user.Username)) {
        Write-Host "`nUnable to automatically resolve a username for '$($user.Name)'." -ForegroundColor Yellow
        $manualEntry = Read-Host "Enter the correct username, or press Enter to skip this user"

        if (![string]::IsNullOrWhiteSpace($manualEntry)) {
            $user.Username = $manualEntry.Trim()
        }
    }
}

# ============================================================
# PART 2: License lookup via Microsoft Graph (always runs)
# ============================================================

# --- Load shared license-lookup helpers ---
. "$PSScriptRoot\LicenseLookup.ps1"

if (-not (Connect-LicenseGraph)) {
    Write-Host "Username resolution completed, but license lookup could not run. Saving CSV without license data." -ForegroundColor Yellow
    $users | Export-Csv -Path $Path -NoTypeInformation
    return
}

# Cache tenant SKU list once (avoids repeated API calls per user/license)
$tenantSkus = Get-TenantSkuCache

foreach ($user in $users) {
    if ([string]::IsNullOrWhiteSpace($user.Username)) {
        $licenseValue = "Skipped - no username resolved"
    }
    else {
        $resolved = Resolve-UserLicense -Username $user.Username -DefaultDomain $DefaultDomain -TenantSkus $tenantSkus
        $licenseValue = $resolved.Licenses
    }

    # Add or overwrite the License column on this row
    if ($user.PSObject.Properties.Name -contains "License") {
        $user.License = $licenseValue
    }
    else {
        $user | Add-Member -NotePropertyName "License" -NotePropertyValue $licenseValue -Force
    }
}

# ============================================================
# PART 3: Save + summary output (all users, not just errors)
# ============================================================

$users | Export-Csv -Path $Path -NoTypeInformation

Write-Host "`n=== Offboard List Summary ===" -ForegroundColor Cyan
$users | Select-Object Name, Username, License | Format-Table -AutoSize

Write-Host "`nDone! Usernames and licenses updated in $Path" -ForegroundColor Green
