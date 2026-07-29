<#
.SYNOPSIS
    Retrieves assigned Microsoft 365 licenses for a list of users via Microsoft Graph.

.DESCRIPTION
    Reads a list of usernames/UPNs (from a text file, CSV, or inline array),
    queries Microsoft Graph for each user's assigned licenses, and outputs a
    friendly summary table.

    Shared lookup logic (friendly SKU names, Graph resolution strategy) lives in
    LicenseLookup.ps1, which must be in the same folder as this script. It is
    also used by Update-OffboardList.ps1 - edit LicenseLookup.ps1 to change
    behavior for both scripts at once.

.NOTES
    Requires: Microsoft.Graph.Users module
        Install-Module Microsoft.Graph.Users -Scope CurrentUser
    Requires delegated or app permission: User.Read.All (or User.ReadBasic.All + LicenseAssignment)
    Any account with User Administrator / License Administrator / Global Administrator role can run this.
    No E5 license required to run this script.

.EXAMPLE
    .\Get-OffboardListLicenses.ps1
    (defaults to reading the "Username" column from C:\offboard\progress.csv)

.EXAMPLE
    .\Get-OffboardListLicenses.ps1 -InputFile .\users.txt -OutputFile .\LicenseReport.csv
#>

param(
    # Path to a CSV file containing a "Username" column. Defaults to the offboard progress file.
    [string]$CsvFile = "C:\offboard\progress.csv",

    # Path to a plain text file with one UPN/username per line (overrides CsvFile if provided)
    [string]$InputFile,

    # Or pass usernames directly (overrides CsvFile if provided)
    [string[]]$UserList,

    # Optional CSV export path
    [string]$OutputFile,

    # Domain to append to bare usernames (e.g. "shastahealth.org") to form a UPN, e.g. atippens -> atippens@shastahealth.org
    # Leave blank to skip this and rely on the onPremisesSamAccountName filter fallback instead.
    [string]$DefaultDomain = "shastahealth.org"
)

# --- Load shared license-lookup helpers ---
. "$PSScriptRoot\LicenseLookup.ps1"

# --- Connect to Graph ---
if (-not (Connect-LicenseGraph)) {
    return
}

# --- Build the list of users to check ---
$users = @()
if ($InputFile) {
    if (-not (Test-Path $InputFile)) {
        Write-Error "Input file not found: $InputFile"
        return
    }
    $users = Get-Content $InputFile | Where-Object { $_.Trim() -ne "" }
}
elseif ($UserList) {
    $users = $UserList
}
else {
    # Default: read the "Username" column from the CSV file
    if (-not (Test-Path $CsvFile)) {
        Write-Error "CSV file not found: $CsvFile"
        return
    }

    $csvData = Import-Csv -Path $CsvFile

    if (-not ($csvData | Get-Member -Name "Username" -MemberType NoteProperty)) {
        Write-Error "No 'Username' column found in $CsvFile. Columns present: $(($csvData | Get-Member -MemberType NoteProperty).Name -join ', ')"
        return
    }

    $users = $csvData |
    Select-Object -ExpandProperty Username |
    Where-Object { $_ -and $_.Trim() -ne "" }

    if ($users.Count -eq 0) {
        Write-Error "No usernames found in the 'Username' column of $CsvFile."
        return
    }

    Write-Host "Loaded $($users.Count) username(s) from $CsvFile" -ForegroundColor Cyan
}

# --- Cache tenant SKU list once (avoids repeated API calls per user/license) ---
$tenantSkus = Get-TenantSkuCache

# --- Query each user ---
$results = foreach ($u in $users) {
    Resolve-UserLicense -Username $u.Trim() -DefaultDomain $DefaultDomain -TenantSkus $tenantSkus
}

# --- Output ---
$results | Format-Table -AutoSize

if ($OutputFile) {
    $results | Export-Csv -Path $OutputFile -NoTypeInformation
    Write-Host "`nExported to $OutputFile" -ForegroundColor Green
}