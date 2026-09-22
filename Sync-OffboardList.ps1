<#
.SYNOPSIS
Pulls open Off-board tickets from Freshservice, appends any new employees found in their
requested items to a tracking CSV, then resolves usernames via Active Directory and looks up
Microsoft 365 license assignments for every unresolved/unlicensed row.

.DESCRIPTION
This is the combined pipeline for the offboarding tracker at $Path:

1. Queries Freshservice for tickets matching $TicketStatus, $Category, and $SubCategory.
2. Pulls requested_items for each matching ticket and extracts every employee_name*
   custom field. Any Ticket+Name pair not already present in the CSV is appended as a new
   row, with all other columns left blank. Existing rows and columns are never overwritten
   or cleared.
3. Reloads the full CSV and resolves the 'Username' column via Active Directory for any row
   still missing one - exact DisplayName match, then 'Last, First', then GivenName+Surname
   concatenation, then interactive fuzzy match, then manual entry as a last resort.
4. Queries Microsoft Graph for each resolved user's assigned licenses and writes a
   friendly-named summary into the 'License' column.
5. Saves the CSV and prints a summary table of every row (Ticket, Name, Username, License).

If the Freshservice call fails, the script warns and proceeds to step 3 using whatever rows
already exist in the CSV, rather than aborting. If the Graph connection cannot be established,
the script saves usernames only and skips license data, matching the standalone script's
original behavior.

License lookup logic lives in LicenseLookup.ps1 (same folder), shared with
Get-OffboardListLicenses.ps1 - update it there to change behavior for both scripts at once.

.PARAMETER Path
The path to the CSV file. Defaults to C:\offboard\progress.csv.

.PARAMETER DefaultDomain
Domain appended to resolved usernames (e.g. "shastahealth.org") to form a UPN for the Graph lookup.

.PARAMETER FreshserviceDomain
Your Freshservice subdomain, e.g. "shastahealth" for shastahealth.freshservice.com.

.PARAMETER FreshserviceApiKey
Your Freshservice API key. Optional - if omitted, the script retrieves it from the
Microsoft.PowerShell.SecretStore vault (secret name "FreshserviceApiKey"), prompting
for it once and storing it for future runs if not already present. Passing this
parameter explicitly bypasses the vault entirely for that run.

.PARAMETER TicketStatus
Freshservice ticket status to query. Defaults to 2 (Open). Pass 3 for Pending, or run twice
with both values if off-board tickets can sit in either status before this script catches them.

.PARAMETER Category
Ticket category to filter for after pulling by status. Defaults to "User Management".

.PARAMETER SubCategory
Ticket sub-category to filter for after pulling by status. Defaults to "Off-board".

.PARAMETER DisableExpired
Switch. When set, disables the AD account and removes M365 licenses for any row whose
End Date has passed and whose 'AD User Disabled' column is still blank. Prompts for
confirmation per user (standard PowerShell ShouldProcess) unless run with -Confirm:$false.
Use -WhatIf to preview affected users with no changes made.

.EXAMPLE
.\Sync-OffboardList.ps1 -FreshserviceApiKey "abcdef123456" -FreshserviceDomain "shastahealth"

.EXAMPLE
.\Sync-OffboardList.ps1 -Path "H:\offboard\Progress.csv" -TicketStatus 3

.EXAMPLE
.\Sync-OffboardList.ps1 -DisableExpired

.NOTES
Requires RSAT tools (Active Directory module) and the Microsoft.Graph.Users module.
Must be run by a user with permission to query Active Directory and a Microsoft Graph
admin role (User Administrator / License Administrator / Global Administrator).
No E5 license is required to run this script.

category/sub_category are not filterable Freshservice API fields - the ticket list is pulled
by status only, then filtered client-side.

For an on-demand license-only check (no AD updates), use Get-UserLicenses.ps1 instead.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [Parameter()]
    [string]$Path = "C:\offboard\progress.csv",

    [Parameter()]
    [string]$DefaultDomain = "shastahealth.org",

    [Parameter()]
    [string]$FreshserviceApiKey,

    [Parameter()]
    [int]$TicketStatus = 2,

    [Parameter()]
    [string]$Category = "User Management",

    [Parameter()]
    [string]$SubCategory = "Off-board",

    [Parameter()]
    [switch]$DisableExpired
)

# ============================================================
# SecretStore setup + Freshservice API key retrieval
# ============================================================

$secretName = "FreshserviceApiKey"
$vaultName  = "SecretStore"

if ([string]::IsNullOrWhiteSpace($FreshserviceApiKey)) {

    # Ensure required modules are installed
    foreach ($moduleName in @('Microsoft.PowerShell.SecretManagement', 'Microsoft.PowerShell.SecretStore')) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            Write-Host "Installing $moduleName..." -ForegroundColor Cyan
            Install-Module -Name $moduleName -Scope CurrentUser -Force -ErrorAction Stop
        }
    }

    Import-Module Microsoft.PowerShell.SecretManagement
    Import-Module Microsoft.PowerShell.SecretStore

    # Register the SecretStore vault if it isn't already
    if (-not (Get-SecretVault -Name $vaultName -ErrorAction SilentlyContinue)) {
        Write-Host "Registering SecretStore vault..." -ForegroundColor Cyan
        Register-SecretVault -Name $vaultName -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
    }

    # First-time vault configuration - no master password, relies on Windows profile
    # protection (DPAPI) so the script can run unattended without a vault prompt each time
    try {
        Get-SecretStoreConfiguration -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "Configuring SecretStore for first use..." -ForegroundColor Cyan
        Set-SecretStoreConfiguration -Authentication None -Interaction None -Confirm:$false
    }

    # Retrieve existing key, or prompt once and store it for future runs
    $existingSecret = Get-Secret -Name $secretName -Vault $vaultName -ErrorAction SilentlyContinue

    if ($existingSecret) {
        $FreshserviceApiKey = [System.Net.NetworkCredential]::new('', $existingSecret).Password
    }
    else {
        $secureKey = Read-Host -AsSecureString -Prompt "Enter your Freshservice API key"
        Set-Secret -Name $secretName -SecureStringSecret $secureKey -Vault $vaultName
        $FreshserviceApiKey = [System.Net.NetworkCredential]::new('', $secureKey).Password
    }
}

$localTz = [System.TimeZoneInfo]::FindSystemTimeZoneById("Pacific Standard Time")

# ============================================================
# PART 1: Pull new Off-board tickets from Freshservice and append to CSV
# ============================================================

$pair = "$FreshserviceApiKey`:X"
$encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
$headers = @{ Authorization = "Basic $encoded" }

try {
    $baseUrl = "https://schc.freshservice.com/api/v2/tickets/filter"
    $queryString = "status:$TicketStatus"
    $escapedQuery = [Uri]::EscapeDataString($queryString)

    $allTickets = @()
    $page = 1

    do {
        $uri = "$baseUrl`?query=`"$escapedQuery`"&page=$page"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers
        $allTickets += $response.tickets
        $page++
    } while ($allTickets.Count -lt $response.total)

    $offboardTickets = $allTickets | Where-Object {
        $_.category -eq $Category -and $_.sub_category -eq $SubCategory
    }

    # Load existing CSV to preserve columns and detect duplicates
    $existingForDedup = @(Import-Csv -Path $Path)
    $existingKeys = @{}
    foreach ($row in $existingForDedup) {
        $existingKeys["$($row.Ticket)|$($row.Name)"] = $true
    }

    if ($existingForDedup.Count -gt 0) {
        $csvColumns = $existingForDedup[0].PSObject.Properties.Name
    }
    else {
        # Header-only CSV: read column names directly instead of relying on data rows
        $csvColumns = (Get-Content -Path $Path -TotalCount 1) -split ','
    }

    $newRows = @()

    foreach ($ticket in $offboardTickets) {
        $riUri = "https://schc.freshservice.com/api/v2/tickets/$($ticket.id)/requested_items"
        $riResponse = Invoke-RestMethod -Uri $riUri -Headers $headers

        foreach ($item in $riResponse.requested_items) {
            $cf = $item.custom_fields
            $nameProps = $cf.PSObject.Properties |
                Where-Object { $_.Name -match '^employee_name(_\d+)?$' -and $_.Value }

            foreach ($prop in $nameProps) {
                $name = $prop.Value
                $key = "$($ticket.id)|$name"

                # employee_name has no suffix but still maps to end_date_time_1;
                # employee_name_2 maps to end_date_time_2, etc.
                if ($prop.Name -match '^employee_name_(\d+)$') {
                    $dateIndex = $Matches[1]
                }
                else {
                    $dateIndex = 1
                }
                $rawEndDate = $cf."end_date_time_$dateIndex"
                if (-not [string]::IsNullOrWhiteSpace($rawEndDate)) {
                    $utcDate = ([datetime]$rawEndDate).ToUniversalTime()
                    $endDate = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcDate, $localTz).ToString("yyyy-MM-dd HH:mm:ss")
                }
                else {
                    $endDate = ""
                }

                if (-not $existingKeys.ContainsKey($key)) {
                    $row = [ordered]@{}
                    foreach ($col in $csvColumns) {
                        switch ($col) {
                            'Ticket'   { $row[$col] = $ticket.id }
                            'Name'     { $row[$col] = $name }
                            'End Date' { $row[$col] = $endDate }
                            default    { $row[$col] = '' }
                        }
                    }
                    $newRows += [PSCustomObject]$row
                    $existingKeys[$key] = $true
                }
            }
        }
    }

    if ($newRows.Count -gt 0) {
        $newRows | Export-Csv -Path $Path -Append -NoTypeInformation
        Write-Host "Added $($newRows.Count) new row(s) from Freshservice." -ForegroundColor Cyan
    }
    else {
        Write-Host "No new Freshservice entries to add." -ForegroundColor Cyan
    }
}
catch {
    Write-Warning "Freshservice pull failed: $_. Continuing with existing CSV rows only."
}

# ============================================================
# PART 2a: Resolve usernames via Active Directory
# ============================================================

Import-Module ActiveDirectory -ErrorAction Stop

try {
    $users = Import-Csv -Path $Path | Where-Object {
        $_.PSObject.Properties.Value |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }
}
catch {
    Write-Error "Failed to load CSV: $_"
    return
}

$adUserCache = Get-ADUser -Filter { GivenName -like "*" -and Surname -like "*" } -Properties GivenName, Surname, sAMAccountName

foreach ($user in $users) {
    $name = $user.Name.Trim()
    $username = $user.Username

    if (![string]::IsNullOrWhiteSpace($username)) {
        continue
    }

    $adUser = Get-ADUser -Filter { DisplayName -eq $name } -Properties sAMAccountName

    if (-not $adUser -and $name -match '^\S+\s+\S+$') {
        $parts = $name -split '\s+'
        $reversedName = "$($parts[1]), $($parts[0])"
        $adUser = Get-ADUser -Filter { DisplayName -eq $reversedName } -Properties sAMAccountName
    }

    if (-not $adUser) {
        $concatMatches = @($adUserCache | Where-Object {
            "$($_.GivenName) $($_.Surname)".Trim() -eq $name
        })
        if ($concatMatches.Count -eq 1) {
            $adUser = $concatMatches[0]
        }
    }

    if ($adUser) {
        $user.Username = $adUser.sAMAccountName
        continue
    }

    $potentialMatches = Get-ADUser -Filter { DisplayName -like "*$name*" } -Properties DisplayName, sAMAccountName

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
    }
    else {
        Write-Host "Invalid choice. Skipped."
        $user.Username = ""
    }
}

# ============================================================
# PART 2b: Manual entry for any names that couldn't be resolved
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
# PART 3: License lookup via Microsoft Graph (always runs)
# ============================================================

. "$PSScriptRoot\LicenseLookup.ps1"

if (-not (Connect-LicenseGraph)) {
    Write-Host "Username resolution completed, but license lookup could not run. Saving CSV without license data." -ForegroundColor Yellow
    $users | Export-Csv -Path $Path -NoTypeInformation
    return
}

$tenantSkus = Get-TenantSkuCache

foreach ($user in $users) {
    if ([string]::IsNullOrWhiteSpace($user.Username)) {
        $licenseValue = "Skipped - no username resolved"
    }
    else {
        $resolved = Resolve-UserLicense -Username $user.Username -DefaultDomain $DefaultDomain -TenantSkus $tenantSkus
        $licenseValue = $resolved.Licenses
    }

    if ($user.PSObject.Properties.Name -contains "License") {
        $user.License = $licenseValue
    }
    else {
        $user | Add-Member -NotePropertyName "License" -NotePropertyValue $licenseValue -Force
    }
}

# ============================================================
# PART 3.5: Disable AD account + remove licenses for expired end dates
# ============================================================

if ($DisableExpired) {
    foreach ($user in $users) {
        if ([string]::IsNullOrWhiteSpace($user.Username)) { continue }
        if ($user.'AD User Disabled' -match '\S') { continue }  # already handled

        # Check actual AD state first - catches accounts disabled outside this script
        # (manually, by another process, or before this script existed)
        try {
            $adAccount = Get-ADUser -Identity $user.Username -Properties Enabled -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not look up AD account for $($user.Name) ($($user.Username)): $_"
            continue
        }

        # Determine whether the end date has passed (needed either way, to decide on licenses)
        $endDateExpired = $false
        if (-not [string]::IsNullOrWhiteSpace($user.'End Date')) {
            $rawDateValue = $user.'End Date'
            if ($rawDateValue -isnot [string]) {
                Write-Warning "'End Date' for $($user.Name) is not a plain text value (found $($rawDateValue.GetType().Name)) — check for a duplicate 'End Date' column in the CSV. Skipping."
                continue
            }

            try {
                $endDateParsed = [datetime]::ParseExact(
                    $rawDateValue,
                    'yyyy-MM-dd HH:mm:ss',
                    [System.Globalization.CultureInfo]::InvariantCulture)
                $endDateExpired = $endDateParsed -lt (Get-Date)
            }
            catch {
                Write-Warning "Could not parse End Date '$rawDateValue' for $($user.Name) — expected format yyyy-MM-dd HH:mm:ss. Skipping."
                continue
            }
        }

        if (-not $endDateExpired) {
            # Not past its date - if already disabled some other way, just record that; no license action
            if (-not $adAccount.Enabled) {
                $user.'AD User Disabled' = "Y"
            }
            continue
        }

        # Past its end date from here on - account should end up disabled AND licenses removed,
        # regardless of whether AD disable already happened via another process
        $target = "$($user.Name) ($($user.Username))"

        if ($adAccount.Enabled) {
            if ($PSCmdlet.ShouldProcess($target, "Disable AD account (end date $($user.'End Date') has passed)")) {
                try {
                    Disable-ADAccount -Identity $user.Username -ErrorAction Stop
                    $user.'AD User Disabled' = "Y"
                }
                catch {
                    $user.'AD User Disabled' = "Error: $($_.Exception.Message)"
                    Write-Warning "Failed to disable AD account for $target : $_"
                    continue
                }
            }
            else {
                continue  # user declined the prompt; don't touch licenses either
            }
        }
        else {
            # Already disabled outside this script, but still past its date - record it
            $user.'AD User Disabled' = "Y"
        }

        # Remove licenses - runs whenever end date has passed, independent of who disabled AD
        if ($PSCmdlet.ShouldProcess($target, "Remove M365 licenses (end date $($user.'End Date') has passed)")) {
            try {
                $upn = "$($user.Username)@$DefaultDomain"
                $skuIds = (Get-MgUserLicenseDetail -UserId $upn -ErrorAction Stop).SkuId
                if ($skuIds) {
                    Set-MgUserLicense -UserId $upn -AddLicenses @() -RemoveLicenses $skuIds -ErrorAction Stop
                    $user.License = "Removed $(Get-Date -Format 'yyyy-MM-dd')"
                }
            }
            catch {
                Write-Warning "Failed to remove licenses for $target : $_"
            }
        }
    }
}

# ============================================================
# PART 4: Save + summary output (all users, not just errors)
# ============================================================

$users | Export-Csv -Path $Path -NoTypeInformation

Write-Host "`n=== Offboard List Summary ===" -ForegroundColor Cyan
$users | Select-Object Ticket, Name, Username, 'End Date', 'AD User Disabled', License | Format-Table -AutoSize

Write-Host "`nDone! Tickets, usernames, and licenses updated in $Path" -ForegroundColor Green