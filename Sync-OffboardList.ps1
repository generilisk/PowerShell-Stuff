<#
.SYNOPSIS
Pulls open Off-board tickets from Freshservice, appends any new employees found in their
requested items to a tracking CSV, resolves usernames via Active Directory, and (with
-DisableExpired) carries out the offboarding steps for everyone past their end date.

.DESCRIPTION
If $Path does not exist, the script creates it (and its containing folder, if needed)
with the standard header row before proceeding.

This is the combined pipeline for the offboarding tracker at $Path:

1. Queries Freshservice for tickets matching $TicketStatus, $Category, and $SubCategory.
2. Pulls requested_items for each matching ticket and extracts every employee_name*
   custom field. Any Ticket+Name pair not already present in the CSV is appended as a new
   row, with all other columns left blank. Existing rows and columns are never overwritten
   or cleared.
3. Reloads the full CSV and resolves the 'Username' column via Active Directory for any row
   still missing one - exact DisplayName match, then 'Last, First', then GivenName+Surname
   concatenation, then interactive fuzzy match, then manual entry as a last resort.
4. Past each user's end date (-DisableExpired only), carries out the offboarding steps, each
   one recorded in its own CSV column:
   - Disables the AD account ('AD User Disabled').
   - Licenses: asks for confirmation, then posts the user's M365 license list to the ticket as
     a private note ('Licenses Noted'), and only if the note posts removes the licenses
     ('Licenses Removed').
   - Groups: same pattern - asks, posts the group list as a private note so whoever takes over
     the role can see which groups to assign ('Groups Noted'), and only if the note posts
     removes the AD group memberships ('Groups Removed'). The primary group is left alone.
   - Files: asks, then runs Remove-OffboardUserData.ps1 under an admin account you are prompted
     for (backs up and removes the profile folders and H Drive), and records the date in
     'Profile Folder Work' only if the HR result file shows success.
   Anything declined or failed stays blank and is asked again on the next run.
5. Saves the CSV and prints a summary table of every row.

The CSV columns are kept in workflow order. The last two ('AD User
Deleted', 'Ticket Closed') are filled in by hand or by other scripts: the AD account is deleted
once the user's files are backed up, and the ticket is closed last.

If the Freshservice call fails, the script warns and proceeds to step 3 using whatever rows
already exist in the CSV, rather than aborting - but -DisableExpired is skipped for that run,
because ticket data can't be trusted. -DisableExpired is also skipped if Microsoft Graph
can't be reached.

Graph connection and license name translation come from LicenseHelpers.psm1 (same folder),
shared with Get-OffboardListLicenses.ps1.

.PARAMETER Path
The path to the CSV file. Defaults to C:\offboard\progress.csv.

.PARAMETER DefaultDomain
Domain appended to resolved usernames (e.g. "shastahealth.org") to form a UPN for the Graph lookup.

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
Switch. When set, for any row whose End Date has passed: disables the AD account, then
removes the M365 licenses and the AD group memberships (except the primary group, e.g. Domain
Users). Each step prompts for confirmation per user (standard PowerShell ShouldProcess) unless
run with -Confirm:$false. On yes, the license or group list is first posted to the ticket as a
private note, and the removal only happens if that note posts successfully.
Use -WhatIf to preview affected users with no changes made.
Each AD disable, license removal, and group removal also adds a private note to the user's
ticket (what was done, when, and by whom).
After those, runs the user file cleanup (Remove-OffboardUserData.ps1) under an admin account you
are prompted for, and records the date in 'Profile Folder Work' if it succeeds.
If the Freshservice pull fails, or Microsoft Graph can't be reached, -DisableExpired is
skipped for that run (nothing is disabled or removed).

.PARAMETER HrArchiveRoot
The HR folder where Remove-OffboardUserData.ps1 archives each user's files. After the cleanup
runs, the script reads <HrArchiveRoot>\<username>\Termination.result to confirm it succeeded.
Must match $hrRoot in Remove-OffboardUserData.ps1.

.EXAMPLE
.\Sync-OffboardList.ps1 -Path "H:\offboard\Progress.csv" -TicketStatus 3

.EXAMPLE
.\Sync-OffboardList.ps1 -DisableExpired

.NOTES
On first run the script sets up its own prerequisites: the required PowerShell modules
are installed to the current user's profile, and the RSAT Active Directory tools are
installed if missing (that one step needs an elevated PowerShell session). Needs
PowerShell Gallery access, and the Freshservice API key (prompted for once, then saved).
Must be run by a user with permission to query Active Directory and a Microsoft Graph
admin role (User Administrator / License Administrator / Global Administrator).
No E5 license is required to run this script.

category/sub_category are not filterable Freshservice API fields - the ticket list is pulled
by status only, then filtered client-side.

Microsoft Graph is only contacted when -DisableExpired is used. For an on-demand license-only
check (no AD updates), use Get-OffboardListLicenses.ps1 instead.
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
    [switch]$DisableExpired,

    [Parameter()]
    [string]$HrArchiveRoot = "\\shared01\hr\_HR Shared\TERMINATION\Termed employee files"
)


# ============================================================
# PART 1: First-run prerequisites (skips anything already present)
# ============================================================

# Active Directory PowerShell module (RSAT). This is a Windows feature, so adding it
# needs an elevated session; if that's missing we stop with clear instructions.
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Error ("The Active Directory PowerShell module (RSAT) is not installed. Run this script once from an " +
            "elevated PowerShell window (Run as administrator) and it will install it, or install it yourself with: " +
            "Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'")
        return
    }
    Write-Host "Installing RSAT Active Directory tools (this can take a few minutes)..." -ForegroundColor Cyan
    try {
        Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -WhatIf:$false -Confirm:$false -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Error "Failed to install RSAT Active Directory tools: $_"
        return
    }
}

# Shared license helper module (must be in the same folder as this script)
Import-Module "$PSScriptRoot\LicenseHelpers.psm1" -Force -ErrorAction Stop

# Microsoft Graph modules used by the license lookup/removal
if (-not (Install-LicenseGraphModule)) { return }

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
# Ensure CSV and its folder exist
# ============================================================

$csvFolder = Split-Path -Path $Path -Parent
if ($csvFolder -and -not (Test-Path -Path $csvFolder)) {
    Write-Host "Creating folder: $csvFolder" -ForegroundColor Cyan
    New-Item -Path $csvFolder -ItemType Directory -Force | Out-Null
}

if (-not (Test-Path -Path $Path)) {
    Write-Host "Creating new CSV: $Path" -ForegroundColor Cyan
    $header = "Ticket,Name,Username,End Date,AD User Disabled,Licenses Noted,Licenses Removed,Groups Noted,Groups Removed,Profile Folder Work,AD User Deleted,Ticket Closed"
    Set-Content -Path $Path -Value $header -Encoding UTF8
}

# ============================================================
# PART 2: Pull new Off-board tickets from Freshservice and append to CSV
# ============================================================

$freshserviceOk = $false   # set to $true only if the whole Freshservice pull succeeds; gates -DisableExpired

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
    $freshserviceOk = $true
}
catch {
    Write-Warning "Freshservice pull failed: $_. Continuing with existing CSV rows only."
    if ($DisableExpired) {
        Write-Warning "-DisableExpired will be skipped this run: with Freshservice unreachable, no accounts, licenses, or groups will be changed."
    }
}

# ============================================================
# PART 3: Resolve usernames via Active Directory
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

# Properties used to tell people apart in the pick lists below
$adProps = 'GivenName', 'Surname', 'DisplayName', 'sAMAccountName', 'Title', 'Department', 'Enabled'
$adUserCache = Get-ADUser -Filter { GivenName -like "*" -and Surname -like "*" } -Properties $adProps

# Helper: one-line description of an AD account, so similar names can be told apart
function Format-AdUserChoice {
    param($AdUser)
    $details = @($AdUser.Title, $AdUser.Department) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $status  = if ($AdUser.Enabled) { 'enabled' } else { 'DISABLED' }
    $text    = "$($AdUser.DisplayName) [$($AdUser.sAMAccountName)]"
    if ($details) { $text += " - $($details -join ', ')" }
    return "$text - $status"
}

# Helper: show a numbered list and let the operator pick one.
# Returns the chosen AD account, or $null if they skip or enter something invalid.
function Select-AdUserMatch {
    param([string]$Name, $Candidates, [string]$Heading)
    Write-Host "`n$Heading '$Name':" -ForegroundColor Yellow
    $i = 1
    foreach ($candidate in $Candidates) {
        Write-Host "$i. $(Format-AdUserChoice $candidate)"
        $i++
    }
    $choice = Read-Host "Enter the number of the correct user, or press Enter to skip"
    if ([string]::IsNullOrWhiteSpace($choice) -or $choice -notmatch '^\d+$') {
        Write-Host "Skipped."
        return $null
    }
    $index = [int]$choice - 1
    if ($index -ge 0 -and $index -lt $Candidates.Count) { return $Candidates[$index] }
    Write-Host "Invalid choice. Skipped."
    return $null
}

# Helper: does a ticket name ("First Last") match an AD DisplayName ("First Last" or "Last, First")?
function Test-NameMatchesDisplayName {
    param([string]$TicketName, [string]$DisplayName)
    $a = ($TicketName  -replace '\s+', ' ').Trim()
    $b = ($DisplayName -replace '\s+', ' ').Trim()
    if ($a -ieq $b) { return $true }
    if ($a -match '^(\S+) (\S+)$' -and $b -ieq "$($Matches[2]), $($Matches[1])") { return $true }
    return $false
}

foreach ($user in $users) {
    $name = $user.Name.Trim()
    $username = $user.Username

    if (![string]::IsNullOrWhiteSpace($username)) {
        continue
    }

    # Look for exact matches: DisplayName, then "Last, First", then GivenName + Surname.
    # Exactly one match is used automatically. MORE than one is ambiguous, so a person always picks:
    # guessing wrong here would put the wrong person's groups on the ticket and could disable them.
    $exactMatches = @(Get-ADUser -Filter { DisplayName -eq $name } -Properties $adProps)

    if ($exactMatches.Count -eq 0 -and $name -match '^\S+\s+\S+$') {
        $parts = $name -split '\s+'
        $reversedName = "$($parts[1]), $($parts[0])"
        $exactMatches = @(Get-ADUser -Filter { DisplayName -eq $reversedName } -Properties $adProps)
    }

    if ($exactMatches.Count -eq 0) {
        $exactMatches = @($adUserCache | Where-Object {
            "$($_.GivenName) $($_.Surname)".Trim() -eq $name
        })
    }

    if ($exactMatches.Count -eq 1) {
        $user.Username = $exactMatches[0].sAMAccountName
        continue
    }

    if ($exactMatches.Count -gt 1) {
        $selected = Select-AdUserMatch -Name $name -Candidates $exactMatches -Heading "More than one Active Directory account matches"
        $user.Username = if ($selected) { $selected.sAMAccountName } else { "" }
        continue
    }

    # No exact match: offer partial matches
    $potentialMatches = @(Get-ADUser -Filter { DisplayName -like "*$name*" } -Properties $adProps)

    if ($potentialMatches.Count -eq 0) {
        Write-Host "No Active Directory match found for '$name'."
        $user.Username = ""
        continue
    }

    $selected = Select-AdUserMatch -Name $name -Candidates $potentialMatches -Heading "Possible matches for"
    $user.Username = if ($selected) { $selected.sAMAccountName } else { "" }
}

# ============================================================
# PART 4: Manual entry for any names that couldn't be resolved
# ============================================================
# A typed username is checked against AD, and the operator confirms the account it belongs to
# before it is used. A typo that happens to be a real account would otherwise point every later
# step (notes, disable, license and group removal) at the wrong person.
foreach ($user in $users) {
    if (-not [string]::IsNullOrWhiteSpace($user.Username)) { continue }

    Write-Host "`nUnable to automatically resolve a username for '$($user.Name)'." -ForegroundColor Yellow

    while ($true) {
        $manualEntry = Read-Host "Enter the correct username, or press Enter to skip this user"
        if ([string]::IsNullOrWhiteSpace($manualEntry)) { break }

        try {
            $found = Get-ADUser -Identity $manualEntry.Trim() -Properties $adProps -ErrorAction Stop
        }
        catch {
            Write-Warning "'$($manualEntry.Trim())' was not found in Active Directory. Check the spelling and try again."
            continue
        }

        Write-Host "  Found: $(Format-AdUserChoice $found)"
        if (-not (Test-NameMatchesDisplayName -TicketName $user.Name -DisplayName $found.DisplayName)) {
            Write-Host "  Note: this name does not match the name on the ticket ('$($user.Name)')." -ForegroundColor Yellow
        }

        $confirm = Read-Host "Use this account for '$($user.Name)'? (y/n)"
        if ($confirm -match '^(y|yes)$') {
            $user.Username = $found.sAMAccountName
            break
        }
    }
}

# Make sure every row has the 'Groups Noted', 'Groups Removed', 'Licenses Noted' and
# 'Licenses Removed' columns (older CSVs won't)
foreach ($user in $users) {
    foreach ($column in 'Groups Noted', 'Groups Removed', 'Licenses Noted', 'Licenses Removed') {
        if ($user.PSObject.Properties.Name -notcontains $column) {
            $user | Add-Member -NotePropertyName $column -NotePropertyValue '' -Force
        }
    }
}

# Put columns in workflow order. Any extra columns someone added by hand are kept, after the standard ones.
# The file itself is rewritten in this order at the next save.
$columnOrder = 'Ticket', 'Name', 'Username', 'End Date', 'AD User Disabled',
               'Licenses Noted', 'Licenses Removed', 'Groups Noted', 'Groups Removed',
               'Profile Folder Work', 'AD User Deleted', 'Ticket Closed'
$users = @($users)
if ($users.Count -gt 0) {
    $extraColumns = @($users[0].PSObject.Properties.Name | Where-Object { $_ -notin $columnOrder })
    $users = @($users | Select-Object -Property ($columnOrder + $extraColumns))
}

# Helper: add a private note to a Freshservice ticket. Throws on failure so each caller decides
# how to report it. Notes are HTML (plain line breaks are ignored), so callers build the body
# with <br>. Sent as JSON (works in Windows PowerShell 5.1 and 7; -Form needs PowerShell 7).
# $headers already holds the API authorization from Part 2.
function Add-TicketNote {
    param(
        [Parameter(Mandatory)][string]$Ticket,
        [Parameter(Mandatory)][string]$Body
    )
    $noteUri = "https://schc.freshservice.com/api/v2/tickets/$Ticket/notes"
    $payload = @{ body = $Body; private = $true } | ConvertTo-Json
    $null = Invoke-RestMethod -Uri $noteUri -Method Post -Headers $headers `
        -ContentType 'application/json; charset=utf-8' -Body $payload -ErrorAction Stop
}

# Helper: save the CSV now instead of waiting for Part 8. Used as a checkpoint so that notes already
# posted and accounts already changed are remembered even if the run is interrupted. Skipped under
# -WhatIf, and a save failure (e.g. the CSV is open in Excel) only warns.
function Save-OffboardCsv {
    if ($WhatIfPreference) { return }
    try { $users | Export-Csv -Path $Path -NoTypeInformation -ErrorAction Stop }
    catch { Write-Warning "Could not save progress to $Path (is it open in another program?): $($_.Exception.Message)" }
}

# Helper: converts S:\foo to \\server\share\foo when S: is a mapped drive. The admin account's
# logon session doesn't have this user's mapped drives, so the child must be given a UNC path.
function Resolve-UncPath {
    param([string]$Path)
    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    $root = [System.IO.Path]::GetPathRoot($full)
    if ($root -match '^([A-Za-z]):\\$') {
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($Matches[1]):'"
        if ($disk -and $disk.ProviderName) {
            return (Join-Path $disk.ProviderName $full.Substring($root.Length))
        }
    }
    return $full
}

# Checkpoint: keep the resolved usernames and the new column order even if the run is interrupted later
Save-OffboardCsv

# ============================================================
# PART 5: Microsoft Graph connection (only needed for -DisableExpired)
# ============================================================

# Destructive actions only run when the Freshservice pull fully succeeded and Graph is reachable
$runDisable = $DisableExpired.IsPresent -and $freshserviceOk

if ($runDisable -and -not (Connect-LicenseGraph -IncludeWrite)) {
    Write-Warning "Could not connect to Microsoft Graph, so -DisableExpired was skipped this run (nothing was disabled or removed)."
    $runDisable = $false
}

# ============================================================
# PART 6: Disable AD account, remove licenses and groups for expired end dates
# ============================================================

if ($runDisable) {
    # try/finally so the CSV is saved even if the run is interrupted (Ctrl+C at a prompt, an error).
    # Without this, accounts already disabled or groups already removed would be forgotten.
    try {
        foreach ($user in $users) {
            if ([string]::IsNullOrWhiteSpace($user.Username)) { continue }
            if ($user.'AD User Deleted' -match '\S') { continue }  # account is gone, nothing left to do

            # Skip rows that are fully handled (disabled AND licenses removed AND groups removed). A row an
            # earlier run disabled that is still waiting on licenses or groups falls through. A failed disable
            # ("Error: ...") stays skipped so it gets looked at manually.
            $disabledState = $user.'AD User Disabled'
            if ($disabledState -match '\S' -and $disabledState -ne 'Y') { continue }
            if ($disabledState -eq 'Y' -and $user.'Groups Removed' -match '\S' -and
                $user.'Licenses Removed' -match '\S') { continue }

            # Check actual AD state first - catches accounts disabled outside this script
            # (manually, by another process, or before this script existed)
            try {
                $adAccount = Get-ADUser -Identity $user.Username -Properties Enabled, PrimaryGroup -ErrorAction Stop
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
                    Write-Warning "'End Date' for $($user.Name) is not a plain text value (found $($rawDateValue.GetType().Name)) - check for a duplicate 'End Date' column in the CSV. Skipping."
                    continue
                }

                $knownFormats = @(
                    'yyyy-MM-dd HH:mm:ss',   # format this script writes
                    'M/d/yyyy H:mm',          # Excel short-date reformat (no seconds, no leading zeros)
                    'M/d/yyyy HH:mm:ss'       # Excel reformat if seconds survive
                )

                $endDateParsed = $null
                foreach ($fmt in $knownFormats) {
                    try {
                        $endDateParsed = [datetime]::ParseExact($rawDateValue, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
                        break
                    }
                    catch {
                        # try next format
                    }
                }

                if (-not $endDateParsed) {
                    Write-Warning "Could not parse End Date '$rawDateValue' for $($user.Name) - unrecognized format. Skipping."
                    continue
                }

                $endDateExpired = $endDateParsed -lt (Get-Date)
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

                    # Record the disable on the ticket. A failed note doesn't undo the disable, so just warn.
                    if (-not [string]::IsNullOrWhiteSpace($user.Ticket)) {
                        $noteBody = "AD account disabled for $([System.Net.WebUtility]::HtmlEncode($user.Name)) ($($user.Username)).<br><br>" +
                                    "Disabled: $(Get-Date -Format 'yyyy-MM-dd HH:mm') by $env:USERNAME<br>" +
                                    "End date: $($user.'End Date')<br>"
                        try { Add-TicketNote -Ticket $user.Ticket -Body $noteBody }
                        catch { Write-Warning "AD account for $target was disabled, but adding the ticket note failed: $($_.Exception.Message)" }
                    }
                }
                elseif (-not $WhatIfPreference) {
                    continue  # user declined the prompt; don't touch licenses or groups either
                }
                # Under -WhatIf, fall through so the license and group steps are previewed too
            }
            else {
                # Already disabled outside this script, but still past its date - record it
                $user.'AD User Disabled' = "Y"
            }

            # Licenses - ask, then note, then remove. The license list goes on the ticket only after the
            # operator says yes, and nothing is removed unless that note posted. Declined or failed rows
            # stay blank and are asked again on the next run. A user with no licenses has nothing to ask
            # about: a "No licenses assigned" note is posted and both columns are filled.
            if ($user.'Licenses Removed' -notmatch '\S') {
                try {
                    $upn = "$($user.Username)@$DefaultDomain"
                    $licenseDetail = @(Get-MgUserLicenseDetail -UserId $upn -ErrorAction Stop)
                    $licenseNames  = @($licenseDetail | ForEach-Object { Get-FriendlyName $_.SkuPartNumber } | Sort-Object)
                    if ($licenseNames.Count -ne $licenseDetail.Count) {
                        throw "could not translate every license to a friendly name"
                    }

                    if ($licenseDetail.Count -eq 0) {
                        if ($WhatIfPreference) {
                            Write-Host "What if: $target has no M365 licenses; a 'No licenses assigned' note would be added to ticket $($user.Ticket)"
                        }
                        else {
                            if ([string]::IsNullOrWhiteSpace($user.Ticket)) { throw "this row has no ticket number to record the license list on" }
                            $noteBody = "License List for $([System.Net.WebUtility]::HtmlEncode($user.Name)) ($($user.Username)):<br><br>No licenses assigned.<br>"
                            Add-TicketNote -Ticket $user.Ticket -Body $noteBody
                            $user.'Licenses Noted'   = Get-Date -Format 'yyyy-MM-dd'
                            $user.'Licenses Removed' = Get-Date -Format 'yyyy-MM-dd'   # nothing to remove
                            Save-OffboardCsv
                        }
                    }
                    elseif ($PSCmdlet.ShouldProcess($target, "Remove M365 licenses (end date $($user.'End Date') has passed)")) {
                        # Step 1: on yes, the list goes on the ticket first (skipped if an earlier run already posted it)
                        if ($user.'Licenses Noted' -notmatch '\S') {
                            if ([string]::IsNullOrWhiteSpace($user.Ticket)) { throw "this row has no ticket number to record the license list on" }
                            $noteBody = "License List for $([System.Net.WebUtility]::HtmlEncode($user.Name)) ($($user.Username)):<br><br>"
                            foreach ($licenseName in $licenseNames) {
                                $noteBody += [System.Net.WebUtility]::HtmlEncode($licenseName) + ";<br>"
                            }
                            Add-TicketNote -Ticket $user.Ticket -Body $noteBody
                            $user.'Licenses Noted' = Get-Date -Format 'yyyy-MM-dd'
                            Save-OffboardCsv   # checkpoint so the note isn't posted twice if the run is interrupted
                        }

                        # Step 2: remove (a failed note throws above, so this is only reached once the list is on the ticket)
                        $null = Set-MgUserLicense -UserId $upn -AddLicenses @() -RemoveLicenses @($licenseDetail.SkuId) -ErrorAction Stop
                        $user.'Licenses Removed' = Get-Date -Format 'yyyy-MM-dd'

                        # Short confirmation note. A failed note doesn't undo the removal, so just warn.
                        $noteBody = "M365 licenses removed for $([System.Net.WebUtility]::HtmlEncode($user.Name)) ($($user.Username)).<br><br>" +
                                    "Removed $($licenseNames.Count) license(s) on $(Get-Date -Format 'yyyy-MM-dd HH:mm') by $env:USERNAME. " +
                                    "The License List note on this ticket shows what they had.<br>"
                        try { Add-TicketNote -Ticket $user.Ticket -Body $noteBody }
                        catch { Write-Warning "Licenses for $target were removed, but adding the ticket note failed: $($_.Exception.Message)" }
                    }
                }
                catch {
                    Write-Warning "License handling did not complete for $target (if the ticket note failed, nothing was removed; it will be retried on the next run): $_"
                }
            }

            # Groups - same pattern: ask, then note, then remove. Disabling an account doesn't take it out of
            # its groups, so without this the user still shows up in mail lists and security groups. The
            # primary group (Domain Users) can't be removed, so it is left out of the list and left alone.
            # Declined or failed rows stay blank and are asked again on the next run.
            if ($user.'Groups Removed' -notmatch '\S') {
                try {
                    $memberOf = @(Get-ADPrincipalGroupMembership -Identity $user.Username -ErrorAction Stop |
                        Where-Object { $_.DistinguishedName -ne $adAccount.PrimaryGroup } |
                        Sort-Object Name)

                    if ($memberOf.Count -eq 0) {
                        if ($WhatIfPreference) {
                            Write-Host "What if: $target has no removable group memberships; a note saying so would be added to ticket $($user.Ticket)"
                        }
                        else {
                            if ([string]::IsNullOrWhiteSpace($user.Ticket)) { throw "this row has no ticket number to record the group list on" }
                            $noteBody = "Group List for $([System.Net.WebUtility]::HtmlEncode($user.Name)) ($($user.Username)):<br><br>No removable group memberships found.<br>"
                            Add-TicketNote -Ticket $user.Ticket -Body $noteBody
                            $user.'Groups Noted'   = Get-Date -Format 'yyyy-MM-dd'
                            $user.'Groups Removed' = Get-Date -Format 'yyyy-MM-dd'   # nothing to remove
                            Save-OffboardCsv
                        }
                    }
                    elseif ($PSCmdlet.ShouldProcess($target, "Remove AD group memberships (end date $($user.'End Date') has passed)")) {
                        # Step 1: on yes, the list goes on the ticket first (skipped if an earlier run already posted it).
                        # Notes are HTML, so each group goes on its own line with <br> and names are HTML-encoded.
                        if ($user.'Groups Noted' -notmatch '\S') {
                            if ([string]::IsNullOrWhiteSpace($user.Ticket)) { throw "this row has no ticket number to record the group list on" }
                            $noteBody = "Group List for $([System.Net.WebUtility]::HtmlEncode($user.Name)) ($($user.Username)):<br><br>"
                            foreach ($group in $memberOf) {
                                $noteBody += [System.Net.WebUtility]::HtmlEncode($group.Name) + ";<br>"
                            }
                            Add-TicketNote -Ticket $user.Ticket -Body $noteBody
                            $user.'Groups Noted' = Get-Date -Format 'yyyy-MM-dd'
                            Save-OffboardCsv   # checkpoint so the note isn't posted twice if the run is interrupted
                        }

                        # Step 2: remove (a failed note throws above, so this is only reached once the list is on the ticket).
                        # One group at a time so a single failure (e.g. a protected group) doesn't stop the rest.
                        $removedGroups = @()
                        $failedGroups  = @()
                        foreach ($group in $memberOf) {
                            try {
                                Remove-ADGroupMember -Identity $group -Members $user.Username -Confirm:$false -ErrorAction Stop
                                $removedGroups += $group.Name
                            }
                            catch {
                                $failedGroups += $group.Name
                                Write-Warning "Could not remove $target from group '$($group.Name)': $($_.Exception.Message)"
                            }
                        }

                        # Only mark the row done when nothing failed; otherwise the remaining groups are retried next run
                        if ($failedGroups.Count -eq 0) { $user.'Groups Removed' = Get-Date -Format 'yyyy-MM-dd' }

                        # Confirmation note (the full list was posted in step 1). Only when something was actually
                        # removed, so a group that can never be removed doesn't add a note on every run.
                        if ($removedGroups.Count -gt 0) {
                            $noteBody = "AD group memberships removed for $([System.Net.WebUtility]::HtmlEncode($user.Name)) ($($user.Username)).<br><br>" +
                                        "Removed from $($removedGroups.Count) group(s) on $(Get-Date -Format 'yyyy-MM-dd HH:mm') by $env:USERNAME. " +
                                        "The Group List note on this ticket shows what they were in.<br>"
                            if ($failedGroups.Count -gt 0) {
                                $noteBody += "<br>Could not remove from these groups (will be retried on the next run):<br>"
                                foreach ($groupName in $failedGroups) {
                                    $noteBody += [System.Net.WebUtility]::HtmlEncode($groupName) + ";<br>"
                                }
                            }
                            try { Add-TicketNote -Ticket $user.Ticket -Body $noteBody }
                            catch { Write-Warning "Groups for $target were processed, but adding the ticket note failed: $($_.Exception.Message)" }
                        }
                    }
                }
                catch {
                    Write-Warning "Group handling did not complete for $target (if the ticket note failed, nothing was removed; it will be retried on the next run): $_"
                }
            }
        }
    }
    finally {
        Save-OffboardCsv
    }
}

# ============================================================
# PART 7: User file cleanup (runs Remove-OffboardUserData.ps1 under the admin account)
# ============================================================

if ($runDisable) {
    try {
        # Ready for file cleanup: disabled, licenses and groups handled, files not yet done
        $cleanupRows = @($users | Where-Object {
            $_.Username -match '\S' -and
            $_.'AD User Disabled' -eq 'Y' -and
            $_.'Licenses Removed' -match '\S' -and
            $_.'Groups Removed' -match '\S' -and
            $_.'Profile Folder Work' -notmatch '\S' -and
            $_.'AD User Deleted' -notmatch '\S'
        })

        # Ask per user first, so the admin credential is only requested if there is work to do
        $approved = @()
        foreach ($user in $cleanupRows) {
            $target = "$($user.Name) ($($user.Username))"
            if ($PSCmdlet.ShouldProcess($target, "Back up and delete profile folders and H Drive")) {
                $approved += $user
            }
        }

        if ($approved.Count -gt 0) {
            $cleanupScript = Join-Path (Resolve-UncPath $PSScriptRoot) 'Remove-OffboardUserData.ps1'
            if (-not (Test-Path -LiteralPath $cleanupScript)) { throw "Cannot find $cleanupScript" }

            $adminCred = Get-Credential -Message 'Admin account for the file cleanup' -UserName "$env:USERDOMAIN\"
            if (-not $adminCred) { throw 'No admin credential supplied.' }

            foreach ($user in $approved) {
                $target   = "$($user.Name) ($($user.Username))"
                $launched = (Get-Date).AddSeconds(-1)
                $argList  = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$cleanupScript`"",
                              '-UserName', $user.Username, '-Unattended')
                try {
                    Write-Host "Running file cleanup for $target ..."
                    # The exit code isn't reliable with -Credential; the result file is the signal.
                    $null = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Credential $adminCred `
                        -WorkingDirectory "$env:SystemRoot\System32" -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
                }
                catch {
                    # Stop here rather than retrying a rejected credential for every remaining user
                    Write-Warning "Could not start the file cleanup for ${target}: $($_.Exception.Message). Remaining users were skipped."
                    break
                }

                # Trust only a result file written after this launch, for this user, that says Success
                $userFolder = Join-Path $HrArchiveRoot $user.Username
                $resultFile = Join-Path $userFolder 'Termination.result'
                $status = $null
                if (Test-Path -LiteralPath $resultFile) {
                    $parts = @((Get-Content -LiteralPath $resultFile -TotalCount 1) -split ',')
                    $stamp = [datetime]::MinValue
                    if ($parts.Count -ge 3 -and $parts[0] -eq $user.Username -and
                        [datetime]::TryParse($parts[2], [ref]$stamp) -and $stamp -ge $launched) {
                        $status = $parts[1]
                    }
                }

                if ($status -eq 'Success') {
                    $user.'Profile Folder Work' = Get-Date -Format 'yyyy-MM-dd'
                    Save-OffboardCsv
                    try {
                        $noteBody = "User files processed for $([System.Net.WebUtility]::HtmlEncode($user.Name)) ($($user.Username)).<br><br>" +
                                    "Profile folders and H Drive were backed up to the HR termination folder where needed and removed from the network shares on $(Get-Date -Format 'yyyy-MM-dd HH:mm') by $env:USERNAME.<br>"
                        Add-TicketNote -Ticket $user.Ticket -Body $noteBody
                    }
                    catch { Write-Warning "File cleanup for $target succeeded, but adding the ticket note failed: $($_.Exception.Message)" }
                }
                else {
                    Write-Warning "File cleanup did not report success for $target. 'Profile Folder Work' left blank; check the newest *-Termination.log in $userFolder"
                }
            }
        }
    }
    catch {
        Write-Warning "File cleanup step did not complete: $_"
    }
    finally {
        Save-OffboardCsv
    }
}

# ============================================================
# PART 8: Save + summary output (all users, not just errors)
# ============================================================

$users | Export-Csv -Path $Path -NoTypeInformation

Write-Host "`n=== Offboard List Summary ===" -ForegroundColor Cyan
$users | Select-Object -Property $columnOrder | Format-Table -AutoSize -Wrap

Write-Host "`nDone! Tickets, usernames, and offboarding progress updated in $Path" -ForegroundColor Green