<#
.SYNOPSIS
Pulls open Off-board tickets from Freshservice, appends any new employees found in their
requested items to a tracking CSV, takes each username from the ticket and verifies it in
Active Directory, and (with -DisableExpired) carries out the offboarding steps for everyone
past their end date.

.DESCRIPTION
If $Path does not exist, the script creates it (and its containing folder, if needed)
with the standard header row before proceeding.

This is the combined pipeline for the offboarding tracker at $Path:

1. Queries Freshservice for tickets matching $TicketStatus, $Category, and $SubCategory.
2. Pulls requested_items for each matching ticket and extracts every employee_name*
   custom field. Any Ticket+Name pair not already present in the CSV is appended as a new
   row, with all other columns left blank (the username typed on the ticket form is
   remembered for step 3). Existing rows and columns are never overwritten or cleared.
3. Reloads the full CSV and fills the 'Username' column, for any row still missing one, from
   the username on the ticket. It is used when AD confirms the account exists and the name
   matches ('First Last' or 'Last, First'); a mismatch shows the AD account and asks first.
   Anything still blank is entered manually and confirmed against AD.
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
   - Deletes the AD account ('AD User Deleted') once everything above is done, after checking
     in AD that it is still disabled and showing its display name, title, and OU to confirm.
   - Closes the ticket ('Ticket Closed'), with a resolution note, once every person listed on
     it has been deleted and every task on the ticket (e.g. IT Security's) is Completed.
   Anything declined or failed stays blank and is asked again on the next run.
5. Saves the CSV and prints a summary table of every row.

The CSV columns are kept in workflow order. Every column is
filled in by this script: the user's files are backed up first, then the AD account is deleted,
and the ticket is closed last.

If the Freshservice call fails, the script warns and proceeds to step 3 using whatever rows
already exist in the CSV, rather than aborting - but -DisableExpired is skipped for that run,
because ticket data can't be trusted. -DisableExpired is also skipped if Microsoft Graph
can't be reached.

Graph connection and license name translation come from LicenseHelpers.psm1 (same folder),
shared with Get-OffboardListLicenses.ps1.

.PARAMETER Path
The path to the CSV file. Defaults to C:\offboard\progress.csv.

.PARAMETER DefaultDomain
Domain appended to resolved usernames (e.g. "jla.org") to form a UPN for the Graph lookup.

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

.PARAMETER AssignGroupId
Freshservice group_id that open Off-board tickets are assigned to (numeric ID, not the name).
Defaults to the Tier 2 group.

.PARAMETER AssignAgentId
Freshservice responder_id (agent ID) that open Off-board tickets are assigned to. The agent must be
a member of the group above.

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
Then deletes the AD account and, once everyone listed on a ticket has been deleted, closes the ticket.
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
    [long]$AssignGroupId = 13000272117,

    [Parameter()]
    [long]$AssignAgentId = 13000472885,

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

$ticketUsernames = @{}   # 'ticket|name' -> username typed on the ticket form; checked against AD in Part 3

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

                # Username typed on the form: employee_username, or employee_username_2 etc. for extra people
                $suffix = if ($dateIndex -eq 1) { '' } else { "_$dateIndex" }
                $ticketUsername = $cf."employee_username$suffix"
                if (-not [string]::IsNullOrWhiteSpace($ticketUsername)) {
                    $ticketUsernames[$key] = "$ticketUsername".Trim()
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
# PART 3: Take usernames from the tickets and verify them in Active Directory
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

# Properties shown when confirming an account
$adProps = 'DisplayName', 'sAMAccountName', 'Title', 'Department', 'Enabled'

# Helper: one-line description of an AD account, so it can be checked against the ticket
function Format-AdUserChoice {
    param($AdUser)
    $details = @($AdUser.Title, $AdUser.Department) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $status  = if ($AdUser.Enabled) { 'enabled' } else { 'DISABLED' }
    $text    = "$($AdUser.DisplayName) [$($AdUser.sAMAccountName)]"
    if ($details) { $text += " - $($details -join ', ')" }
    return "$text - $status"
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

# The username comes from the ticket form. It is used only if AD agrees about who that is: a typo
# that happens to be a real account would otherwise point every later step (notes, disable,
# license and group removal, deletion) at the wrong person. Anything left blank goes to Part 4.
foreach ($user in $users) {
    $name = $user.Name.Trim()

    if (![string]::IsNullOrWhiteSpace($user.Username)) { continue }

    $ticketUsername = $ticketUsernames["$($user.Ticket)|$($user.Name)"]
    if ([string]::IsNullOrWhiteSpace($ticketUsername)) {
        Write-Warning "Ticket $($user.Ticket) has no username for '$name'."
        continue
    }

    $fromTicket = Get-ADUser -Filter { SamAccountName -eq $ticketUsername } -Properties $adProps
    if (-not $fromTicket) {
        Write-Warning "Ticket $($user.Ticket) lists username '$ticketUsername' for '$name', but no such account exists in AD."
    }
    elseif (Test-NameMatchesDisplayName -TicketName $name -DisplayName $fromTicket.DisplayName) {
        $user.Username = $fromTicket.sAMAccountName
    }
    else {
        Write-Host "`nTicket $($user.Ticket) lists username '$ticketUsername' for '$name', but AD has:" -ForegroundColor Yellow
        Write-Host "  $(Format-AdUserChoice $fromTicket)"
        $answer = Read-Host "Use this account for '$name'? (Y/N)"
        if ($answer -match '^(Y|YES)$') { $user.Username = $fromTicket.sAMAccountName }
    }
}

# ============================================================
# PART 4: Manual entry for any names that couldn't be resolved
# ============================================================
# A typed username is checked against AD, and the operator confirms the account it belongs to
# before it is used. A typo that happens to be a real account would otherwise point every later
# step (notes, disable, license and group removal) at the wrong person.
foreach ($user in $users) {
    if (-not [string]::IsNullOrWhiteSpace($user.Username)) { continue }

    Write-Host "`nNo confirmed username for '$($user.Name)'." -ForegroundColor Yellow

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

# Helper: write the resolution and close a Freshservice ticket (status 5). On this account's form the
# Resolution box is a custom field ('resolution' inside custom_fields), not the built-in
# resolution_notes. It is saved first, then the ticket is closed, so a failed close leaves the
# resolution in place (rewriting it on the next run is harmless). A blank description, which
# Freshservice requires to close, is filled in too. Throws on failure, using the
# API's own error message when it gives one, so the caller can warn and leave 'Ticket Closed' blank.
function Close-Ticket {
    param(
        [Parameter(Mandatory)][string]$Ticket,
        [Parameter(Mandatory)][string]$ResolutionNote,
        [Parameter(Mandatory)][string]$Description
    )
    $ticketUri    = "https://schc.freshservice.com/api/v2/tickets/$Ticket"
    $closePayload = @{ status = 5 } | ConvertTo-Json
    try {
        # Freshservice won't close a ticket whose description is blank, and tickets made from the form
        # have none. Fill it in only when it is blank, so a real description is never overwritten.
        $current = Invoke-RestMethod -Uri $ticketUri -Method Get -Headers $headers -ErrorAction Stop
        $fields  = @{ custom_fields = @{ resolution = $ResolutionNote } }
        if ([string]::IsNullOrWhiteSpace($current.ticket.description_text)) { $fields.description = $Description }
        $resolutionPayload = $fields | ConvertTo-Json -Depth 3

        $null = Invoke-RestMethod -Uri $ticketUri -Method Put -Headers $headers `
            -ContentType 'application/json; charset=utf-8' -Body $resolutionPayload -ErrorAction Stop
        $null = Invoke-RestMethod -Uri $ticketUri -Method Put -Headers $headers `
            -ContentType 'application/json; charset=utf-8' -Body $closePayload -ErrorAction Stop

        # A 200 reply doesn't prove the change was applied (Freshservice has been reported to accept
        # updates on service requests without applying them), so check how the ticket looks now.
        $check = Invoke-RestMethod -Uri $ticketUri -Method Get -Headers $headers -ErrorAction Stop
    }
    catch {
        $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        throw $detail
    }

    if ($check.ticket.status -ne 5) {
        throw "Freshservice accepted the update, but the ticket's status is still $($check.ticket.status) (not 5, Closed)."
    }
    if ([string]::IsNullOrWhiteSpace($check.ticket.custom_fields.resolution)) {
        Write-Warning "Ticket $Ticket is closed, but its Resolution field is empty."
    }
}

# Helper: returns the tasks on a ticket that are not yet Completed (task status 1 = Open,
# 2 = In Progress, 3 = Completed). Freshservice will not close a ticket while any are outstanding.
function Get-OpenTicketTasks {
    param([Parameter(Mandatory)][string]$Ticket)
    $response = Invoke-RestMethod -Uri "https://schc.freshservice.com/api/v2/tickets/$Ticket/tasks" `
        -Method Get -Headers $headers -ErrorAction Stop
    return @($response.tasks | Where-Object { -not $_.deleted -and $_.status -ne 3 })
}

# Helper: save the CSV now instead of waiting for Part 10. Used as a checkpoint so that notes already
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
# PART 4a: Drop rows whose ticket is Closed
# ============================================================
# Status 5 = Closed. A failed lookup keeps the row. Each ticket is checked once, however many rows it has.
# Under -WhatIf nothing is dropped; the rows are only reported.

if ($freshserviceOk) {
    $ticketIsClosed = @{}   # ticket number -> $true/$false

    $keepRows = foreach ($row in $users) {
        $ticketNumber = "$($row.Ticket)".Trim()
        if ($ticketNumber -notmatch '^\d+$') { $row; continue }

        if (-not $ticketIsClosed.ContainsKey($ticketNumber)) {
            try {
                $t = (Invoke-RestMethod -Uri "https://schc.freshservice.com/api/v2/tickets/$ticketNumber" -Headers $headers -ErrorAction Stop).ticket
                $ticketIsClosed[$ticketNumber] = ($t.status -eq 5)
            }
            catch {
                Write-Warning "Could not check ticket ${ticketNumber} (rows kept): $($_.Exception.Message)"
                $ticketIsClosed[$ticketNumber] = $false
            }
        }

        if (-not $ticketIsClosed[$ticketNumber]) { $row; continue }

        if ([string]::IsNullOrWhiteSpace($row.'AD User Deleted')) {
            Write-Warning "Ticket $ticketNumber is Closed but $($row.Name) has no AD deletion date."
        }
        if ($WhatIfPreference) {
            Write-Host "What if: row for $($row.Name) (ticket $ticketNumber, Closed) would be dropped."
            $row
        }
        else {
            Write-Host "Dropped $($row.Name) (ticket $ticketNumber is Closed)." -ForegroundColor Cyan
        }
    }

    $users = @($keepRows)
    Save-OffboardCsv
}

# ============================================================
# PART 4b: Low priority once every task is done; set group and agent
# ============================================================
# Priority 1 = Low. Only the open Off-board tickets pulled in Part 2 are touched, and only when
# something needs to change. Problems here warn and never stop the run.

if ($freshserviceOk -and $offboardTickets) {
    foreach ($obTicket in $offboardTickets) {
        try {
            $changes = [ordered]@{}
            if ($obTicket.group_id -ne $AssignGroupId)      { $changes['group_id']     = $AssignGroupId }
            if ($obTicket.responder_id -ne $AssignAgentId)  { $changes['responder_id'] = $AssignAgentId }

            # All tasks Completed (3) and at least one task exists. Same 'deleted' filter as Get-OpenTicketTasks.
            if ($obTicket.priority -ne 1) {
                $ticketTasks = @((Invoke-RestMethod -Uri "https://schc.freshservice.com/api/v2/tickets/$($obTicket.id)/tasks" -Headers $headers -ErrorAction Stop).tasks |
                    Where-Object { -not $_.deleted })
                if ($ticketTasks.Count -gt 0 -and @($ticketTasks | Where-Object { $_.status -ne 3 }).Count -eq 0) { $changes['priority'] = 1 }
            }

            if ($changes.Count -eq 0) { continue }
            if ($WhatIfPreference) {
                Write-Host "What if: ticket $($obTicket.id) would be updated: $($changes.Keys -join ', ')"
                continue
            }

            $ticketUri = "https://schc.freshservice.com/api/v2/tickets/$($obTicket.id)"
            $null = Invoke-RestMethod -Uri $ticketUri -Method Put -Headers $headers `
                -ContentType 'application/json; charset=utf-8' -Body ($changes | ConvertTo-Json) -ErrorAction Stop

            # Same check Close-Ticket does: a 200 reply doesn't prove the change was applied
            $after      = (Invoke-RestMethod -Uri $ticketUri -Method Get -Headers $headers -ErrorAction Stop).ticket
            $notApplied = @($changes.Keys | Where-Object { $after.$_ -ne $changes[$_] })
            if ($notApplied.Count -gt 0) { Write-Warning "Ticket $($obTicket.id): Freshservice accepted the update but did not apply: $($notApplied -join ', ')" }
            else { Write-Host "Ticket $($obTicket.id) updated: $($changes.Keys -join ', ')" -ForegroundColor Cyan }
        }
        catch {
            $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
            Write-Warning "Could not update ticket $($obTicket.id): $detail"
        }
    }
}

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

            $adminCred = Get-Credential -Message "Admin account for the file cleanup (enter it as $env:USERDOMAIN\username)"
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
                    Write-Warning "Detail: $($_.Exception.GetType().FullName), HResult $($_.Exception.HResult), host $($Host.Name), PowerShell $($PSVersionTable.PSVersion), as $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
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
# PART 8: Delete the AD account once everything else is done
# ============================================================

if ($runDisable) {
    try {
        # Ready for deletion: disabled, licenses and groups handled, files backed up, not yet deleted
        $deleteRows = @($users | Where-Object {
            $_.Username -match '\S' -and
            $_.'AD User Disabled' -eq 'Y' -and
            $_.'Licenses Removed' -match '\S' -and
            $_.'Groups Removed' -match '\S' -and
            $_.'Profile Folder Work' -match '\S' -and
            $_.'AD User Deleted' -notmatch '\S'
        })

        foreach ($user in $deleteRows) {
            $target = "$($user.Name) ($($user.Username))"

            # Look at the account as it is now, not as the CSV remembers it
            $adUser = $null
            $gone   = $false
            try { $adUser = Get-ADUser -Identity $user.Username -Properties Enabled, DisplayName, Title -ErrorAction Stop }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] { $gone = $true }
            catch {
                Write-Warning "Could not look up $target in AD, so it was not deleted: $($_.Exception.Message)"
                continue
            }

            if ($gone) {
                # Already deleted by some other means: record it so the ticket can still be closed
                if (-not $WhatIfPreference) {
                    Write-Warning "$target no longer exists in AD; recording it as deleted."
                    $user.'AD User Deleted' = Get-Date -Format 'yyyy-MM-dd'
                    Save-OffboardCsv
                }
                continue
            }

            if ($adUser.Enabled) {
                Write-Warning "$target is still enabled in AD, so it was not deleted."
                continue
            }

            # The prompt shows what AD has next to what the ticket said, so a wrong match is visible here
            $ou = ($adUser.DistinguishedName -split ',', 2)[1]
            if ($PSCmdlet.ShouldProcess("$target - AD: $($adUser.DisplayName), $($adUser.Title), $ou",
                    "DELETE the AD account (all other offboarding steps are done)")) {
                try {
                    Remove-ADUser -Identity $adUser -Confirm:$false -ErrorAction Stop
                    $user.'AD User Deleted' = Get-Date -Format 'yyyy-MM-dd'
                    Save-OffboardCsv   # checkpoint right away: the account is gone
                    try {
                        $noteBody = "AD account deleted for $([System.Net.WebUtility]::HtmlEncode($user.Name)) ($($user.Username)).<br><br>" +
                                    "Deleted on $(Get-Date -Format 'yyyy-MM-dd HH:mm') by $env:USERNAME, after the licenses, groups, and user files were handled.<br>"
                        Add-TicketNote -Ticket $user.Ticket -Body $noteBody
                    }
                    catch { Write-Warning "$target was deleted, but adding the ticket note failed: $($_.Exception.Message)" }
                }
                catch { Write-Warning "Could not delete ${target}: $($_.Exception.Message)" }
            }
        }
    }
    catch {
        Write-Warning "AD account deletion did not complete: $_"
    }
    finally {
        Save-OffboardCsv
    }
}

# ============================================================
# PART 9: Close the ticket once everyone on it has been deleted
# ============================================================

if ($runDisable) {
    try {
        $closeRows = @($users | Where-Object {
            $_.Ticket -match '\S' -and
            $_.'AD User Deleted' -match '\S' -and
            $_.'Ticket Closed' -notmatch '\S'
        })

        foreach ($ticketNumber in @($closeRows | Select-Object -ExpandProperty Ticket -Unique)) {
            $ticketRows = @($users | Where-Object { $_.Ticket -eq $ticketNumber })

            # Older tickets can list several people (and some may still be unresolved or waiting):
            # only close once every row on the ticket has been deleted
            if (@($ticketRows | Where-Object { $_.'AD User Deleted' -notmatch '\S' }).Count -gt 0) { continue }

            # The ticket also has task(s) for other teams (e.g. IT Security) that must be Completed first.
            # Not an error: the ticket is simply checked again on the next run.
            try { $openTasks = @(Get-OpenTicketTasks -Ticket $ticketNumber) }
            catch {
                Write-Warning "Could not check the tasks on ticket ${ticketNumber}, so it was left open: $($_.Exception.Message)"
                continue
            }
            if ($openTasks.Count -gt 0) {
                Write-Host "Ticket $ticketNumber is waiting on $($openTasks.Count) open task(s): $(($openTasks | ForEach-Object { $_.title }) -join '; '). It will be checked again on the next run." -ForegroundColor Yellow
                continue
            }

            if ($PSCmdlet.ShouldProcess("ticket $ticketNumber", "Close the Freshservice ticket (all offboarding steps are done)")) {
                try {
                    $who = ($ticketRows | ForEach-Object { "$($_.Name) ($($_.Username))" }) -join ', '
                    $usernames = ($ticketRows | ForEach-Object { $_.Username }) -join ', '
                    $resolution = "Offboarding complete for $who. All steps are done: AD account disabled, licenses and groups removed, " +
                                  "user files backed up and removed from the network shares, and AD account deleted. " +
                                  "Closed by $env:USERNAME on $(Get-Date -Format 'yyyy-MM-dd')."
                    Close-Ticket -Ticket $ticketNumber -ResolutionNote $resolution -Description "offboarding for $usernames"
                    $closedOn = Get-Date -Format 'yyyy-MM-dd'
                    foreach ($row in $ticketRows) { $row.'Ticket Closed' = $closedOn }
                    Save-OffboardCsv
                }
                catch { Write-Warning "Could not close ticket ${ticketNumber} (left open; 'Ticket Closed' stays blank): $_" }
            }
        }
    }
    catch {
        Write-Warning "Ticket closure did not complete: $_"
    }
    finally {
        Save-OffboardCsv
    }
}

# ============================================================
# PART 10: Save + summary output (all users, not just errors)
# ============================================================

$users | Export-Csv -Path $Path -NoTypeInformation

Write-Host "`n=== Offboard List Summary ===" -ForegroundColor Cyan
$users | Select-Object -Property $columnOrder | Format-Table -AutoSize -Wrap

Write-Host "`nDone! Tickets, usernames, and offboarding progress updated in $Path" -ForegroundColor Green