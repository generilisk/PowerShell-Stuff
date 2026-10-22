<#
.SYNOPSIS
    Employee offboarding – user data cleanup and archival.
.DESCRIPTION
    This script performs the data-handling portion of employee termination:
    - Copies the user's H Drive and plain profile folder (usersprofile$) to an
      HR archive location, then removes the originals
    - Deletes the V6 roaming profile folder (userprofiles$) directly, with no
      archival (roaming profile cache/churn, not something HR needs retained)
    - Logs all actions for audit and troubleshooting purposes
    DESIGN NOTES (Training):
    ------------------------
    Why V5.7 exists:
    - V5.6 proved that adding takeown fixed some stubborn profile roots.
    - But one large profile tree looked like it "looped" because we were doing
      very broad recursive ACL repair on the entire folder tree.
    - The log showed the script spending many minutes inside recursive icacls
      on the whole usersprofile tree rather than focusing on the actual failing
      child path.
    What changed in V5.7:
    - Change strategy from broad whole-tree repair to targeted subtree repair.
    - We now try to extract the actual failing child path from the error text
      and repair only that subtree when possible.
    - We keep a max-attempt loop so the script cannot spin forever.
    - We still keep special handling for literal $RECYCLE.BIN folders.
    - We keep honest final status reporting.

    Why V5.8 exists:
    - The H Drive path was previously assumed from a fixed naming convention
      (\\ntdfs01\HDrive\<username>), which doesn't hold for every account and
      isn't the authoritative source for where a user's home folder lives.

    What changed in V5.8:
    - H Drive path is now resolved from the user's AD HomeDirectory attribute
      (Get-ADUser -Properties HomeDirectory) instead of being constructed from
      a hardcoded root and username.
    - Falls back to the old constructed path (\\ntdfs01\HDrive\<username>) only
      if the AD attribute is empty or the lookup fails, so accounts without a
      HomeDirectory value set still get processed.
    - Removed the pre-flight Validate-ShareRoot check against the old fixed
      \\ntdfs01\HDrive root, since the resolved path is no longer guaranteed
      to live under that share. Existence is still checked per-user inside
      the archive function.
    - Logs which source (AD vs. fallback) was used for the H Drive path on
      each run, for audit purposes.
    - Requires the ActiveDirectory module (RSAT-AD-PowerShell) to be available
      wherever this script runs.

    Why V5.9 exists:
    - The plain profile folder (usersprofile$) previously went straight to
      direct deletion, same as the V6 roaming profile folder. That's no
      longer desired - it needs to be archived to HR first, matching the
      H Drive workflow.
    - The V6 roaming profile folder (userprofiles$) is intentionally left on
      direct deletion - it's roaming profile churn/cache, not something HR
      needs retained.

    What changed in V5.9:
    - Generalized the old Copy-And-RemoveHDrive function into
      Copy-And-RemoveFolder(SourcePath, DestinationFolder, Label), reusable
      for any folder that needs archive-then-delete handling.
    - Copy-And-RemoveFolder does the archive copy first (with ACL repair and
      one retry on failure, same pattern H Drive already used). If the copy
      cannot be completed, the source is left in place and NOT deleted - we
      never delete something we failed to archive.
    - Once archived, deletion of the source is handed off to the existing
      Remove-FolderWithFallback function, so archived folders still get the
      full targeted-repair / multi-attempt delete logic from V5.7, rather
      than falling back to a simpler single-retry pattern.
    - ProfileUser (usersprofile$) now goes through Copy-And-RemoveFolder.
    - ProfileV6 (userprofiles$) is unchanged from V5.8 - still a direct
      Remove-FolderWithFallback call, no archival.
    - No changes to the access-control methods themselves: Repair-PathAccess
      still uses only takeown/icacls via cmd.exe, matching the approved
      methods from the original script.
.VERSION
    5.9
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$UserName,

    # Skips both Read-Host prompts so the script can run in a hidden child process.
    [switch]$Unattended
)

Import-Module ActiveDirectory

# ==============================
# INPUT
# ==============================
$username = $UserName.Trim()

if ([string]::IsNullOrWhiteSpace($username)) {
    Write-Host "[ERROR] Username cannot be blank." -ForegroundColor Red
    return
}

# ==============================
# PATH DEFINITIONS
# ==============================
$hrRoot          = "\\shared01\hr\_HR Shared\TERMINATION\Termed employee files"
$hrUserFolder    = Join-Path $hrRoot $username
$hrHDriveFolder  = Join-Path $hrUserFolder "H Drive"
$hrProfileUserFolder = Join-Path $hrUserFolder "usersprofile"

$userProfile1Root = "\\shared01\userprofiles$"
$userProfile2Root = "\\shared01\usersprofile$"
$userProfile1     = Join-Path $userProfile1Root "${username}.V6"
$userProfile2     = Join-Path $userProfile2Root $username

$adUser = Get-ADUser `
    -Identity $username `
    -Properties DisplayName,HomeDirectory `
    -ErrorAction SilentlyContinue
if (-not $adUser) {
    Write-Host "[ERROR] User '$username' was not found in Active Directory." -ForegroundColor Red
    return
}
Write-Host ""
Write-Host "You are about to process:" -ForegroundColor Yellow
Write-Host "$($adUser.SamAccountName)  ->  $($adUser.DisplayName)" -ForegroundColor Cyan
Write-Host ""

if (-not $Unattended) {
    $confirm = Read-Host "Continue processing this user? (Y/N)"

    if ($confirm -notmatch '^(Y|YES)$') {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }
}
$hDriveSource = "AD"
$hDriveUserPath = if ($adUser -and $adUser.HomeDirectory) {
    $adUser.HomeDirectory
} else {
    $hDriveSource = "Fallback (constructed path)"
    Join-Path "\\ntdfs01\HDrive" $username
}

$logPath         = Join-Path $hrUserFolder "$((Get-Date).ToString('yyyy-MM-dd_HHmmss'))-Termination.log"

$script:OutcomeSummary = [ordered]@{}
$script:HadErrors = $false

# ==============================
# FUNCTIONS
# ==============================
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "[$timestamp] [$Level] $Message"
}

function Set-Outcome {
    param(
        [string]$Key,
        [string]$Status,
        [string]$Detail = ""
    )

    $script:OutcomeSummary[$Key] = [pscustomobject]@{
        Status = $Status
        Detail = $Detail
    }
}

function Validate-ShareRoot {
    param ([string]$Path)

    if (-not (Test-Path $Path)) {
        Write-Host "[ERROR] Cannot access share root: $Path" -ForegroundColor Red
        Write-Log "Share root inaccessible: $Path" "FATAL"
        return $false
    }

    Write-Log "Validated share root: $Path"
    return $true
}

function Invoke-CmdChecked {
    param(
        [string]$Command,
        [string]$FailureMessage,
        [switch]$IgnoreExitCode
    )

    Write-Log "Running native command: $Command"
    cmd.exe /c "$Command >nul 2>&1"
    $exitCode = $LASTEXITCODE

    if (-not $IgnoreExitCode -and $exitCode -ne 0) {
        throw "$FailureMessage (exit code $exitCode)"
    }

    return $exitCode
}

function Get-FailingPathFromError {
    param(
        [string]$RootPath,
        [string]$ErrorText
    )

    if ([string]::IsNullOrWhiteSpace($ErrorText)) {
        return $null
    }

    # First preference: quoted path from the exception.
    $quoted = [regex]::Matches($ErrorText, "'([^']+)'")
    foreach ($m in $quoted) {
        $candidate = $m.Groups[1].Value
        if ($candidate -like "$RootPath*") {
            return $candidate
        }
    }

    # Fallback: any UNC path mention that begins with the root path.
    $escapedRoot = [regex]::Escape($RootPath)
    $m2 = [regex]::Match($ErrorText, "$escapedRoot[^\r\n|]*")
    if ($m2.Success) {
        return $m2.Value.Trim()
    }

    return $null
}

function Repair-PathAccess {
    param(
        [string]$Path,
        [switch]$Recursive
    )

    Write-Log "ACL repair triggered for: $Path" "WARN"

    $takeownCmd = "takeown /F `"$Path`" /D Y"
    $setOwnerCmd = "icacls `"$Path`" /setowner Administrators /Q"
    $grantCmd = "icacls `"$Path`" /grant Administrators:(OI)(CI)F /Q"
    $inheritCmd = "icacls `"$Path`" /inheritance:e /Q"

    if ($Recursive) {
        $takeownCmd += " /R"
        $setOwnerCmd += " /T /C"
        $grantCmd += " /T /C"
        $inheritCmd += " /T /C"
    }

    try {
        Invoke-CmdChecked -Command $takeownCmd -FailureMessage "Failed to take ownership of $Path"
        Invoke-CmdChecked -Command $setOwnerCmd -FailureMessage "Failed to set owner on $Path"
        Invoke-CmdChecked -Command $grantCmd -FailureMessage "Failed to grant Administrators full control on $Path"
        Invoke-CmdChecked -Command $inheritCmd -FailureMessage "Failed to enable inheritance on $Path"
        Write-Log "ACL repair completed for: $Path" "SUCCESS"
    }
    catch {
        Write-Log "ACL repair failed for $Path | Message: $($_.Exception.Message) | FullError: $_" "ERROR"
        throw
    }
}

function Remove-LiteralRecycleBinsUnderPath {
    param([string]$ParentPath)

    $foundAny = $false

    try {
        $recycleCandidates = Get-ChildItem -Path $ParentPath -Directory -Force -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq '$RECYCLE.BIN' }
    }
    catch {
        Write-Log ('Failed while searching for literal $RECYCLE.BIN paths under: {0} | Message: {1} | FullError: {2}' -f $ParentPath, $_.Exception.Message, $_) 'WARN'
        return $false
    }

    foreach ($candidate in $recycleCandidates) {
        $foundAny = $true
        $rbPath = $candidate.FullName
        Write-Log ('Attempting targeted removal of stubborn literal $RECYCLE.BIN path: {0}' -f $rbPath) 'WARN'
        try {
            Repair-PathAccess -Path $rbPath -Recursive
            $escaped = '"' + $rbPath + '"'
            Invoke-CmdChecked -Command "attrib -h -s $escaped /S /D" -FailureMessage "Failed to clear attributes on $rbPath" -IgnoreExitCode | Out-Null
            Invoke-CmdChecked -Command "del /f /q /s $escaped\*" -FailureMessage "Failed to delete files under $rbPath" -IgnoreExitCode | Out-Null
            Invoke-CmdChecked -Command "rd /s /q $escaped" -FailureMessage "Failed to remove $rbPath with rd"
            Write-Log ('Removed stubborn literal $RECYCLE.BIN path: {0}' -f $rbPath) 'SUCCESS'
        }
        catch {
            Write-Log ('Targeted literal $RECYCLE.BIN removal failed for {0} | Message: {1} | FullError: {2}' -f $rbPath, $_.Exception.Message, $_) 'WARN'
        }
    }

    if (-not $foundAny) {
        Write-Log ('No literal $RECYCLE.BIN folders found anywhere under: {0}' -f $ParentPath)
    }

    return $foundAny
}

function Purge-TargetPathNative {
    param([string]$TargetPath)

    if (-not (Test-Path $TargetPath)) {
        Write-Log "Native purge skipped because path no longer exists: $TargetPath"
        return $true
    }

    Write-Log "Starting targeted native purge for: $TargetPath" "WARN"

    try {
        Repair-PathAccess -Path $TargetPath -Recursive
        $escaped = '"' + $TargetPath + '"'

        if ((Get-Item -LiteralPath $TargetPath -Force).PSIsContainer) {
            Invoke-CmdChecked -Command "attrib -h -s $escaped\* /S /D" -FailureMessage "Failed to clear attributes under $TargetPath" -IgnoreExitCode | Out-Null
            Invoke-CmdChecked -Command "del /f /q /s $escaped\*" -FailureMessage "Failed to delete one or more files under $TargetPath" -IgnoreExitCode | Out-Null
            Invoke-CmdChecked -Command "rd /s /q $escaped" -FailureMessage "Failed to remove directory $TargetPath with rd" -IgnoreExitCode | Out-Null
        }
        else {
            Invoke-CmdChecked -Command "attrib -h -s $escaped" -FailureMessage "Failed to clear attributes on $TargetPath" -IgnoreExitCode | Out-Null
            Invoke-CmdChecked -Command "del /f /q $escaped" -FailureMessage "Failed to delete file $TargetPath" -IgnoreExitCode | Out-Null
        }

        Write-Log "Targeted native purge completed for: $TargetPath" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Targeted native purge failed for $TargetPath | Message: $($_.Exception.Message) | FullError: $_" "WARN"
        return $false
    }
}

function Invoke-RobocopyChecked {
    param(
        [string]$Source,
        [string]$Destination
    )

    $args = @(
        $Source,
        $Destination,
        '/E',
        '/COPY:DAT',
        '/DCOPY:DAT',
        '/R:1',
        '/W:1',
        '/XJ',
        '/NFL',
        '/NDL',
        '/NJH',
        '/NJS'
    )

    Write-Log "Starting robocopy from $Source to $Destination"
    & robocopy @args | Out-Null
    $rc = $LASTEXITCODE
    Write-Log "Robocopy exit code for $Source : $rc"

    if ($rc -gt 7) {
        throw "Robocopy failed with exit code $rc"
    }
}

function Remove-FolderWithFallback {
    param(
        [string]$Path,
        [string]$Label
    )

    Write-Log "Checking folder path: $Path"

    if (-not (Test-Path $Path)) {
        Write-Log "Folder not found (skipped): $Path"
        Set-Outcome -Key $Label -Status 'SKIPPED' -Detail 'Folder not found'
        return $true
    }

    try {
        Write-Log "Attempting delete: $Path"
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        Write-Log "Deleted folder: $Path" "SUCCESS"
        Set-Outcome -Key $Label -Status 'SUCCESS' -Detail 'Deleted with normal method'
        return $true
    }
    catch {
        $lastErrorText = $_.ToString()
        Write-Log "Initial delete failed for folder: $Path | Message: $($_.Exception.Message) | FullError: $_" "WARN"
    }

    # Training note:
    # We intentionally cap attempts so the script cannot look like it is looping forever.
    $maxAttempts = 4
    $attempt = 0
    $lastTargetRepaired = $null

    while ($attempt -lt $maxAttempts) {
        $attempt++
        $failingPath = Get-FailingPathFromError -RootPath $Path -ErrorText $lastErrorText
        if ([string]::IsNullOrWhiteSpace($failingPath)) {
            $failingPath = $Path
        }

        # If the same failing path keeps coming back, escalate only once to the root.
        if ($attempt -gt 1 -and $failingPath -eq $lastTargetRepaired -and $failingPath -ne $Path) {
            Write-Log "Same failing child path repeated. Escalating repair scope to root path: $Path" "WARN"
            $failingPath = $Path
        }

        $useRecursiveRepair = $true
        Write-Log "Recovery attempt $attempt of $maxAttempts will target: $failingPath" "WARN"

        try {
            Repair-PathAccess -Path $failingPath -Recursive:$useRecursiveRepair
        }
        catch {
            Write-Log "Repair attempt failed for target path: $failingPath | Message: $($_.Exception.Message) | FullError: $_" "WARN"
        }

        $lastTargetRepaired = $failingPath

        if ($failingPath -like '*$RECYCLE.BIN*' -or $Path -like '*usersprofile$*') {
            [void](Remove-LiteralRecycleBinsUnderPath -ParentPath $Path)
        }

        if ($attempt -ge 3) {
            [void](Purge-TargetPathNative -TargetPath $failingPath)
        }

        try {
            Write-Log "Retrying delete after recovery attempt $attempt $Path"
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            Write-Log "Deleted folder after recovery attempt $attempt $Path" "SUCCESS"
            Set-Outcome -Key $Label -Status 'SUCCESS' -Detail "Deleted after targeted recovery attempt $attempt"
            return $true
        }
        catch {
            $lastErrorText = $_.ToString()
            Write-Log "Delete failed after recovery attempt $attempt for folder: $Path | Message: $($_.Exception.Message) | FullError: $_" "WARN"
        }
    }

    Set-Outcome -Key $Label -Status 'FAILED' -Detail "Folder still exists or access denied: $Path"
    return $false
}

function Copy-And-RemoveFolder {
    param(
        [string]$SourcePath,
        [string]$DestinationFolder,
        [string]$Label
    )

    Write-Log "Checking $Label path: $SourcePath"

    if (-not (Test-Path $SourcePath)) {
        Write-Log "$Label not found. Skipping archive."
        Set-Outcome -Key $Label -Status 'SKIPPED' -Detail 'Folder not found'
        return $true
    }

    try {
        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
        Write-Log "Created/validated HR archive folder for ${Label}: $DestinationFolder"
    }
    catch {
        Write-Log "Failed to create HR archive folder for ${Label}: $DestinationFolder | Message: $($_.Exception.Message) | FullError: $_" "ERROR"
        Set-Outcome -Key $Label -Status 'FAILED' -Detail 'Could not create HR archive folder'
        return $false
    }

    try {
        Invoke-RobocopyChecked -Source $SourcePath -Destination $DestinationFolder
        Write-Log "Copied $Label to HR archive." "SUCCESS"
    }
    catch {
        Write-Log "Initial $Label copy failed for $SourcePath | Message: $($_.Exception.Message) | FullError: $_" "WARN"

        try {
            Repair-PathAccess -Path $SourcePath -Recursive
            Invoke-RobocopyChecked -Source $SourcePath -Destination $DestinationFolder
            Write-Log "Copied $Label to HR archive after ACL repair." "SUCCESS"
        }
        catch {
            Write-Log "$Label copy failed after ACL repair for $SourcePath | Message: $($_.Exception.Message) | FullError: $_" "ERROR"
            Set-Outcome -Key $Label -Status 'FAILED' -Detail 'Archive copy failed; source left in place'
            return $false
        }
    }

    # Archive succeeded - now remove the source using the same targeted-repair
    # multi-attempt delete logic that direct profile deletion already relied on.
    $deleted = Remove-FolderWithFallback -Path $SourcePath -Label $Label
    $deleteOutcome = $script:OutcomeSummary[$Label]

    if ($deleted) {
        Set-Outcome -Key $Label -Status 'SUCCESS' -Detail "Archived to $DestinationFolder; source removed ($($deleteOutcome.Detail))"
        return $true
    }
    else {
        Set-Outcome -Key $Label -Status 'PARTIAL' -Detail "Archived to $DestinationFolder, but source could not be removed ($($deleteOutcome.Detail))"
        return $false
    }
}

# ==============================
# START PROCESS
# ==============================
New-Item -Path $hrUserFolder -ItemType Directory -Force | Out-Null
New-Item -Path $logPath -ItemType File -Force | Out-Null

Write-Log "Termination script started for user: $username"
Write-Log "Resolved paths in use for this run:"
Write-Log "  HR Root: $hrRoot"
Write-Log "  HR User Folder: $hrUserFolder"
Write-Log "  Profile Path 1 (V6, direct delete): $userProfile1"
Write-Log "  Profile Path 2 (archived): $userProfile2"
Write-Log "  Profile User Archive: $hrProfileUserFolder"
Write-Log "  H Drive Path: $hDriveUserPath"
Write-Log "  H Drive Source: $hDriveSource"
Write-Log "  Operator: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"

if (-not (Validate-ShareRoot $hrRoot)) { return }
if (-not (Validate-ShareRoot $userProfile1Root)) { return }
if (-not (Validate-ShareRoot $userProfile2Root)) { return }

Write-Log "Starting profile cleanup"
[void](Remove-FolderWithFallback -Path $userProfile1 -Label 'ProfileV6')
[void](Copy-And-RemoveFolder -SourcePath $userProfile2 -DestinationFolder $hrProfileUserFolder -Label 'ProfileUser')

Write-Log "Processing H Drive"
[void](Copy-And-RemoveFolder -SourcePath $hDriveUserPath -DestinationFolder $hrHDriveFolder -Label 'HDrive')

Write-Log "Outcome summary:"
foreach ($key in $script:OutcomeSummary.Keys) {
    $item = $script:OutcomeSummary[$key]
    Write-Log "  ${key}: $($item.Status) - $($item.Detail)"
}

Write-Log "Termination script completed"

$failedOutcomes = @($script:OutcomeSummary.GetEnumerator() | Where-Object { $_.Value.Status -in @('FAILED','PARTIAL') })
if ($failedOutcomes.Count -gt 0) {
    $script:HadErrors = $true
}

$finalStatus = if ($script:OutcomeSummary.Count -gt 0 -and $failedOutcomes.Count -eq 0) { 'Success' } else { 'Failed' }
Write-Log "RESULT|$username|$finalStatus"
Set-Content -LiteralPath (Join-Path $hrUserFolder 'Termination.result') -Value "$username,$finalStatus,$(Get-Date -Format s)"

if ($failedOutcomes.Count -eq 0) {
    Write-Host "`n[SUCCESS] Offboarding completed for $username" -ForegroundColor Green
}
elseif ($script:OutcomeSummary.Count -gt 0) {
    Write-Host "`n[PARTIAL] Offboarding finished with items requiring review for $username" -ForegroundColor Yellow
    foreach ($entry in $failedOutcomes) {
        Write-Host " - $($entry.Key): $($entry.Value.Status) | $($entry.Value.Detail)" -ForegroundColor Yellow
    }
    Write-Host "Check the log for details: $logPath" -ForegroundColor Yellow
}
else {
    Write-Host "`n[ERROR] Offboarding could not complete for $username" -ForegroundColor Red
    Write-Host "Check the log for details: $logPath" -ForegroundColor Red
}

if ($script:HadErrors -and -not $Unattended) {
    Read-Host "`nErrors were encountered. Press Enter to exit"
}