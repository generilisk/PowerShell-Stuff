<#
.SYNOPSIS
    Syncs scripts that have been pushed to GitHub over to the department share.

.DESCRIPTION
    Downloads a fresh temporary copy of the main branch of the PowerShell-Stuff
    GitHub repo and syncs it to the department share: new and changed files are
    copied (compared by content, not timestamp), and files that were removed from
    the repo are deleted from the share.

    Because it only reads from GitHub, uncommitted or unpushed work is never
    copied, and nothing on H:\ is touched. Only files listed in the manifest
    (.sync-manifest.txt on the share, written by this script) are ever deleted,
    so anything else on the share is left alone. The .git, .github and .vscode
    folders and .gitignore are skipped. Supports -WhatIf.

.EXAMPLE
    .\Sync-PushedScripts.ps1

.NOTES
    Author:   Will Hughes
    Created:  2026-09-23
    Requires: git installed on the machine running the script, and write
              access to the department share.
    Usage:    Run manually after pushing to GitHub.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param ()

$RepoUrl = 'https://github.com/generilisk/PowerShell-Stuff.git'
$Dest    = '\\shastahealth.org\shared\ITS\Scripts'
$Branch  = 'main'
$Repo    = Join-Path $env:TEMP 'PowerShell-Stuff-sync'

# Clear any leftover temp copy from a previous run that crashed
if (Test-Path $Repo) { Remove-Item $Repo -Recurse -Force -WhatIf:$false }

# Download the latest snapshot of main (no history)
git clone --depth 1 --branch $Branch $RepoUrl $Repo --quiet
if ($LASTEXITCODE -ne 0) { throw 'git clone failed' }

# Files and folders that stay off the share
$ExcludeDirs  = '.git', '.github', '.vscode'
$ExcludeFiles = '.gitignore'

# Record of the files this script has put on the share. Only files listed here are ever
# deleted, so anything else that lives on the share is left alone.
$ManifestPath = Join-Path $Dest '.sync-manifest.txt'

# Every file in the fresh copy of main that belongs on the share
$repoRoot = (Resolve-Path -LiteralPath $Repo).Path.TrimEnd('\')
$sourceFiles = @(Get-ChildItem -LiteralPath $repoRoot -Recurse -File -Force | Where-Object {
    $parts = $_.FullName.Substring($repoRoot.Length + 1).Split('\')
    $dirs  = @($parts | Select-Object -SkipLast 1)
    -not ($dirs | Where-Object { $_ -in $ExcludeDirs }) -and ($_.Name -notin $ExcludeFiles)
})
if ($sourceFiles.Count -eq 0) { throw 'No files found in the downloaded copy; nothing will be changed.' }

# Copy new and changed files. Content is compared by hash because every file in a fresh
# clone has a new timestamp, so timestamps can't show what actually changed.
$syncedFiles = [System.Collections.Generic.List[string]]::new()
$copied = 0; $unchanged = 0; $removed = 0; $failed = 0

foreach ($file in $sourceFiles) {
    $rel    = $file.FullName.Substring($repoRoot.Length + 1)
    $target = Join-Path $Dest $rel
    $syncedFiles.Add($rel)

    $existing = Get-Item -LiteralPath $target -ErrorAction SilentlyContinue
    if ($existing -and $existing.Length -eq $file.Length -and
        (Get-FileHash -LiteralPath $file.FullName).Hash -eq (Get-FileHash -LiteralPath $target).Hash) {
        $unchanged++
        continue
    }

    if ($PSCmdlet.ShouldProcess($target, 'Copy')) {
        try {
            $null = New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force
            Copy-Item -LiteralPath $file.FullName -Destination $target -Force -ErrorAction Stop
            Write-Host "Copied  $rel"
            $copied++
        }
        catch {
            Write-Warning "Could not copy ${rel}: $($_.Exception.Message)"
            $failed++
        }
    }
}

# Delete files this script put on the share earlier that are no longer in the repo
$keepInManifest = [System.Collections.Generic.List[string]]::new($syncedFiles)
if (Test-Path -LiteralPath $ManifestPath) {
    $current = [System.Collections.Generic.HashSet[string]]::new([string[]]$syncedFiles, [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($rel in (Get-Content -LiteralPath $ManifestPath | Where-Object { $_.Trim() })) {
        if ($current.Contains($rel)) { continue }
        if ($rel -match '(^|\\)\.\.(\\|$)' -or [System.IO.Path]::IsPathRooted($rel)) {
            Write-Warning "Ignoring suspicious manifest entry: $rel"
            continue
        }
        $target = Join-Path $Dest $rel
        if (-not (Test-Path -LiteralPath $target)) { continue }

        if ($PSCmdlet.ShouldProcess($target, 'Delete (no longer in the repo)')) {
            try {
                Remove-Item -LiteralPath $target -Force -ErrorAction Stop
                Write-Host "Deleted $rel"
                $removed++
                # Remove folders left empty, stopping at the share root
                $dir = Split-Path $target -Parent
                while ($dir.Length -gt $Dest.Length -and -not (Get-ChildItem -LiteralPath $dir -Force)) {
                    Remove-Item -LiteralPath $dir -Force
                    $dir = Split-Path $dir -Parent
                }
            }
            catch {
                Write-Warning "Could not delete ${rel}: $($_.Exception.Message)"
                $keepInManifest.Add($rel)   # still on the share, so keep it on record
                $failed++
            }
        }
        else { $keepInManifest.Add($rel) }
    }
}

if ($PSCmdlet.ShouldProcess($ManifestPath, 'Update sync manifest')) {
    Set-Content -LiteralPath $ManifestPath -Value $keepInManifest -Encoding UTF8
}

# Delete the temp copy
Remove-Item $Repo -Recurse -Force -WhatIf:$false

Write-Host "Sync complete: $copied copied, $unchanged unchanged, $removed deleted, $failed failed."
if ($failed -gt 0) { throw "$failed file operation(s) failed; see the warnings above." }