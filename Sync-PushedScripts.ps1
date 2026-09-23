<#
.SYNOPSIS
    Syncs scripts that have been pushed to GitHub over to the department share.

.DESCRIPTION
    Downloads a fresh temporary copy of the main branch of the PowerShell-Stuff
    GitHub repo, copies it to the department share with robocopy, then deletes
    the temporary copy.

    Because it only reads from GitHub, uncommitted or unpushed work is never
    copied, and nothing on H:\ is touched. Nothing already on the share is
    deleted (no /MIR). The .git and .github folders and .gitignore are skipped.

.EXAMPLE
    .\Sync-PushedScripts.ps1

.NOTES
    Author:   Will Hughes
    Created:  2026-09-23
    Requires: git installed on the machine running the script, and write
              access to the department share.
    Usage:    Run manually after pushing to GitHub.
#>

$RepoUrl = 'https://github.com/generilisk/PowerShell-Stuff.git'
$Dest    = '\\shastahealth.org\shared\ITS\Scripts'
$Branch  = 'main'
$Repo    = Join-Path $env:TEMP 'PowerShell-Stuff-sync'

# Clear any leftover temp copy from a previous run that crashed
if (Test-Path $Repo) { Remove-Item $Repo -Recurse -Force }

# Download the latest snapshot of main (no history)
git clone --depth 1 --branch $Branch $RepoUrl $Repo --quiet
if ($LASTEXITCODE -ne 0) { throw 'git clone failed' }

# Copy to the share; no /MIR, so nothing already on the share gets deleted
robocopy $Repo $Dest /E /XD .git .github /XF .gitignore /R:2 /W:5 /NP
$rc = $LASTEXITCODE    # save robocopy's result before cleanup overwrites it

# Delete the temp copy
Remove-Item $Repo -Recurse -Force

# Robocopy exit codes 0-7 are success; 8 and above are failures
if ($rc -ge 8) { throw "robocopy failed with exit code $rc" }