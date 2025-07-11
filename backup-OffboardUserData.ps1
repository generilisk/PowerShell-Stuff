<#
.SYNOPSIS
    Employee data offboarding script for terminated users.

.DESCRIPTION
    This script handles the data management portion of employee termination by:
    - Backing up user's H Drive to HR folder
    - Removing user profile folders from network shares
    - Logging all file operations for audit purposes
    - Verifying network path accessibility before processing

.EXAMPLE
    .\backup-UserData.ps1
    
    Prompts for a username, backs up the user's H Drive to the HR termination folder,
    removes user profile folders, and logs all operations.

.AUTHOR
    Will Hughes

.VERSION
    1.0.0
#>

function Test-NetworkPath {
    param (
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$LogPath
    )
    
    try {
        $basePath = Split-Path $Path -Parent
        if (-not (Test-Path $basePath)) {
            Add-Content $LogPath "[ERROR] Network path not accessible: $basePath"
            return $false
        }
        Add-Content $LogPath "[SUCCESS] Network path accessible: $basePath"
        return $true
    } catch {
        Add-Content $LogPath "[ERROR] Failed to test network path $Path`: $_"
        return $false
    }
}

function Set-FolderPermissions {
    param (
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Account,
        [Parameter(Mandatory)][string]$LogPath
    )

    if (-not (Test-Path $Path)) {
        Add-Content $LogPath "`n[WARNING] Path does not exist: $Path"
        return
    }

    Add-Content $LogPath "`nSetting permissions for $Account on $Path"

    try {
        # Step 1: Take ownership (makes current user the owner)
        Write-Progress -Activity "Setting Permissions" -Status "Taking ownership of: $Path" -PercentComplete 25
        Write-Host "[INFO] Taking ownership of: $Path" -ForegroundColor Yellow
        Add-Content $LogPath "Taking ownership of: $Path"
        $takeownResult = & takeown /f $Path /r /d y 2>&1
        Add-Content $LogPath "takeown result: $($takeownResult -join "`n")`n"
        
        # Check if takeown was successful
        if ($LASTEXITCODE -ne 0) {
            Add-Content $LogPath "[WARNING] takeown may have failed with exit code: $LASTEXITCODE"
        }
        
        # Step 2: Grant full control to the account (should be same as current user)
        Write-Progress -Activity "Setting Permissions" -Status "Granting full control to: $Account" -PercentComplete 75
        Write-Host "[INFO] Granting full control to: $Account" -ForegroundColor Yellow
        Add-Content $LogPath "Granting full control to: $Account"
        $result = & icacls $Path /grant "${Account}:F" /T /C 2>&1
        Add-Content $LogPath "icacls result: $($result -join "`n")`n"
        
        # Note: No need to setowner since takeown already made us the owner
        # and we're granting permissions to the same account

        Write-Progress -Activity "Setting Permissions" -Status "Completed: $Path" -PercentComplete 100
        Write-Host "[SUCCESS] Permissions updated on: $Path" -ForegroundColor Green
        Add-Content $LogPath "[SUCCESS] Permissions updated on: $Path"
        Write-Progress -Activity "Setting Permissions" -Completed
    } catch {
        Write-Host "[ERROR] Failed to set permissions on ${Path}: $_" -ForegroundColor Red
        Add-Content $LogPath "[ERROR] Failed to set permissions on ${Path}: $_"
        Write-Progress -Activity "Setting Permissions" -Completed
    }
}

function Remove-FolderIfExists {
	param (
		[Parameter(Mandatory)][string]$Path,
		[Parameter(Mandatory)][string]$LogPath
	)

	if (Test-Path $Path) {
		try {
			Remove-Item -Path $Path -Recurse -Force
			Add-Content $LogPath "[SUCCESS] Deleted: $Path"
		} catch {
			Add-Content $LogPath "[ERROR] Failed to delete ${Path}: $_"
		}
	} else {
		Add-Content $LogPath "[INFO] Skipped delete (not found): $Path"
	}
}

function Copy-HDrive {
    param (
        [Parameter(Mandatory)]
        [string]$Source,
        [Parameter(Mandatory)]
        [string]$Destination,
        [Parameter(Mandatory)]
        [string]$LogPath
    )

    if (-not (Test-Path $Source)) {
        Add-Content -Path $LogPath -Value "[WARNING] Source path does not exist: $Source"
        return $false
    }

    try {
        New-Item -Path $Destination -ItemType Directory -Force | Out-Null
        Copy-Item -Path "$Source\*" -Destination $Destination -Recurse -Force -ErrorAction Stop
        Add-Content -Path $LogPath -Value "Copied H Drive from ${Source} to ${Destination}"
        return $true
    } catch {
        Add-Content -Path $LogPath -Value "[ERROR] Failed to copy H Drive: $_"
        return $false
    }
}

# === INPUTS ===
$username = Read-Host "Enter the username (e.g., jsmith)"
# Get the current security context (works with "Run as")
$currentUser = whoami
$accountName = $currentUser.Split('\\')[-1]  # Extract just the username part
Write-Host "Using current security context for permissions: $currentUser" -ForegroundColor Green
Write-Host "Account name: $accountName" -ForegroundColor Green

# === SETUP BASIC PATHS FOR LOGGING ===
$hrFolderRoot = "\\shared01\hr\_HR Shared\TERMINATION\Termed employee files"
$hrUserFolder = Join-Path $hrFolderRoot $username
$logPath = Join-Path $hrUserFolder "Termed Script.log"

# === VERIFY USERNAME ===
try {
    $user = Get-ADUser -Identity $username -Properties DisplayName
    $fullName = $user.DisplayName
    Write-Host "`n=== USER VERIFICATION ===" -ForegroundColor Cyan
    Write-Host "Username entered: $username" -ForegroundColor Yellow
    Write-Host "Full Name: $fullName" -ForegroundColor Yellow
    
    # Confirmation prompt
    $confirmation = Read-Host "`nIs this the correct user? (y/n)"
    if ($confirmation -notmatch '^[Yy]') {
        Write-Host "Operation cancelled by user." -ForegroundColor Red
        Read-Host "`nPress Enter to exit"
        return
    }
    
    Write-Host "[SUCCESS] User verification confirmed" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Could not verify user: $username" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    
    # Ask if they want to continue anyway
    $continueAnyway = Read-Host "`nUnable to verify username. Continue anyway? (y/n)"
    if ($continueAnyway -notmatch '^[Yy]') {
        Write-Host "Operation cancelled." -ForegroundColor Red
        Read-Host "`nPress Enter to exit"
        return
    }
    
    Write-Host "[WARNING] Proceeding without user verification" -ForegroundColor Yellow
}

# === PATHS ===
$userProfilePath  = "\\shared01\f$\Shares\UserProfiles$\${username}.V6"
$userProfile2Path = "\\shared01\e$\UsersProfile\${username}"
$hDrivePath       = "\\ntdfs01\e$\shares\HDrive\${username}"
$hrHDriveFolder   = Join-Path $hrUserFolder "H Drive"

# === CREATE HR USER FOLDER ===
if (-not (Test-Path $hrUserFolder)) {
	New-Item -Path $hrUserFolder -ItemType Directory -Force | Out-Null
	$folderMessage = "HR folder created at: $hrUserFolder"
} else {
	$folderMessage = "HR folder already exists at: $hrUserFolder"
}

# Create or overwrite log file
New-Item -Path $logPath -ItemType File -Force | Out-Null

# Write initial log entries
Add-Content $logPath ("=" * 50)
Add-Content $logPath "Termination Script Started: $(Get-Date)"
Add-Content $logPath ("=" * 50)
Add-Content $logPath $folderMessage

# === VALIDATE NETWORK PATHS ===
Add-Content $logPath "`nValidating network path accessibility..."

if (-not (Test-NetworkPath -Path $userProfilePath -LogPath $logPath)) {
    Write-Host "[ERROR] Cannot access network path for user profile: $userProfilePath" -ForegroundColor Red
    Add-Content $logPath "[FATAL] Script terminating due to network path access failure"
    Read-Host "`nPress Enter to exit"
    return
}

if (-not (Test-NetworkPath -Path $userProfile2Path -LogPath $logPath)) {
    Write-Host "[ERROR] Cannot access network path for user profile 2: $userProfile2Path" -ForegroundColor Red
    Add-Content $logPath "[FATAL] Script terminating due to network path access failure"
    Read-Host "`nPress Enter to exit"
    return
}

if (-not (Test-NetworkPath -Path $hDrivePath -LogPath $logPath)) {
    Write-Host "[ERROR] Cannot access network path for H Drive: $hDrivePath" -ForegroundColor Red
    Add-Content $logPath "[FATAL] Script terminating due to network path access failure"
    Read-Host "`nPress Enter to exit"
    return
}

if (-not (Test-NetworkPath -Path $hrFolderRoot -LogPath $logPath)) {
    Write-Host "[ERROR] Cannot access network path for HR folder: $hrFolderRoot" -ForegroundColor Red
    Add-Content $logPath "[FATAL] Script terminating due to network path access failure"
    Read-Host "`nPress Enter to exit"
    return
}

Add-Content $logPath "[SUCCESS] All network paths validated successfully"
Write-Host "[SUCCESS] Network path validation completed" -ForegroundColor Green

# === SET PERMISSIONS ===
Write-Host "`n=== SETTING PERMISSIONS ==="  -ForegroundColor Cyan
Write-Host "[1/3] Processing User Profile Path..." -ForegroundColor Yellow
Set-FolderPermissions -Path $userProfilePath -Account $accountName -LogPath $logPath

Write-Host "[2/3] Processing User Profile 2 Path..." -ForegroundColor Yellow
Set-FolderPermissions -Path $userProfile2Path -Account $accountName -LogPath $logPath

Write-Host "[3/3] Processing H Drive Path..." -ForegroundColor Yellow
Set-FolderPermissions -Path $hDrivePath -Account $accountName -LogPath $logPath

# === DELETE FOLDERS ===
Write-Host "`n=== DELETING FOLDERS ==="  -ForegroundColor Cyan
Write-Host "[1/2] Deleting User Profile Path..." -ForegroundColor Yellow
Remove-FolderIfExists -Path $userProfilePath -LogPath $logPath

Write-Host "[2/2] Deleting User Profile 2 Path..." -ForegroundColor Yellow
Remove-FolderIfExists -Path $userProfile2Path -LogPath $logPath

# === COPY H DRIVE ===
Write-Host "`n=== COPYING H DRIVE ==="  -ForegroundColor Cyan
Write-Host "Copying H Drive to HR folder..." -ForegroundColor Yellow
$copySucceeded = Copy-HDrive -Source $hDrivePath -Destination $hrHDriveFolder -LogPath $logPath

if ($copySucceeded) {
    try {
        Remove-Item -Path $hDrivePath -Recurse -Force -ErrorAction Stop
        Add-Content -Path $logPath -Value "Deleted original H Drive folder: ${hDrivePath}"
    } catch {
        Add-Content -Path $logPath -Value "[ERROR] Failed to delete original H Drive folder: $_"
    }
} else {
    Add-Content -Path $logPath -Value "[INFO] H Drive copy failed. Original not deleted."
}

Read-Host "`nPress Enter to exit"