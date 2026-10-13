<#
.SYNOPSIS
Copies all AD group memberships from one user to another.
.DESCRIPTION
Adds the target user to every group the source user belongs to. Existing
memberships on the target are left alone. Use -WhatIf to preview.
.EXAMPLE
.\Copy-UsersGroupMembership.ps1 -SourceUser bruwayne -TargetUser diprince
.EXAMPLE
.\Copy-UsersGroupMembership.ps1 -SourceUser bruwayne -TargetUser diprince -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceUser,
    
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetUser
)

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

# Helper function to get username if missing
function Get-UsernameIfMissing {
    param(
        [string]$username,
        [string]$prompt
    )
    
    if ([string]::IsNullOrWhiteSpace($username)) {
        $username = Read-Host -Prompt $prompt
    }
    return $username
}

# Get usernames using helper function
$SourceUser = Get-UsernameIfMissing -username $SourceUser -prompt "Enter the source username (groups will be copied FROM this user)"
$TargetUser = Get-UsernameIfMissing -username $TargetUser -prompt "Enter the target username (groups will be copied TO this user)"

if ([string]::IsNullOrWhiteSpace($SourceUser) -or [string]::IsNullOrWhiteSpace($TargetUser)) {
    Write-Warning "Both a source and a target username are required."
    return
}

# Verify both users exist and get the source user's groups
try {
    $null = Get-ADUser -Identity $TargetUser -ErrorAction Stop
    $groupsSourceUser = Get-ADPrincipalGroupMembership -Identity $SourceUser -ErrorAction Stop
}
catch {
    Write-Warning "Could not look up user: $($_.Exception.Message)"
    return
}

# Create an array to hold copied groups
$copiedGroups = @()
$failedGroups = @()

# Add groups from the source user to the target user
foreach ($group in $groupsSourceUser) {
    try {
        if ($PSCmdlet.ShouldProcess($TargetUser, "Add to group $($group.Name)")) {
            Add-ADGroupMember -Identity $group -Members $TargetUser -ErrorAction Stop
            $copiedGroups += $group.Name
        }
    }
    catch {
        if ($_.Exception.Message -like "*already a member*") {
            Write-Host "$TargetUser is already a member of $($group.Name), no copy needed."
        }
        else {
            # Handle other errors here
            Write-Warning "Failed to add $TargetUser to $($group.Name): $($_.Exception.Message)"
            $failedGroups += $group.Name
        }
    }
}

# Display the copied groups
Write-Host "Copied Groups from $($SourceUser) to $($TargetUser):"
$copiedGroups | ForEach-Object { Write-Host $_ }
if ($failedGroups.Count -gt 0) {
    Write-Host "Failed groups:"
    $failedGroups | ForEach-Object { Write-Host $_ }
}
Read-Host -Prompt "Press Enter to exit"