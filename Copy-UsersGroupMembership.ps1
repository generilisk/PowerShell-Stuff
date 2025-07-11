param(
    [string]$SourceUser,
    [string]$TargetUser
)

# Helper function to get username if missing
function Get-UsernameIfMissing {
    param(
        [string]$username,
        [string]$prompt
    )
    
    if ([string]::IsNullOrEmpty($username)) {
        $username = Read-Host -Prompt $prompt
    }
    return $username
}

# Get usernames using helper function
$sourceUser = Get-UsernameIfMissing -username $SourceUser -prompt "Enter the source username (groups will be copied FROM this user)"
$targetUser = Get-UsernameIfMissing -username $TargetUser -prompt "Enter the target username (groups will be copied TO this user)"

# Get AD group memberships for the source user
$groupsSourceUser = Get-ADPrincipalGroupMembership -Identity $sourceUser

# Create an array to hold copied groups
$copiedGroups = @()

# Add groups from the source user to the target user
foreach ($group in $groupsSourceUser) {
    try {
        Add-ADGroupMember -Identity $group.Name -Members $targetUser
        $copiedGroups += $group.Name
    }
    catch {
        if ($_.Exception.Message -like "*already a member*") {
            Write-Host "$targetUser is already a member of $group, no copy needed."
        }
        else {
            # Handle other errors here
            Write-Host "An error occurred: $($_.Exception.Message)"
        }
    }
}

# Display the copied groups
Write-Host "Copied Groups from $($sourceUser) to $($targetUser):"
$copiedGroups | ForEach-Object { Write-Host $_ }
Read-Host -Prompt "Press any key to continue"