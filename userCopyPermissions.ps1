# Prompt for the source username
$sourceUser = Read-Host "Enter the source username (groups will be copied FROM this user)"

# Prompt for the target username
$targetUser = Read-Host "Enter the target username (groups will be copied TO this user)"

# Get AD group memberships for the source user
$groupsSourceUser = Get-ADPrincipalGroupMembership -Identity $sourceUser

# Create an array to hold copied groups
$copiedGroups = @()

# Add groups from the source user to the target user
foreach ($group in $groupsSourceUser) {
    $copiedGroups += $group.Name
    Add-ADGroupMember -Identity $group.Name -Members $targetUser
}

# Display the copied groups
Write-Host "Copied Groups from $($sourceUser) to $($targetUser):"
$copiedGroups | ForEach-Object { Write-Host $_ }