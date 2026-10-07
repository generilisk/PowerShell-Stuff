# Prompt user for AD group name
$groupName = Read-Host "Enter the Active Directory group name"

# Get users in the specified group
try {
    $groupMembers = Get-ADGroupMember -Identity $groupName -Recursive | Select-Object Name | Sort-Object Name
    if ($groupMembers.Count -eq 0) {
        Write-Host "No users found in the specified group."
    }
    else {
        Write-Host "Users in the group '$groupName':"
        $groupMembers | Format-Table -AutoSize
    }
}
catch {
    Write-Host "Error: $_"
}