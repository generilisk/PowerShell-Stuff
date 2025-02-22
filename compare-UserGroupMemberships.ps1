<#
.Synopsis
    Compare group memberships between two users.
.DESCRIPTION
    This script compares the group memberships of two Active Directory users.
.PARAMETER User1
    The first user for comparison.
.PARAMETER User2
    The second user for comparison.
.EXAMPLE
    Compare-UserGroups.ps1
    This example prompts the user to enter two usernames for comparison and then displays the comparison of group memberships between those users.
.EXAMPLE
    Compare-UserGroups.ps1 -User1 "user1" -User2 "user2"
    This example compares the group memberships between "user1" and "user2".
#>

#Script Name: Compare-UserGroups.ps1
#Created by: Will Hughes
#Date: 2024-05-02
#Patch Notes:

param (
    [string]$User1,
    [string]$User2
)

if (-not $User1) {
    $User1 = Read-Host "Enter the first username for comparison"
}

if (-not $User2) {
    $User2 = Read-Host "Enter the second username for comparison"
}

# Get AD group memberships for User1
$groupsUser1 = (Get-ADPrincipalGroupMembership -Identity $User1 | Select-Object -ExpandProperty Name) | Sort-Object

# Get AD group memberships for User2
$groupsUser2 = (Get-ADPrincipalGroupMembership -Identity $User2 | Select-Object -ExpandProperty Name) | Sort-Object

# Display group memberships comparison
Write-Host "`n**Group Memberships Comparison between $($User1) and $($User2):**"

# Groups present in both users
$groupsBoth = $groupsUser1 | Where-Object { $groupsUser2 -contains $_ }
if ($groupsBoth) {
    Write-Host "`n**Present in both users:**"
    $groupsBoth | ForEach-Object {
        Write-Host $_
    }
} else {
    Write-Host "`nNo groups found in common between $($User1) and $($User2)."
}

# Groups present only in User1
Write-Host "`n**Present only in $($User1):**"
$groupsOnlyUser1 = Compare-Object $groupsUser1 $groupsUser2 | Where-Object {$_.SideIndicator -eq "<="}
$groupsOnlyUser1 | ForEach-Object {
    Write-Host $_.InputObject
}

# Groups present only in User2
Write-Host "`n**Present only in $($User2):**"
$groupsOnlyUser2 = Compare-Object $groupsUser1 $groupsUser2 | Where-Object {$_.SideIndicator -eq "=>"}
$groupsOnlyUser2 | ForEach-Object {
    Write-Host $_.InputObject
}

Read-Host -Prompt "`nPress enter to continue..."