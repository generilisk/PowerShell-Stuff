# Prompt for the badge number to search for
$badgeNumber = Read-Host "Enter the badge number you want to search for"

# Search for users whose extensionAttribute3 matches the entered badge number
$user = Get-ADUser -Filter { extensionAttribute3 -eq $badgeNumber }

# Output the user(s) found
if ($user) {
    foreach ($u in $user) {
        Write-Host "$badgeNumber belongs to $($u.Name)"
    }
} else {
    Write-Host "No user found with the badge number: $badgeNumber"
}
Read-Host "Press Enter to continue..."