<#
users.csv should have "UserName,BadgeNumber" header and format. For example:

UserName,BadgeNumber
alagrant,10248
dennedry,10001
esattler,10249
henrywu,10003
imalcolm,10250
jhammond,10002

#>


# Import the Active Directory module
Import-Module ActiveDirectory

# Specify the path to the CSV file containing user information
$csvPath = "C:\PShellData\users.csv"

# Read the CSV file
$users = Import-Csv $csvPath

# Loop through each user in the CSV file
foreach ($user in $users) {
    # Retrieve the user by their username
    $adUser = Get-ADUser -Filter "SamAccountName -eq '$($user.UserName)'" -Properties extensionAttribute3

    if ($adUser) {
        # User found, update the extensionAttribute3 attribute
        $adUser.extensionAttribute3 = $user.BadgeNumber
        Set-ADUser -Instance $adUser
        Write-Host "Updated extensionAttribute3 for user $($adUser.Name)"
    }
    else {
        Write-Host "User with username $($user.UserName) not found"
    }
}
pause