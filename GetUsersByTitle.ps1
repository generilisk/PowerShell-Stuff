# Import the Active Directory module if not already loaded
Import-Module ActiveDirectory

# Define the title you want to filter by
$titleFilter = "*visit coordinator*"

# Get a list of users with titles containing the specified filter
$users = Get-ADUser -Filter {Title -like $titleFilter} -Properties Department, Title, OfficePhone

# Display the user information
$users | Select-Object Name, Department, Title, OfficePhone | Format-Table -AutoSize
