<#
.SYNOPSIS
    This PowerShell script retrieves the next available computer name following a specific naming convention in Active Directory and renames the computer accordingly.

.DESCRIPTION
    The script connects to Active Directory and retrieves computer objects whose names follow the convention specified by the $Prefix parameter. 
    It then identifies the next available sequential number in the sequence and generates the next computer name accordingly. 
    If no existing computer names match the pattern, it defaults to the next sequential number after the highest number found.
    After determining the new computer name, the script prompts the user to confirm before renaming the computer and initiating a restart.

.PARAMETER Prefix
    Specifies the prefix part of the computer names. Default is "AMEDW0".

.EXAMPLE
    .\Get-NextComputerName.ps1 -Prefix "AMEDW0"
    Retrieves the next available computer name following the "AMEDW0" naming convention and renames the computer accordingly.

.NOTES
    Script Name: Get-NextComputerName.ps1
    Created by: Will Hughes
    Date: 2024-05-07
    Patch Notes: 
#>

Clear-Host

param (
    [string]$Prefix = "AMEDW0"
)

# Search for all computer objects in Active Directory
$computers = Get-ADComputer -Filter "Name -like '$Prefix*'" | Select-Object -ExpandProperty Name

# Output the list of computer names (for debugging)
Write-Output "Computers:"
$computers = $computers | Sort-Object | ForEach-Object { Write-Output $_ }
$computers

# If no computers found, default to the next sequential number after the highest number found
if ($computers.Length -eq 0) {
    $highestNumber = 0
} else {
    # Extract the numbers from the existing computer names
    $numbers = $computers | ForEach-Object { [int]($_ -replace '[^\d]','') }

    # Find the missing sequential number
    $highestNumber = ($numbers | Measure-Object -Maximum).Maximum

    $nextNumber = 1
    while ($numbers -contains $nextNumber) {
        $nextNumber++
    }
}

# Format the next computer name
$nextComputerName = "{0}{1:D3}" -f $Prefix, $nextNumber

# Output the next computer name
Write-Output "Next computer name:"
Write-Output $nextComputerName

# Prompt the user to continue before restarting
Read-Host "Press Enter to rename and restart the computer..."

# Rename the computer and restart
#Rename-Computer -NewName $nextComputerName -Force -Restart
