<#
.Synopsis
    Offboarding script for handling AD user information and Exchange content searches.
.DESCRIPTION
    Retrieves AD user information and performs a content search in Exchange for offboarding purposes.
.EXAMPLE
    .\export-ExchangeData.ps1 -accountName "jdoe"
#>

#Script Name: Offboarding Script
#Created by: Will Hughes
#Date: 2024-10-07
#Patch Notes: 


param(
    [string]$accountName
)

# Import AD module if not already imported
Import-Module ActiveDirectory

# Install ExchangeOnlineManagement module if not installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -AllowClobber -Force
}
Import-Module ExchangeOnlineManagement

# Check if $accountName was provided
if (-not $accountName) {
    throw "Username is required."
}

# Retrieve the AD user and display the DisplayName and mail
try {
    $adUser = Get-ADUser -Identity $accountName -Properties DisplayName, Mail
    $adMail = $adUser.Mail
																   
    if ($adUser) {
        Write-Host "$accountName is linked to: $($adUser.DisplayName), $adMail"
    } else {
        Write-Host "Account $accountName not found in Active Directory."
        exit
    }
} catch {
    Write-Host "Error retrieving user $accountName`: $_"
    exit
}

# Confirm if the AD user is correct - default to "no"
$confirmation = Read-Host "Is this correct? (y/n)"
if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
    Write-Host ""
    Write-Host -ForegroundColor Black -BackgroundColor Red "---------- !!!!! Operation cancelled  !!!!! ----------"
    Write-Host ""
    exit
}

# If confirmed, continue with offboard operations
Write-Host "Continuing with offboard operations for $accountName..."


<# This does not support MFA
# Prompt for credentials
$credentials = Get-Credential
# Connect to Exchange Online using the credential object
Connect-ExchangeOnline -Credential $credentials -ShowProgress $true
#>

# Connect to Exchange Online for mailbox-related tasks
Connect-ExchangeOnline

# Connect to Microsoft Purview (Compliance Center) for compliance searches
Connect-IPPSSession

# Create a content search
$searchName = $accountName
$description = "Off-boarding for $($adUser.DisplayName)"
New-ComplianceSearch -Name $searchName -ExchangeLocation $adMail -Description $description

# Start the content search
Start-ComplianceSearch -Identity $searchName

# Poll the search status every 10 seconds until it is completed
$status = Get-ComplianceSearch -Identity $searchName
while ($status.Status -ne "Completed") {
    Write-Host "Search is still running... Waiting for completion."
    Start-Sleep -Seconds 10
    $status = Get-ComplianceSearch -Identity $searchName
}

# Once the search is completed, export the search results
Write-Host "Search completed! Proceeding with export..."
New-ComplianceSearchAction -SearchName $searchName -Export -ExchangeArchiveFormat SingleFolderPst

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false