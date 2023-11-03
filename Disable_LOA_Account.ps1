Function SyncAD{
# Sync AD with AzureAD with ADCONNECT
    param(
        $creds,
        $ADConnectServer    
    )
    # Import Modules
    $session = New-PSSession -ComputerName $ADConnectServer -Credential $creds
    Invoke-Command -Session $session -ScriptBlock {Import-Module -Name 'ADSync'}
    Invoke-Command -Session $session -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Initial}
    Remove-PSSession $session
}

Function DisableAccount{
  param($username)
    try{
        # Block ActiveSync Devices tied to the Account
        $MobileDevices = Get-MobileDevice -Mailbox $username
        foreach($device in $MobileDevices){
          Set-CASMailbox -Identity $username -ActiveSyncBlockedDeviceIDs $device.DeviceId
          write-host "Blocked $device.DeviceId from connecting with ActiveSync" -ForegroundColor Green
        }
        write-host "Disabling Exchange ActiveSync Devices was successful" -ForegroundColor Green

    }catch{
        write-host "Disabling Exchange ActiveSync Devices Failed, Please log into ECP and do it manually!!!" -ForegroundColor Red
    }

    #Block ActiveSync, OWA, OWA for Devices
    try{
        Set-CASMailbox -Identity $username -ActiveSyncEnabled $false
        write-host "Disabling Exchange ActiveSync was successful" -ForegroundColor Green
        Set-CASMailbox -Identity $username -OWAforDevicesEnabled $false
        write-host "Disabling Exchange OWA for Devices was successful" -ForegroundColor Green
        Set-CasMailbox -Identity $username -OWAEnabled $false
        write-host "Disabling Exchange OWA was successful" -ForegroundColor Green
    }catch{
        write-host "Disabling Exchange ActiveSync, OWA for Devices, or OWA has Failed, Please log into ECP and do it manually!!!" -ForegroundColor Red
    }
}

Function DisableExchangeAccounts{
    param($username,$eol,$UserCredential)
    
    if($eol -eq $true){
        
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        Import-PSSession $Session  -DisableNameChecking -AllowClobber
        DisableAccount -username "$username@shastahealth.org"
        Remove-PSSession $Session
    }else{
        DisableAccount -username $username
    }
}

Function DisableUserAccount{
    Param(
        $username
    )
    # Disable Active Directory Account
    try{
        Disable-ADAccount -Identity $username
        write-host "Active Directory Account has been disabled." -ForegroundColor Green
    }catch{
        write-host "Failed to disable Active Directory Account. Please do it manually!!!" -ForegroundColor Red
    }
}

Function ImportADCommands{
    param($Server)

    $sessionAD = New-PSSession -ComputerName $Server
    Invoke-Command { Import-Module ActiveDirectory } -Session $sessionAD
    Export-PSSession -Session $sessionAD -CommandName *-AD* -OutputModule RemAD -AllowClobber -Force | Out-Null
    Remove-PSSession -Session $sessionAD
}

Function FindLocation{
    param($users)
    #Location is set to On Prem by Default
    $Location = $true
    $ADUser = Get-ADUser -Identity $users -Properties *

    if($ADUser.msExchRemoteRecipientType){
        # User is in the cloud
        $Location = $false
    }
    return $Location
}

Function ConfirmDisableCloud{
    param($username)
    
    Try{
        write-host "Installing AzureAD Module.." -ForegroundColor Yellow
        Install-Module -Name AzureAD
    }Catch{ 
        Write-host "Failed installing AzureAD Module. Script may still connect." -ForegroundColor Magenta
    }
    write-host "Type in Office 365 Credentials to Connect to AzureAD" -ForegroundColor Yellow
    Connect-MsolService
    $upn = "$username@shastahealth.org"
    $user = Get-MsolUser -UserPrincipalName $upn
    Do{ 
        write-host "Syncing....."
        Start-Sleep -Seconds 30
     }While($user.BlockCredential -eq $true)
    
    return $true
}

Function ImportExchangeCommands{
    param($creds)
    $UserCredential = $creds
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://emails01.shastahealth.org/PowerShell/ -Authentication Kerberos -Credential $UserCredential
    Import-PSSession $Session -AllowClobber
    #Exit-PSSession
}

$ADConnectServer = "ADFS03"
$userAcct = Read-Host -Prompt "What Username do you want to disable?"
write-host "You typed:   $userAcct   " -ForegroundColor White -BackgroundColor DarkMagenta
$confirmation = Read-Host -Prompt "Is the above username correct? Press 'Y' to confirm."

if($confirmation.ToUpper() -eq "Y" -and $userAcct -ne $null){
    
    $creds = Get-Credential -Message "Enter in Administrator AD Credentials"
    $o365creds = Get-Credential -Message "Type in Office 365 Credentials with @shastahealth.org"
    $EmailFrom = $o365creds.UserName
    ImportADCommands -Server "NtSCHCDCMAIN" -creds $creds
    ImportExchangeCommands -creds $creds
    $UserLocation = FindLocation -users $userAcct

    If($UserLocation -eq $true){
        #User is on Prem
        write-host "Users Mailbox is On-Prem" -ForegroundColor Yellow
        DisableExchangeAccounts -username $userAcct -eol $false -UserCredential $creds           
    }else{
        # User is in the cloud
        write-host "Users Mailbox is in the Cloud" -ForegroundColor Yellow
        DisableExchangeAccounts -username $userAcct -eol $true -UserCredential $o365creds
    }

    #Disable AD Account
    DisableUserAccount -username $userAcct   
    
   
    # Syn no matter where Account is to disable account online services
    Write-Host "Performing a Sync to Azure to Block Account in Cloud Servies..." -ForegroundColor Green
    # Start a Azure Sync
    SyncAD -creds $o365creds -ADConnectServer $ADConnectServer
        
    $ADDisableResult = ConfirmDisableCloud -username $userAcct
    if($ADDisableResult -eq $true){
        write-host "Sync Completed. Account $userAcct has been disabled in the cloud." -ForegroundColor Green
    }else{
        write-host "Sync Failed. Please disable account $userAcct manually and rerun a sync to Azure." -ForegroundColor Red   
    }
    #MoveDisabledUser -username $userAcct -credentials $creds -Path $DisabledOU  # Still need to ad exporting of Email before move.
}else{
    write-host "Disabling Canceled!!" -ForegroundColor Red
}