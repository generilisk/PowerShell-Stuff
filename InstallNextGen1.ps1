# Step 1: Run VC_redist.x64.exe as admin
Start-Process -FilePath "\\shastahealth.org\shared\ITS\Store\Software\ODBC\17\VC_redist.x64.exe" -ArgumentList "/q" -Verb RunAs -Wait

# Step 2: Install Sqlnclix64.msi -> Redo this one
Start-Process -FilePath "msiexec.exe" -ArgumentList "/i \\shastahealth.org\shared\ITS\Store\Software\ODBC\17\Sqlnclix64.msi /qn" -Wait

# Step 3: Run Create_ODBC.exe from Prod folder
Start-Process -FilePath "\\shastahealth.org\shared\ITS\Store\Software\ODBC\17\Prod\Create_ODBC.exe" -Wait

# Step 4: Run Create_ODBC.exe from Report folder
Start-Process -FilePath "\\shastahealth.org\shared\ITS\Store\Software\ODBC\17\Report\Create_ODBC.exe" -Wait

# Step 5: Merge ODBCx64.reg and ODBCx86.reg into the registry
regedit.exe /s "\\shastahealth.org\shared\ITS\Store\Software\ODBC\17\ODBCx64.reg"
regedit.exe /s "\\shastahealth.org\shared\ITS\Store\Software\ODBC\17\ODBCx86.reg"

# Step 6: Copy NGConfig.ini into c:\windows
Copy-Item "\\ngroot\NextGenRoot\NGUtilitiesConfig\Prod\NGConfig.ini" "C:\Windows" -Force

# Step 7: Give "Users" group Modify permissions for NGConfig.ini
    # Get the current ACL of the file
    $acl = Get-Acl -Path "C:\Windows\NGConfig.ini"

    # Define the permissions for the Users group
    $permission = "Users","Modify","None","Allow"

    # Convert the permission string to a FileSystemRights object
    $FileSystemRights = [System.Security.AccessControl.FileSystemRights] $permission[1]

    # Create an Access Control Type object
    $AccessControlType = [System.Security.AccessControl.AccessControlType]::Allow

    # Create a new Access Rule
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($permission[0], $FileSystemRights, $AccessControlType)

    # Add the new rule to the ACL
    $acl.AddAccessRule($rule)

    # Apply the modified ACL to the file
    Set-Acl -Path "C:\Windows\NGConfig.ini" -AclObject $acl

# Step 8: Install vcredist_x64.exe as admin
Start-Process -FilePath "\\shastahealth.org\shared\ITS\Store\Software\NextGen\Prereqs\vcredist_2005\vcredist_x64.exe" -ArgumentList "/q" -Verb RunAs -Wait

# Step 9: Install vcredist_x86.exe as admin
Start-Process -FilePath "\\shastahealth.org\shared\ITS\Store\Software\NextGen\Prereqs\vcredist_2005\vcredist_x86.exe" -ArgumentList "/q" -Verb RunAs -Wait

# Step 10: Enable MSMQ-Container feature
Dism /online /Enable-Feature /FeatureName:MSMQ-Container

# Step 11: Enable MSMQ-ADIntegration feature
Dism /online /Enable-Feature /FeatureName:MSMQ-ADIntegration

# Step 12: Enable NetFx3 feature
Dism /online /Enable-Feature /FeatureName:NetFx3

# Step 13: Add NetFx3 capability from source
#DISM /Online /Add-Capability /CapabilityName:NetFx3~~~~ /Source:"\\Rdmps02\sources\os\win10 x64 ent\sources\sxs"

# Step 14: Add NetFx3 capability from source
#Add-WindowsCapability -Online -Name NetFx3~~~~ -Source "\\Rdmps02\sources\os\zWin10DOTNET"

# Step 15: Copy EDR Files
$sourcePathEDR = "\\ngroot\NextGenRoot\Prod\EDR"
$destinationPathEDR = "C:\Nextgen"
robocopy $sourcePathEDR $destinationPathEDR /E

# Step 16: Copy Custom Provider View
$sourcePathEHR = "\\ngroot\nextgenroot\Prod\EHR"
$destinationPathEHR = "C:\NextGen"
robocopy $sourcePathEHR $destinationPathEHR /E

# Step 17: Install Custom Font
$scriptPathFont = "\\shastahealth.org\shared\ITS\Store\Software\NextGen\scripts\NGFont\install_font.ps1"
Start-Process -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList "-executionpolicy bypass -File `"$scriptPathFont`"" -Verb RunAs

# Step 18: Copy Nextgen Shortcut
$sourcePathShortcut = "\\shastahealth.org\shared\ITS\Store\Software\Shortcuts\NextGen 5.lnk"
$destinationPathShortcut = "C:\Users\Public\Desktop"
Copy-Item -Path $sourcePathShortcut -Destination $destinationPathShortcut

# Step 19: Run setup.exe as admin
Start-Process -FilePath "\\ngroot\NextGenRoot\Install\Install\NextGen Setup\setup.exe" -Verb RunAs