# Step 1: Run VC_redist.x64.exe as admin
Start-Process -FilePath "\\rdmps02\sources\Software\ODBC\17\VC_redist.x64.exe" -ArgumentList "/q" -Verb RunAs -Wait

# Step 2: Install Sqlnclix64.msi
Start-Process -FilePath "msiexec.exe" -ArgumentList "/i \\rdmps02\sources\Software\ODBC\17\Sqlnclix64.msi /qn" -Wait

# Step 3: Run Create_ODBC.exe from Prod folder
Start-Process -FilePath "\\rdmps02\sources\Software\ODBC\17\Prod\Create_ODBC.exe" -Wait

# Step 4: Run Create_ODBC.exe from Report folder
Start-Process -FilePath "\\rdmps02\sources\Software\ODBC\17\Report\Create_ODBC.exe" -Wait

# Step 5: Merge ODBCx64.reg and ODBCx86.reg into the registry
regedit.exe /s "\\rdmps02\sources\Software\ODBC\17\ODBCx64.reg"
regedit.exe /s "\\rdmps02\sources\Software\ODBC\17\ODBCx86.reg"

# Step 6: Copy NGConfig.ini into c:\windows
Copy-Item "\\ngroot\NextGenRoot\NGUtilitiesConfig\Prod\NGConfig.ini" "C:\Windows" -Force

# Step 7: Give "Users" group Modify permissions for NGConfig.ini
$filePath = "C:\Windows\NGConfig.ini"
$acl = Get-Acl $filePath
$permission = "Users", "Modify", "ContainerInherit, ObjectInherit", "None", "Allow"
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.SetAccessRule($rule)
Set-Acl $filePath $acl

# Step 8: Install vcredist_x64.exe as admin
Start-Process -FilePath "\\rdmps02\sources\Software\NextGen\Prereqs\vcredist_2005\vcredist_x64.exe" -ArgumentList "/q" -Verb RunAs -Wait

# Step 9: Install vcredist_x86.exe as admin
Start-Process -FilePath "\\rdmps02\sources\Software\NextGen\Prereqs\vcredist_2005\vcredist_x86.exe" -ArgumentList "/q" -Verb RunAs -Wait

# Step 10: Enable MSMQ-Container feature
Dism /online /Enable-Feature /FeatureName:MSMQ-Container

# Step 11: Enable MSMQ-ADIntegration feature
Dism /online /Enable-Feature /FeatureName:MSMQ-ADIntegration

# Step 12: Enable NetFx3 feature
Dism /online /Enable-Feature /FeatureName:NetFx3

# Step 13: Add NetFx3 capability from source
DISM /Online /Add-Capability /CapabilityName:NetFx3~~~~ /Source:"\\Rdmps02\sources\os\win10 x64 ent\sources\sxs"

# Step 14: Add NetFx3 capability from source
Add-WindowsCapability -Online -Name NetFx3~~~~ -Source "\\Rdmps02\sources\os\zWin10DOTNET"

# Step 15: Run setup.exe as admin
Start-Process -FilePath "\\ngroot\NextGenRoot\Install\Install\NextGen Setup\setup.exe" -Verb RunAs
