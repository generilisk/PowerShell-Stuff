Function SetupODBC{
    net use j: \\Rdmps01\sources\software\ODBC\17
    J:\VC_redist.x64.exe
    J:\sqlnclix64.msi
    J:\Prod\Create_ODBC.exe
    J:\Report\Create_ODBC.exe
}
Function SetupNGConfig{
    cmd /c copy "\\ngroot\NextGenRoot\NGUtilitiesConfig\Prod\NGConfig.ini" C:\windows
    $NewAcl = Get-Acl -Path "C:\windows\NGConfig.ini"
    # Set properties
    $identity = "BUILTIN\Users"
    $fileSystemRights = "Modify"
    $type = "Allow"
    # Create new rule
    $fileSystemAccessRuleArgumentList = $identity, $fileSystemRights, $type
    $fileSystemAccessRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $fileSystemAccessRuleArgumentList
    # Apply new rule
    $NewAcl.SetAccessRule($fileSystemAccessRule)
    Set-Acl -Path "C:\windows\NGConfig.ini" -AclObject $NewAcl
}
Function SetupMSMQ{
    Dism /online /Enable-Feature /FeatureName:MSMQ-Container
    Dism /online /Enable-Feature /FeatureName:MSMQ-ADIntegration
}
Function SetupDotNet3Point5{
    Dism /online /Enable-Feature /FeatureName:NetFx3
    DISM /Online /Add-Capability /CapabilityName:NetFx3~~~~ /Source:"\\Rdmps01\sources\os\win10 x64 ent\sources\sxs"
    Add-WindowsCapability -Online -Name NetFx3~~~~ -Source \\Rdmps01\sources\os\zWin10DOTNET
}


SetupODBC
SetupNGConfig
SetupMSMQ
SetupDotNet3Point5