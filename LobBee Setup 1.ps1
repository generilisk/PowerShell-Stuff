Set-Itemproperty -path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'AutoAdminLogon' -value "1"
Set-Itemproperty -path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultDomainName' -value "shastahealth.org"
Set-Itemproperty -path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultPassword' -value "Us3r1035"
Set-Itemproperty -path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'DefaultUserName' -value "LobbyUser"

Copy-Item '\\rdmps01\sources\images\S2LobbyBees\WhiteLogo.png' 'C:\ProgramData\S2\MagicMonitor\magic_views\media\WhiteLogo.png'
Copy-Item '\\rdmps01\sources\images\S2LobbyBees\slideshow.mp4' 'C:\ProgramData\S2\MagicMonitor\magic_views\media\slideshow.mp4'
Copy-Item '\\rdmps01\sources\images\S2LobbyBees\media.ini' 'C:\ProgramData\S2\MagicMonitor\magic_views\media\media.ini'
Copy-Item '\\rdmps01\sources\images\S2LobbyBees\Magic Monitor.lnk' 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\Magic Monitor.lnk'

net localgroup "Administrators" "schc\LobbyUser" /add

$creds = Get-Credential
$NamePostChange = Read-Host "What is the new PC name?"
Rename-Computer -DomainCredential $creds -NewName $NamePostChange -Restart