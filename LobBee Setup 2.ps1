reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableLockWorkstation /d 0x1 /t REG_DWORD
reg add 'HKCU\Control Panel\Desktop' /v ScreenSaverIsSecure /d 1 /t REG_SZ

net localgroup "Administrators" "schc\LobbyUser" /delete
shutdown -t 0 -r