Clear-Host
$Beehive = @(
'EMEDB001',
'RDENB001',
'RFAMB001',
'RFELB001',
'RMATB001',
'RPCNB001',
'RPEDB001',
'RRESB001',
'RTMDB001',
'RURGB001',
'RXRYB001',
'SMEDB001'
)

ForEach($Bee in $Beehive)
{
    xcopy C:\Users\whughes\Videos\slideshow.mp4 "\\$Bee\c`$\ProgramData\S2\MagicMonitor\magic_views\media" /R /U /Y /D 
    shutdown /t 0 /r /f /m \\$bee
}