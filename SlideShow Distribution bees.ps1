clear
$Beehive = @(
#'RXRYB001',
'RFAMB001'
)

ForEach($Bee in $Beehive)
{
#    xcopy C:\Users\whughes\Videos\slideshow.mp4 "\\$Bee\c`$\ProgramData\S2\MagicMonitor\magic_views\media"# /R /U /Y /D 
    shutdown /t 0 /r /f /m \\$bee
}
