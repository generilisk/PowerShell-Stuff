#$Locations = LobbyPCN, LobbyMat, LobbySLC

$UsualPath = "H:\PowerShell\Lobby Slideshow.ppsx"
$Usual = Get-Item -Path $UsualPath
if(Test-Path $UsualPath){
        Copy-Item -Path $UsualPath -Destination "H:\PowerShell\test\Lobby Slideshow.ppsx" -Force
        Write-Host "I think it worked?"
    } else{
        Write-Host $UsualPath "Doesn't Exist!"
    }
#$Usual |FL  