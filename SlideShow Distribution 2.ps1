#$Complist = @("LOBBYFEL", "LOBBYFPB", "LOBBYMAT", "LOBBYPCN", "LOBBYPEDS", "LOBBYRDEN", "LOBBYRES", "LOBBYSLC", "LOBBYUC")
#$Complist = @("RITSW24","ABC123","RITSW77")
$Complist = @("LOBBYFEL")
$SourceLoc = 'C:\Sign\'
$SourceFile = 'Lobby Sign.ppsx'
$DestLoc = 'C:\sign\'
$DestFile = 'Lobby Sign.ppsx'
$SourcePath = "$SourceLoc$SourceFile"
$NetLoc = 'C$\sign\'


function TestAndCopy {
    
    ForEach($Device In $Complist) {
    	$DestPath = "$DestLoc$DestFile"
        $NetPath = "\\$Device\$NetLoc$DestFile"
        Write-Host "Copying to $Device"
        $Session = New-PSSession -ComputerName "$Device"
        If(Test-Connection $Device -Quiet) {
            If(Test-Path $NetPath -PathType Container){
            }Else {
            New-Item -Path "\\$Device\c$\" -Name "sign" -ItemType "directory"
            }
            Try {
                Copy-Item $SourcePath -Destination $DestPath -ToSession $Session
                Write-Host "Successfully copied to $Device"
            }Catch {
                "Unable to copy to $Device"
            }

        }Else {
            Write-Host "Connection to "$Device" has failed."
        }
        Write-Host ""
        Restart-Computer $Device
	}
}
cls

If(Test-Path $SourcePath -PathType Leaf) {
    TestAndCopy
    } else {
    Write-Host "Lobby Sign file not found"
    }

#"LOBBYRAD"