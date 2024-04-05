param(
    [Parameter(Mandatory = $True)]
    [string[]]$LogName,
    [string[]]$computerName = $env:COMPUTERNAME,
    [int32]$Newest = 500

)
foreach ($target in $computerName) {
    #doSomeStuff
}
#Select-Object example
Get-Service | Select-Object DisplayName, Status, StartType
if (3 -lt 4) {
    Write-Host "Math's not broke"
}
elseif (3 -eq 4) {
    Write-Host"Math broke some"
}
else {
    Write-Host"Math done broke" 
}

# $_ means current object in the pipeline
#ForEach-Object has an alias of ForEach. ForEach is also a seperate cmdlet. pshell determines via context.
2, 5, 6, 8, 9 | ForEach-Object { $_ * 3 }
#is the same as 2, 5, 6, 8, 9 | ForEach { $_ * 3 }

