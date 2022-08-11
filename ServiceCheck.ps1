Clear-Host
$medlist = "medical01","medical02","medical04","medical06","medical07","medical08","medical09","medical10","medical11","medical12","medical13","medical14"
$Service = Read-Host -Prompt "Which Service?"

$Cred = Get-Credential
ForEach ($medserver in $medlist)
{

    $ScriptBlock =
    {
        param($Service)
        Get-Service -Name "$Service"
    }
    $ServiceStatus = Invoke-Command -ComputerName $medserver -Credential $Cred -ScriptBlock $ScriptBlock -ArgumentList "$Service" | Select-Object Status
    Write-Host $medserver -NoNewline
    Write-Host " - " -NoNewline
    Write-Host $ServiceStatus
}
