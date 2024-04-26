Clear-Host
$MedCollection = @(
'Medical01',
'Medical02',
'Medical04',
'Medical06',
'Medical07',
'Medical08',
'Medical09',
'Medical10',
'Medical11',
'Medical12',
'Medical13',
'Medical14'
)

ForEach($MedServer in $MedCollection)
{
    Write-Host $MedServer
    $users = quser /server:$MedServer
    foreach($user in $users)
    {
        $Parsed_user = $user -split '\s+'
        #Get UserName
        #Write-Host $user
        Write-Host $Parsed_user[1]
    }
    Write-Host "`n"
    cmd /c 'pause'
}