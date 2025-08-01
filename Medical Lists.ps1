# Connect to the Remote Desktop Connection Broker
$brokerArray = @(
    "rdbroker1.shastahealth.org"
    ,"rdbroker2.shastahealth.org"
)
#$connectionBroker = "rdbroker1.shastahealth.org"
$collectionName = "medical"
$fullServerArrray =@()





function addServersFromCollectionToArray([string]$connectionBroker) {
    # Specify the collection name
    # Retrieve the servers in the specified collection
    $servers = Get-RDSessionHost -CollectionName $collectionName -ConnectionBroker $connectionBroker
    # Return the server names
    $arrayOfServersInCollection = $servers | Select-Object -ExpandProperty SessionHost
    Write-Host $arrayOfServersInCollection
}
foreach ($broker in $brokerArray) {
    addServersFromCollectionToArray $broker
}
Write-Host $fullServerArrray