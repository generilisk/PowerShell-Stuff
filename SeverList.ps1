# Connect to the Remote Desktop Connection Broker
$brokerArray = @(
    "rdbroker1.shastahealth.org"
    ,"rdbroker2.shastahealth.org"
)
$collectionName = "medical"
$global:fullServerArray = @()

function addServersFromCollectionToArray([string]$connectionBroker) {
    # Retrieve the session hosts from the specified collection on the connection broker
													  
    $servers = Get-RDSessionHost -CollectionName $collectionName -ConnectionBroker $connectionBroker

    # Extract the session host names from the retrieved servers
    $arrayOfServersInCollection = $servers | Select-Object -ExpandProperty SessionHost

    # Strip the domain suffix from the server names and sort them
    $strippedServerNames = $arrayOfServersInCollection -replace "\.shastahealth\.org$" | Sort-Object

    # Add the stripped server names to the global server array
    foreach ($strippedServer in $strippedServerNames){
        $global:fullServerArray += $strippedServer
    }
}

# Iterate through the connection brokers and retrieve the session hosts for each
foreach ($broker in $brokerArray) {
    addServersFromCollectionToArray $broker
}

# Sort the full server array
$fullServerArray = $fullServerArray | Sort-Object