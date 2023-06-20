# Connect to the Remote Desktop Connection Broker
$brokerArray = @(
	"rdbroker1.shastahealth.org"
	,"rdbroker2.shastahealth.org"
)
$collectionName = "medical"
$fullServerArray = [System.Collections.ArrayList]::new()

function addServersFromCollectionToArray([string]$connectionBroker) {
	# Retrieve the session hosts from the specified collection on the connection broker
	$servers = Get-RDSessionHost -CollectionName $collectionName -ConnectionBroker $connectionBroker

	# Extract the session host names from the retrieved servers, strip the domain suffix, and sort them
	$strippedServerNames = $servers | Select-Object -ExpandProperty SessionHost -replace "\.shastahealth\.org$" | Sort-Object

	# Add the stripped server names to the global server array
	$fullServerArray.AddRange($strippedServerNames)
}

# Iterate through the connection brokers and retrieve the session hosts for each
foreach ($broker in $brokerArray) {
	addServersFromCollectionToArray $broker
}

# Sort the full server array
$fullServerArray.Sort()