# Define the module
# Define the initialization function
function Initialize-ServerList {
	# Connect to the Remote Desktop Connection Broker
	[pscredential]$creds = Get-Credential
	$brokerArray = @(
		"rdbroker1.shastahealth.org"
		,"rdbroker2.shastahealth.org"
	)
	$collection = "medical"
	$fullServerArray = [System.Collections.ArrayList]::new()

	# Iterate through the connection brokers and retrieve the session hosts for each
	foreach ($broker in $brokerArray) {
		$collection = "medical"
		$servers = Get-RDSessionHost -CollectionName $collection -ConnectionBroker $broker -Credential [pscredential]$creds
		$strippedServerNames = $servers | Select-Object -ExpandProperty SessionHost -replace "\.shastahealth\.org$" | Sort-Object
		$fullServerArray.AddRange($strippedServerNames)
	}

	# Sort the full server array
	$fullServerArray.Sort()

	# Return the sorted server array
	return $fullServerArray
}

# Export the functions
Export-ModuleMember -Function 'Initialize-ServerList'