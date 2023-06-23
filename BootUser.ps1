$brokerArray = @(
	"rdbroker1.shastahealth.org"
	,"rdbroker2.shastahealth.org"
)

$sessionData = @() # Initialize an empty array

ForEach ($connectionBroker in $brokerArray) {
	$sessionData += Get-RDUserSession -ConnectionBroker $connectionBroker
}

$sessionData = $sessionData | Sort-Object -Property Username

$filteredUser = $sessionData | Where-Object { $_.Username -eq "whughes" }
$filteredUser
if ($filteredUser) {
    Invoke-RDUserLogoff -HostServer $filteredUser.HostServer -UnifiedSessionID $filteredUser.SessionID
}