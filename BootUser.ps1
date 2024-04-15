Add-Type -AssemblyName System.Windows.Forms

# Define the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "User Logoff"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"

# Define labels
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(370, 20)
$label.Text = "Enter username:"
$form.Controls.Add($label)

# Define text box for input
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 40)
$textBox.Size = New-Object System.Drawing.Size(370, 20)
$form.Controls.Add($textBox)

# Define button
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(150, 80)
$button.Size = New-Object System.Drawing.Size(100, 30)
$button.Text = "OK"
$button.Add_Click({
    $username = $textBox.Text
    $filteredUser = $sessionData | Where-Object { $_.Username -eq $username }
    
    if ($filteredUser.Count -gt 0) {
        # Display user sessions
        $sessionForm = New-Object System.Windows.Forms.Form
        $sessionForm.Text = "User Sessions"
        $sessionForm.Size = New-Object System.Drawing.Size(400, 300)
        $sessionForm.StartPosition = "CenterScreen"

        $sessionLabel = New-Object System.Windows.Forms.Label
        $sessionLabel.Location = New-Object System.Drawing.Point(10, 20)
        $sessionLabel.Size = New-Object System.Drawing.Size(370, 20)
        $sessionLabel.Text = "Select sessions to log off:"
        $sessionForm.Controls.Add($sessionLabel)

        $sessionListBox = New-Object System.Windows.Forms.CheckedListBox
        $sessionListBox.Location = New-Object System.Drawing.Point(10, 40)
        $sessionListBox.Size = New-Object System.Drawing.Size(370, 150)
        foreach ($user in $filteredUser) {
            $sessionListBox.Items.Add("$($user.HostServer) - $($user.SessionID)")
        }
        $sessionForm.Controls.Add($sessionListBox)

        $sessionButton = New-Object System.Windows.Forms.Button
        $sessionButton.Location = New-Object System.Drawing.Point(150, 200)
        $sessionButton.Size = New-Object System.Drawing.Size(100, 30)
        $sessionButton.Text = "Log off"
        $sessionButton.Add_Click({
            foreach ($index in $sessionListBox.CheckedIndices) {
                $selectedSession = $filteredUser[$index]
                Invoke-RDUserLogoff -HostServer $selectedSession.HostServer -UnifiedSessionID $selectedSession.SessionID
            }
            $form.Close()
            $sessionForm.Close()
        })
        $sessionForm.Controls.Add($sessionButton)

        $sessionForm.ShowDialog() | Out-Null
    } else {
        [System.Windows.Forms.MessageBox]::Show("No sessions found for user: $username", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($button)

# Load data
$brokerArray = @(
    "rdbroker1.shastahealth.org",
    "rdbroker2.shastahealth.org"
)

$sessionData = @() # Initialize an empty array

ForEach ($connectionBroker in $brokerArray) {
    $sessionData += Get-RDUserSession -ConnectionBroker $connectionBroker
}

$sessionData = $sessionData | Sort-Object -Property Username

# Show the form
$result = $form.ShowDialog()
