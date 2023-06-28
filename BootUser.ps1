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
    if ($filteredUser) {
        # Display user information
        $infoForm = New-Object System.Windows.Forms.Form
        $infoForm.Text = "User Information"
        $infoForm.Size = New-Object System.Drawing.Size(400, 200)
        $infoForm.StartPosition = "CenterScreen"

        $infoLabel = New-Object System.Windows.Forms.Label
        $infoLabel.Location = New-Object System.Drawing.Point(10, 20)
        $infoLabel.Size = New-Object System.Drawing.Size(370, 20)
        $infoLabel.Text = "User Information:"
        $infoForm.Controls.Add($infoLabel)

        $infoTextBox = New-Object System.Windows.Forms.TextBox
        $infoTextBox.Location = New-Object System.Drawing.Point(10, 40)
        $infoTextBox.Size = New-Object System.Drawing.Size(370, 80)
        $infoTextBox.Multiline = $true
        $infoTextBox.ReadOnly = $true
        $infoTextBox.Text = "Username: $($filteredUser.UserName)`r`nHost Server: $($filteredUser.HostServer)`r`nCollection: $($filteredUser.CollectionName)`r`nUnified Session ID: $($filteredUser.UnifiedSessionID)"
        $infoForm.Controls.Add($infoTextBox)

        $buttonPanel = New-Object System.Windows.Forms.Panel
        $buttonPanel.Location = New-Object System.Drawing.Point(10, 130)
        $buttonPanel.Size = New-Object System.Drawing.Size(370, 40)
        $infoForm.Controls.Add($buttonPanel)

        $logoffButton = New-Object System.Windows.Forms.Button
        $logoffButton.Location = New-Object System.Drawing.Point(0, 0)
        $logoffButton.Size = New-Object System.Drawing.Size(100, 30)
        $logoffButton.Text = "Logoff"
        $logoffButton.Add_Click({
            Invoke-RDUserLogoff -HostServer $filteredUser.HostServer -UnifiedSessionID $filteredUser.SessionID
            $form.Close()
            $infoForm.Close()
        })
        $buttonPanel.Controls.Add($logoffButton)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(120, 0)
        $okButton.Size = New-Object System.Drawing.Size(100, 30)
        $okButton.Text = "OK"
        $okButton.Add_Click({
            $form.Close()
            $infoForm.Close()
        })
        $buttonPanel.Controls.Add($okButton)

        $infoForm.ShowDialog() | Out-Null
    } else {
        [System.Windows.Forms.MessageBox]::Show("No session found for user: $username", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
